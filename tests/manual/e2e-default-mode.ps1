# E2E: mode resolution after v1.6.0 (ADR-015).
# Phase A: NO native-host.cfg at all -> add-in must run in the NATIVE host mode
#          (panel is a child of the top-level Word window, no CTP title bar).
# Phase B: mode=ctp -> add-in must fall back to the CLASSIC Custom Task Pane
#          (no WinForms child directly under the Word top-level window).
# The user's own cfg is backed up and restored, temp docs only, Word closed at end.
# ASCII only. PowerShell 5.1. Run with Word CLOSED.
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$results = New-Object System.Collections.ArrayList

function Step($name, $ok, $detail) {
    $tag = 'FAIL'
    if ($ok) { $tag = 'OK  ' }
    $line = "[$tag] $name"
    if ($detail) { $line = "$line -- $detail" }
    [void]$results.Add($line)
    Write-Host $line
}

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class WD {
    [DllImport("shcore.dll")] public static extern int SetProcessDpiAwareness(int v);
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder sb, int n);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr h, uint f);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L; public int T; public int R; public int B; }
    public static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
}
"@
[WD]::SetProcessDpiAwareness(2) | Out-Null

# WinForms child hosted DIRECTLY by the top-level Word window == native host.
# In CTP mode the same control lives inside Word's task pane host, so its parent
# is a pane window, not OpusApp -> this returns nothing.
function Get-NativePanel([IntPtr]$wordTop) {
    $list = New-Object System.Collections.ArrayList
    $cb = [WD+EnumProc]{
        param($h, $l)
        $cls = [WD]::Cls($h)
        if ($cls.StartsWith('WindowsForms10') -and [WD]::IsWindowVisible($h)) {
            $r = New-Object WD+RECT
            [void][WD]::GetWindowRect($h, [ref]$r)
            [void]$list.Add([pscustomobject]@{
                Hwnd = $h; Parent = [WD]::GetParent($h)
                L = $r.L; T = $r.T; W = ($r.R - $r.L); H = ($r.B - $r.T)
            })
        }
        return $true
    }
    [void][WD]::EnumChildWindows($wordTop, $cb, [IntPtr]::Zero)
    return @($list | Where-Object { $_.Parent -eq $wordTop } | Select-Object -First 1)
}

$p = Get-Process WINWORD -ErrorAction SilentlyContinue
if ($p) { Write-Host 'ABORT: WINWORD is running.'; exit 2 }

$installDir = Join-Path $env:LOCALAPPDATA 'TabsForWord'
Copy-Item (Join-Path $repo 'src\TabsForWord\bin\Release\TabsForWord.dll') (Join-Path $installDir 'TabsForWord.dll') -Force

$cfgPath = Join-Path $installDir 'native-host.cfg'
$cfgBackup = $null
$cfgExisted = Test-Path $cfgPath
if ($cfgExisted) { $cfgBackup = Get-Content $cfgPath -Raw }

$logFile = Join-Path $installDir ('Logs\TabsForWord-' + (Get-Date -Format 'yyyyMMdd') + '.log')
$script:logStart = 0
function Mark-Log {
    $script:logStart = 0
    if (Test-Path $logFile) { $script:logStart = (Get-Item $logFile).Length }
}
function Read-NewLog {
    if (-not (Test-Path $logFile)) { return '' }
    $fs = [System.IO.File]::Open($logFile, 'Open', 'Read', 'ReadWrite')
    try {
        [void]$fs.Seek($script:logStart, 'Begin')
        $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8)
        return $sr.ReadToEnd()
    } finally { $fs.Dispose() }
}

# One Word session: start, look at the panel, quit. Returns log tail + panel info.
function Invoke-WordSession {
    Mark-Log
    $word = $null
    $panel = $null
    try {
        Start-Process 'winword.exe' -ArgumentList '/w'
        $deadline = (Get-Date).AddSeconds(45)
        while ((Get-Date) -lt $deadline) {
            Start-Sleep -Seconds 2
            try {
                $word = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
                if ($word -and $word.Documents.Count -ge 1) { break }
            } catch { $word = $null }
        }
        if (-not $word) { throw 'Could not attach to Word' }
        Start-Sleep -Seconds 4
        $wordTop = [WD]::GetAncestor([IntPtr]$word.ActiveWindow.Hwnd, 2)
        [void][WD]::SetForegroundWindow($wordTop)
        $word.ActiveWindow.WindowState = 1
        Start-Sleep -Seconds 2
        $panel = Get-NativePanel $wordTop
        foreach ($doc in @($word.Documents)) { $doc.Close([ref]0) }
        Start-Sleep -Seconds 2
        $word.Quit()
        Start-Sleep -Seconds 3
    } finally {
        try { if ($word) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } } catch {}
    }
    Start-Sleep -Seconds 2
    $still = Get-Process WINWORD -ErrorAction SilentlyContinue
    if ($still) { try { $still | Wait-Process -Timeout 10 } catch {} }
    return [pscustomobject]@{ Log = (Read-NewLog); Panel = $panel }
}

try {
    # ---------- Phase A: no config at all ----------
    if (Test-Path $cfgPath) { Remove-Item $cfgPath -Force }
    $a = Invoke-WordSession
    Step 'A: add-in started (no cfg)' ($a.Log -match 'Add-in started') ''
    Step 'A: mode resolved to Native' ($a.Log -match 'Tab host mode: Native') ''
    Step 'A: reserve enabled by default' ($a.Log -match 'Tab host mode: Native \(dump=0, reserve=1\)') ''
    Step 'A: panel hosted inside Word window' ($null -ne $a.Panel) "rect=$($a.Panel.L),$($a.Panel.T) $($a.Panel.W)x$($a.Panel.H)"
    Step 'A: panel not at the very top of the window' ($null -ne $a.Panel -and $a.Panel.T -gt 0) "top=$($a.Panel.T)"
    $errA = @($a.Log -split "`n" | Where-Object { $_ -match '\[ERROR\]' })
    Step 'A: no ERROR lines' ($errA.Count -eq 0) "errors=$($errA.Count)"

    # ---------- Phase B: classic mode escape hatch ----------
    Set-Content -Path $cfgPath -Value 'mode=ctp' -Encoding Ascii
    $b = Invoke-WordSession
    Step 'B: add-in started (mode=ctp)' ($b.Log -match 'Add-in started') ''
    Step 'B: mode resolved to CustomTaskPane' ($b.Log -match 'Tab host mode: CustomTaskPane') ''
    Step 'B: no native panel under the Word window' ($null -eq $b.Panel) ''
    $errB = @($b.Log -split "`n" | Where-Object { $_ -match '\[ERROR\]' })
    Step 'B: no ERROR lines' ($errB.Count -eq 0) "errors=$($errB.Count)"
} catch {
    Step 'scenario exception' $false $_.Exception.Message
    try { Get-Process WINWORD -ErrorAction SilentlyContinue | Out-Null } catch {}
} finally {
    # User's own config comes back exactly as it was (or disappears again).
    if ($cfgExisted) { Set-Content -Path $cfgPath -Value $cfgBackup -NoNewline }
    elseif (Test-Path $cfgPath) { Remove-Item $cfgPath -Force }
}

Write-Host ''
Write-Host '=== DEFAULT MODE SUMMARY ==='
$results | ForEach-Object { Write-Host $_ }
$failCount = @($results | Where-Object { $_ -like '`[FAIL*' }).Count
Write-Host "FAILURES: $failCount"
exit $failCount
