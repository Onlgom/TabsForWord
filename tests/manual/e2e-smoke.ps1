# SMOKE E2E: fast sanity on real Word (~60-90 s, ONE Word session).
# Default check after ANY change that touches painting, geometry or interaction:
#   add-in loads, panel appears and paints, tab click switches, context menu
#   opens/closes, collapse/expand works, Word exits cleanly, no ERROR in log.
# Deep per-feature suites live in the other e2e-*.ps1 (see docs/TESTING.md).
# ASCII only. PowerShell 5.1.
$ErrorActionPreference = 'Stop'
$dir = Join-Path $env:TEMP 'TabsForWord-e2e'
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force $dir | Out-Null }
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

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class W {
    [DllImport("shcore.dll")] public static extern int SetProcessDpiAwareness(int v);
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder sb, int n);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr h, uint f);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint flags, int dx, int dy, uint data, UIntPtr extra);
    [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L; public int T; public int R; public int B; }
    public static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
    public static void Click(int x, int y) {
        SetCursorPos(x, y);
        System.Threading.Thread.Sleep(150);
        mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero);
        System.Threading.Thread.Sleep(80);
        mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero);
    }
    public static void RightClick(int x, int y) {
        SetCursorPos(x, y);
        System.Threading.Thread.Sleep(150);
        mouse_event(0x0008, 0, 0, 0, UIntPtr.Zero);
        System.Threading.Thread.Sleep(80);
        mouse_event(0x0010, 0, 0, 0, UIntPtr.Zero);
    }
    public static void Key(byte vk) {
        keybd_event(vk, 0, 0, UIntPtr.Zero);
        System.Threading.Thread.Sleep(70);
        keybd_event(vk, 0, 2, UIntPtr.Zero);
    }
    public static IntPtr FindMenuOfProcess(uint pid) {
        IntPtr found = IntPtr.Zero;
        EnumWindows(delegate(IntPtr h, IntPtr l) {
            if (!IsWindowVisible(h)) return true;
            if (Cls(h) != "#32768") return true;
            uint p; GetWindowThreadProcessId(h, out p);
            if (p != pid) return true;
            found = h; return false;
        }, IntPtr.Zero);
        return found;
    }
}
"@
[W]::SetProcessDpiAwareness(2) | Out-Null

# Waiting for the process menu to appear or disappear instead of fixed pauses
function Wait-Menu([uint32]$wordPid, [bool]$shouldBeOpen, [int]$timeoutMs = 5000) {
    $deadline = [Environment]::TickCount + $timeoutMs
    while ([Environment]::TickCount -lt $deadline) {
        $isOpen = ([W]::FindMenuOfProcess($wordPid) -ne [IntPtr]::Zero)
        if ($isOpen -eq $shouldBeOpen) { return $true }
        Start-Sleep -Milliseconds 100
    }
    return $false
}

function Get-Host2([IntPtr]$wordTop) {
    $list = New-Object System.Collections.ArrayList
    $cb = [W+EnumProc]{
        param($h, $l)
        $cls = [W]::Cls($h)
        if ($cls.StartsWith('WindowsForms10') -and [W]::IsWindowVisible($h)) {
            $r = New-Object W+RECT
            [void][W]::GetWindowRect($h, [ref]$r)
            [void]$list.Add([pscustomobject]@{
                Hwnd = $h; Parent = [W]::GetParent($h)
                L = $r.L; T = $r.T; R = $r.R; W = ($r.R - $r.L); H = ($r.B - $r.T)
            })
        }
        return $true
    }
    [void][W]::EnumChildWindows($wordTop, $cb, [IntPtr]::Zero)
    return @($list | Where-Object { $_.Parent -eq $wordTop } | Select-Object -First 1)
}

$p = Get-Process WINWORD -ErrorAction SilentlyContinue
if ($p) { Write-Host 'ABORT: WINWORD is running.'; exit 2 }

$installDir = Join-Path $env:LOCALAPPDATA 'TabsForWord'
Copy-Item (Join-Path $repo 'src\TabsForWord\bin\Release\TabsForWord.dll') (Join-Path $installDir 'TabsForWord.dll') -Force

$cfgPath = Join-Path $installDir 'native-host.cfg'
$cfgBackup = $null
if (Test-Path $cfgPath) { $cfgBackup = Get-Content $cfgPath -Raw }
Set-Content -Path $cfgPath -Value "mode=native`r`ndump=0`r`nreserve=1" -Encoding Ascii

$logFile = Join-Path $installDir ('Logs\TabsForWord-' + (Get-Date -Format 'yyyyMMdd') + '.log')
$logStart = 0
if (Test-Path $logFile) { $logStart = (Get-Item $logFile).Length }
function Read-NewLog {
    if (-not (Test-Path $logFile)) { return '' }
    $fs = [System.IO.File]::Open($logFile, 'Open', 'Read', 'ReadWrite')
    try {
        [void]$fs.Seek($script:logStart, 'Begin')
        $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8)
        return $sr.ReadToEnd()
    } finally { $fs.Dispose() }
}

$word = $null
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
    Start-Sleep -Seconds 3
    Step 'add-in started' ((Read-NewLog) -match 'Add-in started') ''

    # a second document, so there is something to switch between (unsaved - faster)
    $null = $word.Documents.Add()
    Start-Sleep -Seconds 2

    $wordPid = [uint32](Get-Process WINWORD).Id
    $wordTop = [W]::GetAncestor([IntPtr]$word.ActiveWindow.Hwnd, 2)
    [void][W]::SetForegroundWindow($wordTop)
    $word.ActiveWindow.WindowState = 1
    Start-Sleep -Seconds 2

    $s = Get-Host2 $wordTop
    Step 'panel present' ($null -ne $s) "rect=$($s.L),$($s.T) $($s.W)x$($s.H)"
    $expandedH = $s.H

    # the strip really is painted: the column contains the bottom line (darker than the background)
    $bmp = New-Object System.Drawing.Bitmap(1, $s.H)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen(($s.L + [int]($s.W * 0.6)), $s.T, 0, 0, $bmp.Size)
    $g.Dispose()
    $minLum = 999
    for ($y = 0; $y -lt $s.H; $y++) {
        $c = $bmp.GetPixel(0, $y)
        $lum = ($c.R + $c.G + $c.B) / 3.0
        if ($lum -lt $minLum) { $minLum = $lum }
    }
    $bmp.Dispose()
    Step 'panel painted (border line found)' ($minLum -lt 235) "minLum=$([int]$minLum)"

    # a click on the body of the first tab switches the document
    $before = $word.ActiveDocument.Name
    [W]::Click($s.L + 30, $s.T + [int]($s.H * 0.55))
    Start-Sleep -Seconds 2
    $after = $word.ActiveDocument.Name
    Step 'tab click switches document' ($after -ne $before) "$before -> $after"

    # SDI: every document has its OWN window - after switching we work with the new
    # active window (it may be non-maximised and somewhere else on screen);
    # we maximise it so that the strip coordinates are predictable.
    $wordTop = [W]::GetAncestor([IntPtr]$word.ActiveWindow.Hwnd, 2)
    [void][W]::SetForegroundWindow($wordTop)
    $word.ActiveWindow.WindowState = 1
    Start-Sleep -Seconds 2

    # the context menu opens and closes with Esc
    $sNow = Get-Host2 $wordTop
    [W]::RightClick($sNow.L + 30, $sNow.T + [int]($sNow.H * 0.55))
    $opened = Wait-Menu $wordPid $true
    Step 'context menu opens' $opened ''
    [W]::Key(0x1B)  # Esc
    $closed = Wait-Menu $wordPid $false 3000
    Step 'context menu closes (Esc)' $closed ''

    # collapse (the button at the right edge) and expand again
    $sNow = Get-Host2 $wordTop
    $expandedH = $sNow.H
    [W]::Click($sNow.R - 25, $sNow.T + [int]($sNow.H * 0.5))
    Start-Sleep -Milliseconds 1500
    $sCol = Get-Host2 $wordTop
    Step 'collapse shrinks panel' ($null -ne $sCol -and $sCol.H -lt ($expandedH * 0.7)) "H=$expandedH -> $($sCol.H)"
    [W]::Click($sCol.R - 25, $sCol.T + [int]($sCol.H * 0.5))
    Start-Sleep -Milliseconds 1500
    $sExp = Get-Host2 $wordTop
    Step 'expand restores panel' ($null -ne $sExp -and [Math]::Abs($sExp.H - $expandedH) -le 2) "H=$($sCol.H) -> $($sExp.H)"

    foreach ($doc in @($word.Documents)) { $doc.Close([ref]0) }
    Start-Sleep -Seconds 2
    $word.Quit()
    Start-Sleep -Seconds 3
} catch {
    Step 'scenario exception' $false $_.Exception.Message
    try { if ($word) { $word.Quit() } } catch {}
} finally {
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch {}
    if ($null -ne $cfgBackup) { Set-Content -Path $cfgPath -Value $cfgBackup -NoNewline }
    elseif (Test-Path $cfgPath) { Remove-Item $cfgPath -Force }   # cfg-fayla ne bylo - ne ostavlyaem svoy
}

Start-Sleep -Seconds 2
$still = Get-Process WINWORD -ErrorAction SilentlyContinue
Step 'WINWORD exited cleanly' (-not $still) ''

$tail = Read-NewLog
$errors = @($tail -split "`n" | Where-Object { $_ -match '\[ERROR\]' })
Step 'no ERROR lines in add-in log' ($errors.Count -eq 0) "errors=$($errors.Count)"

Write-Host ''
Write-Host '=== SMOKE SUMMARY ==='
$results | ForEach-Object { Write-Host $_ }
$failCount = @($results | Where-Object { $_ -like '`[FAIL*' }).Count
Write-Host "FAILURES: $failCount"
exit $failCount
