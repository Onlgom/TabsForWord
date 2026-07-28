# E2E: the File menu (Backstage) is a normal state, not a locator failure.
# Opens the full-page Word UI, checks the strip hides while it is up and comes
# back after it closes, and that the add-in wrote no window-tree dump, no
# "no content zone" storm and did NOT fall back to the CTP mode (stage 34).
# ASCII only. PowerShell 5.1.
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
public static class W {
    [DllImport("shcore.dll")] public static extern int SetProcessDpiAwareness(int v);
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder sb, int n);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr h, uint f);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr h);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint f, int dx, int dy, uint d, IntPtr e);
    [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, IntPtr extra);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L; public int T; public int R; public int B; }
    public static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
}
"@
[W]::SetProcessDpiAwareness(2) | Out-Null

function Get-AllChildren([IntPtr]$top) {
    $list = New-Object System.Collections.ArrayList
    $cb = [W+EnumProc]{
        param($h, $l)
        $r = New-Object W+RECT
        [void][W]::GetWindowRect($h, [ref]$r)
        [void]$list.Add([pscustomobject]@{
            Hwnd = $h; Parent = [W]::GetParent($h); Cls = [W]::Cls($h); Vis = [W]::IsWindowVisible($h)
            L=$r.L; T=$r.T; R=$r.R; B=$r.B; W=($r.R-$r.L); H=($r.B-$r.T)
        })
        return $true
    }
    [void][W]::EnumChildWindows($top, $cb, [IntPtr]::Zero)
    return $list
}

function Get-Strip([IntPtr]$top) {
    return (Get-AllChildren $top | Where-Object { $_.Parent -eq $top -and $_.Cls.StartsWith('WindowsForms10') } | Select-Object -First 1)
}

function Get-FullPage([IntPtr]$top) {
    return (Get-AllChildren $top | Where-Object { $_.Cls -eq 'FullpageUIHost' -and $_.Vis -and $_.W -gt 400 } | Select-Object -First 1)
}

function Send-Escape() {
    [W]::keybd_event(0x1B, 0, 0, [IntPtr]::Zero)
    [W]::keybd_event(0x1B, 0, 2, [IntPtr]::Zero)
}

# Ctrl+O: in Word 365 it opens the Backstage "Open" page. Language-independent,
# unlike the Alt+F key tip (which is Alt+F only on an English ribbon).
function Send-CtrlO() {
    [W]::keybd_event(0x11, 0, 0, [IntPtr]::Zero)
    [W]::keybd_event(0x4F, 0, 0, [IntPtr]::Zero)
    [W]::keybd_event(0x4F, 0, 2, [IntPtr]::Zero)
    [W]::keybd_event(0x11, 0, 2, [IntPtr]::Zero)
}

function Click-At($x, $y) {
    [void][W]::SetCursorPos($x, $y)
    Start-Sleep -Milliseconds 150
    [W]::mouse_event(0x02, 0, 0, 0, [IntPtr]::Zero)
    [W]::mouse_event(0x04, 0, 0, 0, [IntPtr]::Zero)
}

if (Get-Process WINWORD -ErrorAction SilentlyContinue) { Write-Host 'ABORT: WINWORD is running.'; exit 2 }

$installDir = Join-Path $env:LOCALAPPDATA 'TabsForWord'
Copy-Item (Join-Path $repo 'src\TabsForWord\bin\Release\TabsForWord.dll') (Join-Path $installDir 'TabsForWord.dll') -Force

$logsDir = Join-Path $installDir 'Logs'
$dumpsBefore = @(Get-ChildItem $logsDir -Filter 'WindowTree-*.txt' -ErrorAction SilentlyContinue).Count
$logFile = Join-Path $logsDir ('TabsForWord-' + (Get-Date -Format 'yyyyMMdd') + '.log')
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
$opened = $false
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

    $top = [W]::GetAncestor([IntPtr]$word.ActiveWindow.Hwnd, 2)
    [void][W]::BringWindowToTop($top)
    [void][W]::SetForegroundWindow($top)
    $word.ActiveWindow.WindowState = 1
    Start-Sleep -Seconds 3

    # Exactly ONE Word window is open on purpose: that is the case in which the
    # old code counted every poll inside the File menu as a native-mode failure
    # and after five of them switched the add-in to the CTP fallback.
    Step 'exactly one Word window (the single-window case)' ($word.Windows.Count -eq 1) "windows=$($word.Windows.Count)"

    $s0 = Get-Strip $top
    Step 'strip visible before the File menu' ($null -ne $s0 -and $s0.Vis -and $s0.H -gt 10) `
        "strip=$($s0.L),$($s0.T) $($s0.W)x$($s0.H) vis=$($s0.Vis)"

    # --- Open the full-page UI ---
    Send-CtrlO
    Start-Sleep -Seconds 3
    $fp = Get-FullPage $top
    if (-not $fp) {
        # Fallback: click the File tab (the leftmost tab of the ribbon dock).
        Send-Escape; Start-Sleep -Seconds 1
        $dock = (Get-AllChildren $top | Where-Object { $_.Cls -eq 'MsoCommandBarDock' -and $_.Vis -and $_.H -gt 60 -and $_.W -gt 400 } | Select-Object -First 1)
        if ($dock) {
            Click-At ($dock.L + [int]($dock.W * 0.02) + 20) ($dock.T + [int]($dock.H * 0.28))
            Start-Sleep -Seconds 3
            $fp = Get-FullPage $top
        }
    }
    $opened = $null -ne $fp
    Step 'full-page Word UI opened' $opened $(if ($fp) { "FullpageUIHost=$($fp.L),$($fp.T) $($fp.W)x$($fp.H)" } else { 'BLOCKED: could not open the File menu' })

    if ($opened) {
        # Long enough for more than five ~1 Hz polls: the old code would have
        # dropped into the CTP fallback somewhere in here.
        Start-Sleep -Seconds 10
        $s1 = Get-Strip $top
        Step 'strip hidden while the File menu is up' ($null -ne $s1 -and -not $s1.Vis) "vis=$($s1.Vis)"

        # --- Close it ---
        Send-Escape
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $back = $null
        while ($sw.Elapsed.TotalSeconds -lt 20) {
            Start-Sleep -Milliseconds 300
            $cand = Get-Strip $top
            if ($cand -and $cand.Vis -and $cand.H -gt 10) { $back = $cand; break }
        }
        $sw.Stop()
        Step 'strip comes back after the File menu closes' ($null -ne $back) `
            "after $([Math]::Round($sw.Elapsed.TotalSeconds,1)) s"
        if ($back) {
            $wwf = (Get-AllChildren $top | Where-Object { $_.Cls -eq '_WwF' -and $_.Vis } | Select-Object -First 1)
            Step 'strip is back at the top of the document area' `
                ($null -ne $wwf -and [Math]::Abs($back.L - $wwf.L) -le 2 -and $back.T -lt $wwf.T -and ($wwf.T - $back.T) -le ($back.H + 4)) `
                "strip=$($back.L),$($back.T) $($back.W)x$($back.H), _WwF top=$($wwf.T)"
        }
    }

    foreach ($doc in @($word.Documents)) { $doc.Close([ref]0) }
    Start-Sleep -Seconds 2
    $word.Quit()
    Start-Sleep -Seconds 3
} catch {
    Step 'scenario exception' $false $_.Exception.Message
    try { if ($word) { $word.Quit() } } catch {}
} finally {
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch {}
}

Start-Sleep -Seconds 2
$still = Get-Process WINWORD -ErrorAction SilentlyContinue
Step 'WINWORD exited cleanly' (-not $still) ''

$tail = Read-NewLog
$lines = @($tail -split "`n")
$errors = @($lines | Where-Object { $_ -match '\[ERROR\]' })
Step 'no ERROR lines in add-in log' ($errors.Count -eq 0) "errors=$($errors.Count)"
Step 'no fallback to the CTP mode' (-not ($tail -match 'falling back to CustomTaskPane')) ''

if ($opened) {
    $fullPageLines = @($lines | Where-Object { $_ -match 'Full-page Word UI is open' })
    Step 'the log explains the pause (one INFO line per episode)' ($fullPageLines.Count -ge 1 -and $fullPageLines.Count -le 3) `
        "lines=$($fullPageLines.Count)"
    # Two are normal: the window being built at startup and the window being closed.
    $noZone = @($lines | Where-Object { $_ -match 'Locator: no content zone found' })
    Step 'the File menu produced no WARN storm' ($noZone.Count -le 2) "no-content-zone WARNs=$($noZone.Count)"
}

$dumpsAfter = @(Get-ChildItem $logsDir -Filter 'WindowTree-*.txt' -ErrorAction SilentlyContinue).Count
Step 'no window-tree dump written' ($dumpsAfter -eq $dumpsBefore) "dumps $dumpsBefore -> $dumpsAfter"

Write-Host ''
Write-Host '=== SUMMARY ==='
$results | ForEach-Object { Write-Host $_ }
$failCount = @($results | Where-Object { $_ -like '`[FAIL*' }).Count
Write-Host "FAILURES: $failCount"
exit $failCount
