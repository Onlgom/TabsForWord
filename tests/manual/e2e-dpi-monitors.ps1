# E2E probe: per-monitor DPI behaviour of the native tab host.
# Auto-enumerates every connected monitor (position + per-monitor DPI), then
# moves the Word window onto each one and measures:
#   - GetDpiForWindow of the Word top-level window (must equal the monitor DPI),
#   - physical height of the host surface,
#   - actual drawn panel height (bottom border line scan) => content scale.
# The panel must scale proportionally to monitor DPI relative to a baseline
# (the lowest-DPI monitor); returning to the baseline must restore the size.
# ASCII only. PowerShell 5.1. Read-only probe: changes nothing permanently.
# NOTE: needs >=2 monitors with DIFFERENT scales to exercise DPI virtualisation;
#       with a single scale it still checks cross-monitor moves and reports so.
$ErrorActionPreference = 'Stop'
$dir = Join-Path $env:TEMP 'TabsForWord-e2e'
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force $dir | Out-Null }
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
using System.Collections.Generic;
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
    [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr h, int x, int y, int w, int hh, bool repaint);
    [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr h);
    [DllImport("shcore.dll")] public static extern int GetDpiForMonitor(IntPtr hmon, int type, out uint dx, out uint dy);
    public delegate bool MonProc(IntPtr hmon, IntPtr hdc, ref RECT r, IntPtr d);
    [DllImport("user32.dll")] public static extern bool EnumDisplayMonitors(IntPtr hdc, IntPtr rect, MonProc cb, IntPtr d);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L; public int T; public int R; public int B; }
    public static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
    public static List<string> Mons = new List<string>();
    static bool MonCb(IntPtr h, IntPtr dc, ref RECT r, IntPtr d) {
        uint dx, dy; GetDpiForMonitor(h, 0, out dx, out dy);
        Mons.Add(string.Format("{0};{1};{2};{3};{4}", r.L, r.T, r.R, r.B, dx));
        return true;
    }
    public static List<string> EnumMonitors() {
        Mons.Clear();
        EnumDisplayMonitors(IntPtr.Zero, IntPtr.Zero, MonCb, IntPtr.Zero);
        return Mons;
    }
}
"@
[W]::SetProcessDpiAwareness(2) | Out-Null   # per-monitor aware: all coords are physical

function Get-HostSurface([IntPtr]$wordTop) {
    $list = New-Object System.Collections.ArrayList
    $cb = [W+EnumProc]{
        param($h, $l)
        $cls = [W]::Cls($h)
        if ($cls.StartsWith('WindowsForms10')) {
            $r = New-Object W+RECT
            [void][W]::GetWindowRect($h, [ref]$r)
            [void]$list.Add([pscustomobject]@{
                Hwnd = $h; Parent = [W]::GetParent($h)
                L = $r.L; T = $r.T; W = ($r.R - $r.L); H = ($r.B - $r.T)
                Visible = [W]::IsWindowVisible($h)
            })
        }
        return $true
    }
    [void][W]::EnumChildWindows($wordTop, $cb, [IntPtr]::Zero)
    return @($list | Where-Object { $_.Parent -eq $wordTop })
}

# Scan a column of pixels inside the host surface for the darkest row
# (bottom border line of the panel on a light background) => drawn height.
function Measure-DrawnPanel($s, $name) {
    $bmp = New-Object System.Drawing.Bitmap($s.W, $s.H)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($s.L, $s.T, 0, 0, $bmp.Size)
    $g.Dispose()
    $bmp.Save((Join-Path $dir "$name.png"), [System.Drawing.Imaging.ImageFormat]::Png)

    $cols = @([int]($s.W - 150), [int]($s.W - 250), [int]($s.W * 0.6))
    $bestY = -1; $bestLum = 999
    foreach ($x in $cols) {
        if ($x -lt 0 -or $x -ge $s.W) { continue }
        for ($y = 3; $y -lt $s.H; $y++) {
            $c = $bmp.GetPixel($x, $y)
            $lum = ($c.R + $c.G + $c.B) / 3.0
            if ($lum -lt $bestLum) { $bestLum = $lum; $bestY = $y }
        }
    }
    $bmp.Dispose()
    return [pscustomobject]@{ DrawnHeight = $bestY + 1; LineLum = [int]$bestLum }
}

function Measure-State([IntPtr]$wordTop, [string]$name) {
    $dpi = [W]::GetDpiForWindow($wordTop)
    $s = (Get-HostSurface $wordTop) | Where-Object { $_.Visible } | Select-Object -First 1
    if (-not $s) { return [pscustomobject]@{ Name=$name; Dpi=$dpi; HostH=-1; DrawnH=-1; HostDpi=0; StripDpi=0 } }
    # The bottom-border pixel scan can catch a not-yet-repainted (blank white)
    # panel right after a move; retry until a dark border line is found.
    $m = $null
    for ($att = 0; $att -lt 4; $att++) {
        $m = Measure-DrawnPanel $s $name
        if ($m.LineLum -le 235 -and $m.DrawnHeight -ge 20) { break }
        Start-Sleep -Milliseconds 1200
        $s2 = (Get-HostSurface $wordTop) | Where-Object { $_.Visible } | Select-Object -First 1
        if ($s2) { $s = $s2 }
    }

    $hostDpi = [W]::GetDpiForWindow($s.Hwnd)
    $stripDpi = 0
    $stripList = New-Object System.Collections.ArrayList
    $cb2 = [W+EnumProc]{ param($h, $l) [void]$stripList.Add($h); return $true }
    [void][W]::EnumChildWindows($s.Hwnd, $cb2, [IntPtr]::Zero)
    if ($stripList.Count -gt 0) { $stripDpi = [W]::GetDpiForWindow($stripList[0]) }

    return [pscustomobject]@{
        Name = $name; Dpi = $dpi
        HostL = $s.L; HostT = $s.T; HostW = $s.W; HostH = $s.H
        DrawnH = $m.DrawnHeight; LineLum = $m.LineLum
        HostDpi = $hostDpi; StripDpi = $stripDpi
    }
}

$p = Get-Process WINWORD -ErrorAction SilentlyContinue
if ($p) { Write-Host 'ABORT: WINWORD is running.'; exit 2 }

# --- Enumerate monitors and derive one window stop fully inside each ---
$stops = New-Object System.Collections.ArrayList
$idx = 0
foreach ($line in [W]::EnumMonitors()) {
    $f = $line -split ';'
    $L = [int]$f[0]; $T = [int]$f[1]; $R = [int]$f[2]; $B = [int]$f[3]; $D = [int]$f[4]
    $w = [Math]::Min(1500, ($R - $L) - 120)
    $h = [Math]::Min(850,  ($B - $T) - 120)
    if ($w -lt 400 -or $h -lt 300) { continue }
    $x = [Math]::Max($L + 60, 0); if ($x + $w -gt $R) { $x = $R - $w - 60 }
    $y = [Math]::Max($T + 60, 0); if ($y + $h -gt $B) { $y = $B - $h - 60 }
    $idx++
    [void]$stops.Add([pscustomobject]@{ Key=("mon$idx-" + [int]($D*100/96) + 'pct'); X=$x; Y=$y; W=$w; H=$h; Dpi=$D })
}
# Baseline = lowest DPI monitor; scaling factors then run >= 1.
$stops = @($stops | Sort-Object Dpi)
$distinctDpi = @($stops | Select-Object -ExpandProperty Dpi | Sort-Object -Unique)
Write-Host ("Monitors: " + (($stops | ForEach-Object { "$($_.Key)@$($_.Dpi)dpi [$($_.X),$($_.Y) $($_.W)x$($_.H)]" }) -join '  '))
if ($stops.Count -lt 1) { Write-Host 'ABORT: no usable monitor.'; exit 3 }

$repo = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$installDir = Join-Path $env:LOCALAPPDATA 'TabsForWord'
Copy-Item (Join-Path $repo 'src\TabsForWord\bin\Release\TabsForWord.dll') (Join-Path $installDir 'TabsForWord.dll') -Force
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

$alphaFile = Join-Path $dir 'dpi-alpha.docx'
$bravoFile = Join-Path $dir 'dpi-bravo.docx'
Remove-Item $alphaFile, $bravoFile -Force -ErrorAction SilentlyContinue

$word = $null
$measured = @{}
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

    $d1 = $word.Documents.Add(); $d1.Content.Text = 'Alpha'
    $d1.SaveAs([string]$alphaFile)
    Start-Sleep -Milliseconds 700
    $d2 = $word.Documents.Add(); $d2.Content.Text = 'Bravo'
    $d2.SaveAs([string]$bravoFile)
    Start-Sleep -Seconds 2

    $wordTop = [W]::GetAncestor([IntPtr]$word.ActiveWindow.Hwnd, 2)
    $word.ActiveWindow.WindowState = 0   # wdWindowStateNormal - so it can be moved
    Start-Sleep -Seconds 1
    [void][W]::SetForegroundWindow($wordTop)

    foreach ($st in $stops) {
        [void][W]::MoveWindow($wordTop, $st.X, $st.Y, $st.W, $st.H, $true)
        Start-Sleep -Seconds 5
        $m = Measure-State $wordTop ("dpi-" + $st.Key)
        $measured[$st.Key] = $m
        Write-Host ("{0}: windowDPI={1} (expect {2}), host H={3}, drawn H={4} (line lum {5}), hostWndDPI={6}, stripWndDPI={7}" -f `
            $st.Key, $m.Dpi, $st.Dpi, $m.HostH, $m.DrawnH, $m.LineLum, $m.HostDpi, $m.StripDpi)
    }

    # Return to the baseline (lowest-DPI) monitor and confirm size restores.
    $base = $stops[0]
    [void][W]::MoveWindow($wordTop, $base.X, $base.Y, $base.W, $base.H, $true)
    Start-Sleep -Seconds 5
    $back = Measure-State $wordTop 'dpi-return'
    Write-Host ("return->$($base.Key): windowDPI={0}, host H={1}, drawn H={2}" -f $back.Dpi, $back.HostH, $back.DrawnH)

    # --- Assessment ---
    $baseM = $measured[$base.Key]
    $baseH = $baseM.HostH
    $baseD = $baseM.DrawnH
    foreach ($st in $stops) {
        $m = $measured[$st.Key]
        Step "$($st.Key): Word window reports monitor DPI $($st.Dpi)" ($m.Dpi -eq $st.Dpi) "got $($m.Dpi)"
        $expH = [int][Math]::Round($baseH * $st.Dpi / $base.Dpi)
        $expD = [int][Math]::Round($baseD * $st.Dpi / $base.Dpi)
        Step "$($st.Key): host height ~$expH px (base $baseH x $($st.Dpi)/$($base.Dpi))" ([Math]::Abs($m.HostH - $expH) -le 2) "got $($m.HostH)"
        Step "$($st.Key): drawn content ~$expD px (base $baseD x $($st.Dpi)/$($base.Dpi))" ([Math]::Abs($m.DrawnH - $expD) -le 3) "got $($m.DrawnH)"
    }
    Step 'returning to baseline restores size' ([Math]::Abs($back.HostH - $baseH) -le 2 -and [Math]::Abs($back.DrawnH - $baseD) -le 3) `
        "host $baseH->$($back.HostH), drawn $baseD->$($back.DrawnH)"
    if ($distinctDpi.Count -lt 2) {
        Write-Host 'NOTE: all monitors share one scale - DPI virtualisation was NOT exercised (cross-monitor moves only).'
    } else {
        Write-Host ("NOTE: exercised distinct scales: " + (($distinctDpi | ForEach-Object { ('{0}%' -f [int]($_ * 100 / 96)) }) -join ', '))
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
$reposLines = @($tail -split "`n" | Where-Object { $_ -match 'repositioned|Reserve: anchor moved|DPI' })
Write-Host ''
Write-Host '=== Host reposition / DPI log lines (last 20) ==='
$reposLines | Select-Object -Last 20 | ForEach-Object { Write-Host $_ }
$errors = @($tail -split "`n" | Where-Object { $_ -match '\[ERROR\]' })
Step 'no ERROR lines in add-in log' ($errors.Count -eq 0) "errors=$($errors.Count)"
if ($errors.Count -gt 0) { $errors | Select-Object -First 8 | ForEach-Object { Write-Host "  $_" } }

Write-Host ''
Write-Host '=== SUMMARY ==='
$results | ForEach-Object { Write-Host $_ }
$failCount = @($results | Where-Object { $_ -like '`[FAIL*' }).Count
Write-Host "FAILURES: $failCount"
Write-Host "Screenshots in: $dir"
exit $failCount
