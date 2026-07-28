# E2E: overflow scrolling with many tabs (24 docs => ‹ › arrows).
# Repro/regression for: active LAST tab + scroll-left clicks caused tabs to
# spill right of the viewport over the arrows. Also checks horizontal
# touchpad scroll (WM_MOUSEHWHEEL to the strip window).
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
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder sb, int n);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr h, uint f);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint flags, int dx, int dy, uint data, UIntPtr extra);
    [DllImport("user32.dll")] public static extern IntPtr SendMessageW(IntPtr h, uint msg, IntPtr w, IntPtr l);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L; public int T; public int R; public int B; }
    public static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
    public static void Click(int x, int y) {
        SetCursorPos(x, y);
        System.Threading.Thread.Sleep(120);
        mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero);
        System.Threading.Thread.Sleep(60);
        mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero);
    }
    // Touchpad horizontal scrolling: WM_MOUSEHWHEEL (0x020E), delta in HIWORD(wParam)
    public static void HWheel(IntPtr hwnd, int delta, int x, int y) {
        IntPtr wParam = new IntPtr(((delta & 0xFFFF) << 16));
        IntPtr lParam = new IntPtr(((y & 0xFFFF) << 16) | (x & 0xFFFF));
        SendMessageW(hwnd, 0x020E, wParam, lParam);
    }
}
"@
[W]::SetProcessDpiAwareness(2) | Out-Null

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
    $hostS = @($list | Where-Object { $_.Parent -eq $wordTop }) | Select-Object -First 1
    $strip = @($list | Where-Object { $_.Parent -eq $hostS.Hwnd }) | Select-Object -First 1
    return @($hostS, $strip)
}

function Snap-Strip($s, $name) {
    $bmp = New-Object System.Drawing.Bitmap($s.W, $s.H)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($s.L, $s.T, 0, 0, $bmp.Size)
    $g.Dispose()
    $bmp.Save((Join-Path $dir "$name.png"), [System.Drawing.Imaging.ImageFormat]::Png)
    return (Join-Path $dir "$name.png")
}

# Artefact: "tab" pixels in the right-hand button zone (to the right of the right
# arrow there is only the strip background, the separator and grey buttons; blue
# frames or text there are a bug). The zone is the last N pixels (the buttons), and
# it stops short of the scroll zone, where the blue active caption is legitimate.
function Count-BluePixels($pngPath, $zoneWidthPx) {
    $bmp = [System.Drawing.Bitmap]::FromFile($pngPath)
    $count = 0
    $x0 = [Math]::Max(0, $bmp.Width - $zoneWidthPx)
    for ($x = $x0; $x -lt $bmp.Width; $x += 2) {
        for ($y = 2; $y -lt $bmp.Height - 2; $y += 2) {
            $c = $bmp.GetPixel($x, $y)
            if ($c.B -gt ([int]$c.R + 40) -and $c.B -gt 140) { $count++ }
        }
    }
    $bmp.Dispose()
    return $count
}

$p = Get-Process WINWORD -ErrorAction SilentlyContinue
if ($p) { Write-Host 'ABORT: WINWORD is running.'; exit 2 }

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

    # 23 extra documents => guaranteed overflow with scrolling
    for ($i = 0; $i -lt 23; $i++) { $null = $word.Documents.Add() }
    Start-Sleep -Seconds 3

    $wordTop = [W]::GetAncestor([IntPtr]$word.ActiveWindow.Hwnd, 2)
    [void][W]::SetForegroundWindow($wordTop)
    $word.ActiveWindow.WindowState = 1
    Start-Sleep -Seconds 2

    $pair = Get-Host2 $wordTop
    $s = $pair[0]; $strip = $pair[1]
    Step 'host + strip present' ($null -ne $s -and $null -ne $strip) "host=$($s.W)x$($s.H)"

    # the active tab is the last one (the last Add is the active document);
    # auto-scroll is already at the right edge. We scroll left by clicking the arrow.
    $yMid = $s.T + [int]($s.H * 0.5)
    $arrowLeftX = $s.L + 26        # PadLeft(14)+btn/2(13) at 100%, with margin
    $png0 = Snap-Strip $s 'e2e8-before-scroll'
    for ($i = 0; $i -lt 6; $i++) {
        [W]::Click($arrowLeftX, $yMid)
        Start-Sleep -Milliseconds 350
    }
    Start-Sleep -Milliseconds 600
    $png1 = Snap-Strip $s 'e2e8-after-left-clicks'
    $blue1 = Count-BluePixels $png1 105
    Step 'no tab artifacts over right buttons after scroll-left' ($blue1 -le 2) "bluePx=$blue1 (zone: right 105px)"

    # --- horizontal scrolling with a "touchpad" (WM_MOUSEHWHEEL into the strip window) ---
    # the cursor is moved BELOW the strip: hover and tooltips must not affect the shots.
    # Determinism through end stops: scroll to the right stop -> reference; scroll left;
    # scroll to the stop again -> the image must match the reference.
    [void][W]::SetCursorPos(($s.L + [int]($s.W * 0.5)), ($s.T + $s.H + 120))
    Start-Sleep -Milliseconds 500
    $cx = [int]($s.L + $s.W * 0.5)
    for ($i = 0; $i -lt 25; $i++) { [W]::HWheel($strip.Hwnd, 120, $cx, $yMid) }   # to the right stop
    Start-Sleep -Milliseconds 800
    $pngH0 = Snap-Strip $s 'e2e8-hwheel-at-max'
    for ($i = 0; $i -lt 5; $i++) { [W]::HWheel($strip.Hwnd, -120, $cx, $yMid) ; Start-Sleep -Milliseconds 60 }
    Start-Sleep -Milliseconds 600
    $pngH1 = Snap-Strip $s 'e2e8-hwheel-left'
    for ($i = 0; $i -lt 25; $i++) { [W]::HWheel($strip.Hwnd, 120, $cx, $yMid) }   # to the stop again
    Start-Sleep -Milliseconds 800
    $pngH2 = Snap-Strip $s 'e2e8-hwheel-back-to-max'

    function Get-Hash($path) { (Get-FileHash $path -Algorithm MD5).Hash }
    Step 'horizontal touchpad wheel scrolls tabs' ((Get-Hash $pngH0) -ne (Get-Hash $pngH1)) ''
    Step 'hwheel to the stop restores the picture' ((Get-Hash $pngH0) -eq (Get-Hash $pngH2)) ''
    $blueH = Count-BluePixels $pngH1 105
    Step 'no artifacts after hwheel scroll' ($blueH -le 2) "bluePx=$blueH"

    foreach ($doc in @($word.Documents)) { $doc.Close([ref]0) }
    Start-Sleep -Seconds 3
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
$errors = @($tail -split "`n" | Where-Object { $_ -match '\[ERROR\]' })
Step 'no ERROR lines in add-in log' ($errors.Count -eq 0) "errors=$($errors.Count)"

Write-Host ''
Write-Host '=== SUMMARY ==='
$results | ForEach-Object { Write-Host $_ }
$failCount = @($results | Where-Object { $_ -like '`[FAIL*' }).Count
Write-Host "FAILURES: $failCount"
exit $failCount
