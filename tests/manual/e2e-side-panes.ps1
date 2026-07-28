# E2E: tab strip must not cover Word side panes (Navigation/Editor docks).
# Opens the Navigation pane programmatically, verifies the tab host shrinks to
# the document area and does not intersect the dock; closing restores width.
# Also captures a zoomed shot of the right-side buttons for visual check.
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
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L; public int T; public int R; public int B; }
    public static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
}
"@
[W]::SetProcessDpiAwareness(2) | Out-Null

function Get-Children([IntPtr]$top) {
    $list = New-Object System.Collections.ArrayList
    $cb = [W+EnumProc]{
        param($h, $l)
        if (-not [W]::IsWindowVisible($h)) { return $true }
        $r = New-Object W+RECT
        [void][W]::GetWindowRect($h, [ref]$r)
        [void]$list.Add([pscustomobject]@{
            Hwnd = $h; Parent = [W]::GetParent($h); Cls = [W]::Cls($h)
            L=$r.L; T=$r.T; R=$r.R; B=$r.B; W=($r.R-$r.L); H=($r.B-$r.T)
        })
        return $true
    }
    [void][W]::EnumChildWindows($top, $cb, [IntPtr]::Zero)
    return @($list | Where-Object { $_.Parent -eq $top })
}

function Get-Host2([IntPtr]$top) {
    return (Get-Children $top | Where-Object { $_.Cls.StartsWith('WindowsForms10') } | Select-Object -First 1)
}

function Intersects($a, $b) {
    return ($a.L -lt $b.R) -and ($b.L -lt $a.R) -and ($a.T -lt $b.B) -and ($b.T -lt $a.B)
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
    Start-Sleep -Seconds 4

    $top = [W]::GetAncestor([IntPtr]$word.ActiveWindow.Hwnd, 2)
    [void][W]::SetForegroundWindow($top)
    $word.ActiveWindow.WindowState = 1
    Start-Sleep -Seconds 2

    $h0 = Get-Host2 $top
    $wwf0 = (Get-Children $top | Where-Object { $_.Cls -eq '_WwF' } | Select-Object -First 1)
    Step 'host present (baseline)' ($null -ne $h0 -and $null -ne $wwf0) "host=$($h0.L),$($h0.T) $($h0.W)x$($h0.H)"
    $baseW = $h0.W

    # snapshot of the right-hand strip buttons (a ~130 px zone) for a visual check
    try {
        $zw = 130; $zx = $h0.R - $zw
        $bmp = New-Object System.Drawing.Bitmap($zw, $h0.H)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.CopyFromScreen($zx, $h0.T, 0, 0, $bmp.Size)
        $big = New-Object System.Drawing.Bitmap($bmp, ($zw * 3), ($h0.H * 3))
        $big.Save((Join-Path $dir 'e2e7-buttons-zoom.png'), [System.Drawing.Imaging.ImageFormat]::Png)
        $g.Dispose(); $bmp.Dispose(); $big.Dispose()
    } catch { }

    # --- Open the navigation pane (on the left) ---
    $word.ActiveWindow.DocumentMap = $true
    Start-Sleep -Seconds 3

    $h1 = Get-Host2 $top
    $wwf1 = (Get-Children $top | Where-Object { $_.Cls -eq '_WwF' } | Select-Object -First 1)
    # a side dock: MsoCommandBarDock taller than 300 (neither the ribbon nor the status bar)
    $dock = (Get-Children $top | Where-Object { $_.Cls -eq 'MsoCommandBarDock' -and $_.H -gt 300 } | Select-Object -First 1)
    Step 'navigation dock appeared' ($null -ne $dock) "dock=$($dock.L),$($dock.T) $($dock.W)x$($dock.H)"
    Step 'host shrank to document area' ($h1.W -lt $baseW -and [Math]::Abs($h1.L - $wwf1.L) -le 2) `
        "host=$($h1.L)..$($h1.R) w=$($h1.W), _WwF=$($wwf1.L)..$($wwf1.R)"
    Step 'host does NOT cover the side pane' (-not (Intersects $h1 $dock)) `
        "hostX=$($h1.L)..$($h1.R), dockX=$($dock.L)..$($dock.R)"

    # --- Close the navigation pane ---
    $word.ActiveWindow.DocumentMap = $false
    Start-Sleep -Seconds 3
    $h2 = Get-Host2 $top
    Step 'closing pane restores full width' ([Math]::Abs($h2.W - $baseW) -le 2) "w=$($h1.W) -> $($h2.W)"

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
$errors = @($tail -split "`n" | Where-Object { $_ -match '\[ERROR\]' })
Step 'no ERROR lines in add-in log' ($errors.Count -eq 0) "errors=$($errors.Count)"

Write-Host ''
Write-Host '=== SUMMARY ==='
$results | ForEach-Object { Write-Host $_ }
$failCount = @($results | Where-Object { $_ -like '`[FAIL*' }).Count
Write-Host "FAILURES: $failCount"
exit $failCount
