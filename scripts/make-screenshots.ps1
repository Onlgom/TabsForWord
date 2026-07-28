# make-screenshots.ps1 - produces the README screenshots in docs/images.
#
# Runs a real Word with a few neutral demo documents and captures three shots:
# the tab bar in place, the all-tabs menu, the tab context menu. Real Word
# because that IS the product: a mock-up would show a strip that nothing hosts.
#
# Privacy: the capture starts BELOW the Word title bar on purpose - that row
# carries the signed-in Microsoft account name. The demo documents are created
# by this script, so no real document name reaches an image.
#
# The user settings it has to touch (interface language, tab colours, pinning)
# are backed up and restored, exactly as the E2E scripts do.
#
# Requirements: Word closed, do not touch the mouse or the keyboard while it runs.
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File scripts\make-screenshots.ps1
param([string]$OutDir)
$ErrorActionPreference = 'Stop'

$repo = Split-Path -Parent $PSScriptRoot
if (-not $OutDir) { $OutDir = Join-Path $repo 'docs\images' }
$work = Join-Path $env:TEMP 'TabsForWord-shots'

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class S {
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
    [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr h);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint f, int dx, int dy, uint d, UIntPtr e);
    [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint f, UIntPtr e);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L; public int T; public int R; public int B; }
    public static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
    public static void Click(int x, int y) {
        SetCursorPos(x, y); System.Threading.Thread.Sleep(150);
        mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero); System.Threading.Thread.Sleep(80);
        mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero);
    }
    public static void RightClick(int x, int y) {
        SetCursorPos(x, y); System.Threading.Thread.Sleep(150);
        mouse_event(0x0008, 0, 0, 0, UIntPtr.Zero); System.Threading.Thread.Sleep(80);
        mouse_event(0x0010, 0, 0, 0, UIntPtr.Zero);
    }
    public static void Key(byte vk) {
        keybd_event(vk, 0, 0, UIntPtr.Zero); System.Threading.Thread.Sleep(70);
        keybd_event(vk, 0, 2, UIntPtr.Zero); System.Threading.Thread.Sleep(180);
    }
}
"@
[S]::SetProcessDpiAwareness(2) | Out-Null

function Get-Strip([IntPtr]$top) {
    $found = $null
    $cb = [S+EnumProc]{
        param($h, $l)
        if ([S]::GetParent($h) -ne $top) { return $true }
        if (-not [S]::Cls($h).StartsWith('WindowsForms10')) { return $true }
        $r = New-Object S+RECT
        [void][S]::GetWindowRect($h, [ref]$r)
        $script:found = [pscustomobject]@{ L=$r.L; T=$r.T; R=$r.R; B=$r.B; W=($r.R-$r.L); H=($r.B-$r.T) }
        return $false
    }
    [void][S]::EnumChildWindows($top, $cb, [IntPtr]::Zero)
    return $script:found
}

# Saves a region of the SCREEN, scaled down to a sane width for a README.
function Save-Shot($x, $y, $w, $h, $name, $targetWidth = 1200) {
    $bmp = New-Object System.Drawing.Bitmap($w, $h)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($x, $y, 0, 0, $bmp.Size)
    $g.Dispose()
    $out = $bmp
    if ($targetWidth -gt 0 -and $w -gt $targetWidth) {
        $nh = [int][Math]::Round($h * $targetWidth / $w)
        $scaled = New-Object System.Drawing.Bitmap($targetWidth, $nh)
        $sg = [System.Drawing.Graphics]::FromImage($scaled)
        $sg.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $sg.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $sg.DrawImage($bmp, 0, 0, $targetWidth, $nh)
        $sg.Dispose(); $bmp.Dispose()
        $out = $scaled
    }
    $path = Join-Path $OutDir $name
    $out.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    Write-Host ("  {0}  {1}x{2}" -f $name, $out.Width, $out.Height)
    $out.Dispose()
}

if (Get-Process WINWORD -ErrorAction SilentlyContinue) { Write-Host 'ABORT: WINWORD is running.'; exit 2 }

New-Item -ItemType Directory -Force $OutDir | Out-Null
New-Item -ItemType Directory -Force $work | Out-Null

$installDir = Join-Path $env:LOCALAPPDATA 'TabsForWord'
Copy-Item (Join-Path $repo 'src\TabsForWord\bin\Release\TabsForWord.dll') (Join-Path $installDir 'TabsForWord.dll') -Force

# --- back up the settings this script overwrites ---
$cfgNames = @('language.cfg', 'tab-colors.cfg', 'tab-order.cfg')
$backup = @{}
foreach ($n in $cfgNames) {
    $p = Join-Path $installDir $n
    if (Test-Path $p) { $backup[$n] = Get-Content $p -Raw -Encoding UTF8 } else { $backup[$n] = $null }
}

# --- demo documents (neutral English names; nothing real is opened) ---
$docs = @('Annual report.docx', 'Meeting notes.docx', 'Contract draft.docx', 'Presentation outline.docx')
$paths = @()
$word = $null
try {
    # English interface for the shots: the README is in English.
    Set-Content -Path (Join-Path $installDir 'language.cfg') -Value @('# language', 'lang=en') -Encoding UTF8
    # One coloured tab and one pinned tab, so a single shot shows more than plain
    # tabs. The colour is one of the eight presets from the menu (the green one) -
    # a screenshot should show what the product actually offers.
    Set-Content -Path (Join-Path $installDir 'tab-colors.cfg') `
        -Value @('# colours', ('1E8E3E|' + (Join-Path $work $docs[2]))) -Encoding UTF8
    Set-Content -Path (Join-Path $installDir 'tab-order.cfg') `
        -Value @('# order', ('PIN|' + (Join-Path $work $docs[0]))) -Encoding UTF8

    Start-Process 'winword.exe' -ArgumentList '/w'
    $deadline = (Get-Date).AddSeconds(45)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 2
        try { $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application'); if ($word -and $word.Documents.Count -ge 1) { break } } catch { $word = $null }
    }
    if (-not $word) { throw 'Could not attach to Word' }
    Start-Sleep -Seconds 3

    $texts = @(
        'Annual report',
        'Meeting notes',
        'Contract draft',
        'Presentation outline'
    )
    for ($i = 0; $i -lt $docs.Count; $i++) {
        $p = Join-Path $work $docs[$i]
        Remove-Item $p -Force -ErrorAction SilentlyContinue
        $d = $word.Documents.Add()
        $d.Content.Text = $texts[$i] + "`r`n`r`nThis document exists only to fill the tab bar in the screenshots."
        $d.SaveAs([string]$p)
        $paths += $p
        Start-Sleep -Milliseconds 700
    }
    # Close the empty document Word created with /w, if it is still around.
    foreach ($d in @($word.Documents)) { if (-not $d.Path) { $d.Close([ref]0) } }
    Start-Sleep -Seconds 2

    # The last document gets an unsaved change, so the amber dot is visible.
    $word.ActiveDocument.Content.InsertAfter(' ')
    Start-Sleep -Milliseconds 500

    $top = [S]::GetAncestor([IntPtr]$word.ActiveWindow.Hwnd, 2)
    [void][S]::BringWindowToTop($top); [void][S]::SetForegroundWindow($top)
    $word.ActiveWindow.WindowState = 1
    Start-Sleep -Seconds 3

    $strip = Get-Strip $top
    if (-not $strip) { throw 'the tab strip was not found' }
    Write-Host ("strip: {0},{1} {2}x{3}" -f $strip.L, $strip.T, $strip.W, $strip.H)

    # 1. The tab bar in place. The top margin is small on purpose: it takes in the
    #    row of ribbon group labels sitting right above the strip - enough to place
    #    the strip visually - without cutting a row of ribbon buttons in half.
    #    The title bar, and the signed-in account name in it, stays out of frame.
    $x = $strip.L
    $y = $strip.T - 16
    $w = $strip.W
    $h = 360

    # Park the pointer away from the strip first. A tab tooltip shows the FULL
    # PATH of the document, so a pointer left hovering over a tab puts the build
    # machine's user profile path straight into the picture - which is exactly
    # what happened on the first attempt at these shots.
    [void][S]::SetCursorPos(($strip.L + [int]($strip.W / 2)), ($strip.B + 260))
    Start-Sleep -Seconds 3
    Save-Shot $x $y $w $h 'tab-bar.png'

    # 2. The all-tabs menu: the second button from the right (collapse is last).
    $btn = 26
    [S]::Click(($strip.R - [int]($btn * 1.9)), ($strip.T + [int]($strip.H / 2)))
    Start-Sleep -Seconds 1
    Save-Shot $x $y $w $h 'all-tabs-menu.png'
    [S]::Key(0x1B)   # Esc
    Start-Sleep -Seconds 1

    # 3. The context menu of a tab (right-click on the second tab).
    [S]::RightClick(($strip.L + 260), ($strip.T + [int]($strip.H / 2)))
    Start-Sleep -Seconds 1
    Save-Shot $x $y $w $h 'tab-menu.png'
    [S]::Key(0x1B)
    Start-Sleep -Seconds 1

    foreach ($d in @($word.Documents)) { $d.Close([ref]0) }
    Start-Sleep -Seconds 2
    $word.Quit()
    Start-Sleep -Seconds 3
}
catch {
    Write-Host ('FAILED: ' + $_.Exception.Message) -ForegroundColor Red
    try { if ($word) { $word.Quit() } } catch { }
}
finally {
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch { }
    Start-Sleep -Seconds 2
    # Restore exactly what was there: a setting left behind is a bug (TESTING.md).
    foreach ($n in $cfgNames) {
        $p = Join-Path $installDir $n
        if ($null -ne $backup[$n]) { Set-Content -Path $p -Value $backup[$n] -NoNewline -Encoding UTF8 }
        else { Remove-Item $p -Force -ErrorAction SilentlyContinue }
    }
    foreach ($p in $paths) { Remove-Item $p -Force -ErrorAction SilentlyContinue }
}

Write-Host ''
Write-Host "Screenshots written to $OutDir" -ForegroundColor Green
Write-Host 'Look at every file before committing: nothing personal must be in frame.'
