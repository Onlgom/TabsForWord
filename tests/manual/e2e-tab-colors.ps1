# E2E: tab colors via right-click context menu (set preset, reset, persistence).
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
Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class W {
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
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
    [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr FindWindowW(string cls, string title);
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
        System.Threading.Thread.Sleep(180);
    }
    public static bool MenuOpen() { return FindWindowW("#32768", null) != IntPtr.Zero; }
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
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
[W]::SetProcessDPIAware() | Out-Null

# Polling the Word process menu instead of fixed pauses (fewer timing flakes)
function Wait-Menu([uint32]$wordPid, [bool]$shouldBeOpen, [int]$timeoutMs = 5000) {
    $deadline = [Environment]::TickCount + $timeoutMs
    while ([Environment]::TickCount -lt $deadline) {
        $isOpen = ([W]::FindMenuOfProcess($wordPid) -ne [IntPtr]::Zero)
        if ($isOpen -eq $shouldBeOpen) { return $true }
        Start-Sleep -Milliseconds 100
    }
    return $false
}

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

function Get-ScreenPixel($x, $y) {
    $bmp = New-Object System.Drawing.Bitmap(1, 1)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($x, $y, 0, 0, $bmp.Size)
    $c = $bmp.GetPixel(0, 0)
    $g.Dispose(); $bmp.Dispose()
    return $c
}

function Snap($name) {
    try {
        $vs = [System.Windows.Forms.SystemInformation]::VirtualScreen
        $bmp = New-Object System.Drawing.Bitmap($vs.Width, $vs.Height)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.CopyFromScreen($vs.X, $vs.Y, 0, 0, $bmp.Size)
        $bmp.Save((Join-Path $dir "$name.png"), [System.Drawing.Imaging.ImageFormat]::Png)
        $g.Dispose(); $bmp.Dispose()
        Write-Host "Screenshot: $name.png"
    } catch { Write-Host "Screenshot failed: $($_.Exception.Message)" }
}

$p = Get-Process WINWORD -ErrorAction SilentlyContinue
if ($p) { Write-Host 'ABORT: WINWORD is running.'; exit 2 }

$installDir = Join-Path $env:LOCALAPPDATA 'TabsForWord'
Copy-Item (Join-Path $repo 'src\TabsForWord\bin\Release\TabsForWord.dll') (Join-Path $installDir 'TabsForWord.dll') -Force

# keep the user's native-host.cfg intact: back up, set test mode, restore at the end
$cfgPath = Join-Path $installDir 'native-host.cfg'
$cfgBackup = $null
if (Test-Path $cfgPath) { $cfgBackup = Get-Content $cfgPath -Raw }
Set-Content -Path $cfgPath -Value "mode=native`r`ndump=0`r`nreserve=1" -Encoding Ascii

# clean color store so the test starts from a known state (back up as well)
$colorsPath = Join-Path $installDir 'tab-colors.cfg'
$colorsBackup = $null
if (Test-Path $colorsPath) { $colorsBackup = Get-Content $colorsPath -Raw; Remove-Item $colorsPath -Force }

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

$alphaFile = Join-Path $dir 'v4-color-alpha.docx'
$bravoFile = Join-Path $dir 'v4-color-bravo.docx'
Remove-Item $alphaFile, $bravoFile -Force -ErrorAction SilentlyContinue

$word = $null
try {
    # ---------- session 1: set color, check paint + store, reset, set again ----------
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
    Step 'add-in started (native mode)' ((Read-NewLog) -match 'Native host mode enabled') ''

    # close the blank startup document: the first tab must be the SAVED alpha
    # (persistence is keyed by file path; unsaved docs keep color in memory only)
    foreach ($doc in @($word.Documents)) {
        if (-not $doc.Path) { $doc.Close([ref]0) }
    }
    Start-Sleep -Seconds 1

    $d1 = $word.Documents.Add(); $d1.Content.Text = 'Alpha'
    $d1.SaveAs([string]$alphaFile)
    Start-Sleep -Milliseconds 900
    $d2 = $word.Documents.Add(); $d2.Content.Text = 'Bravo'
    $d2.SaveAs([string]$bravoFile)
    Start-Sleep -Seconds 2

    $activeTop = [W]::GetAncestor([IntPtr]$word.ActiveWindow.Hwnd, 2)
    [void][W]::SetForegroundWindow($activeTop)
    $word.ActiveWindow.WindowState = 1
    Start-Sleep -Seconds 2

    $s = (Get-HostSurface $activeTop)[0]
    Step 'host present' ($null -ne $s -and $s.Visible) "rect=$($s.L),$($s.T) $($s.W)x$($s.H)"

    # first tab (alpha, inactive): click at body center; sample fill above the
    # text line (x=L+30 works at 100-150% DPI, y=30% is fill without glyphs)
    $tabX = $s.L + 30
    $tabY = $s.T + [int]($s.H * 0.55)
    $pxX = $s.L + 30
    $pxY = $s.T + [int]($s.H * 0.30)

    $before = Get-ScreenPixel $pxX $pxY
    $wordPid = [uint32](Get-Process WINWORD | Select-Object -First 1).Id

    # --- right-click the first tab -> context menu appears
    [W]::RightClick($tabX, $tabY)
    Step 'context menu opened' (Wait-Menu $wordPid $true) ''
    Snap 'e2e4-menu-open'

    # --- menu: Pin / --- / Close / Close others / --- / Tab color / --- / Open folder
    #     Down x4 -> "Tab color", Right -> submenu (first item = Red), Enter
    [W]::Key(0x28); [W]::Key(0x28); [W]::Key(0x28); [W]::Key(0x28)   # Down x4
    [W]::Key(0x27)                                    # Right (open submenu)
    Snap 'e2e4-submenu-open'
    [W]::Key(0x0D)                                    # Enter -> "Red" preset
    Start-Sleep -Seconds 2

    $log1 = Read-NewLog
    Step 'color set logged (D93025)' ($log1 -match 'Tab color set: D93025') ''
    Step 'menu dismissed after pick' (Wait-Menu $wordPid $false 3000) ''

    $after = Get-ScreenPixel $pxX $pxY
    $reddish = ($after.R - $after.B) -gt 20 -and $after.R -gt 225
    Step 'tab fill turned pastel red' $reddish "before=$($before.R),$($before.G),$($before.B) after=$($after.R),$($after.G),$($after.B)"
    Snap 'e2e4-tab-colored'

    $cfgHasColor = (Test-Path $colorsPath) -and ((Get-Content $colorsPath -Raw) -match 'D93025\|.*alpha')
    Step 'color persisted to tab-colors.cfg' $cfgHasColor ''

    # --- reset: right-click same tab, Down x4, Right, Up (wrap to "Reset"), Enter
    [W]::RightClick($tabX, $tabY)
    if (-not (Wait-Menu $wordPid $true)) { Step 'reset: menu opened' $false ''; throw 'menu did not open' }
    [W]::Key(0x28); [W]::Key(0x28); [W]::Key(0x28); [W]::Key(0x28)
    [W]::Key(0x27)
    [W]::Key(0x26)                                    # Up -> wraps to last item = Reset
    [W]::Key(0x0D)
    Start-Sleep -Seconds 2
    $log2 = Read-NewLog
    Step 'color reset logged' ($log2 -match 'Tab color cleared') ''
    $cfgCleared = -not ((Test-Path $colorsPath) -and ((Get-Content $colorsPath -Raw) -match 'alpha'))
    Step 'reset removed entry from tab-colors.cfg' $cfgCleared ''

    # --- set color again (green: Down x3 into submenu after opening) for restart check
    [W]::RightClick($tabX, $tabY)
    if (-not (Wait-Menu $wordPid $true)) { Step 'green: menu opened' $false ''; throw 'menu did not open' }
    [W]::Key(0x28); [W]::Key(0x28); [W]::Key(0x28); [W]::Key(0x28)
    [W]::Key(0x27)
    [W]::Key(0x28); [W]::Key(0x28); [W]::Key(0x28)    # Down x3 -> Green (4th preset)
    [W]::Key(0x0D)
    Start-Sleep -Seconds 2
    $log3 = Read-NewLog
    Step 'second color set logged (1E8E3E)' ($log3 -match 'Tab color set: 1E8E3E') ''

    foreach ($doc in @($word.Documents)) { $doc.Close([ref]0) }
    Start-Sleep -Seconds 2
    $word.Quit()
    Start-Sleep -Seconds 3
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch {}
    $word = $null

    $s1errors = @((Read-NewLog) -split "`n" | Where-Object { $_ -match '\[ERROR\]' })
    Step 'no ERROR lines in add-in log (session 1)' ($s1errors.Count -eq 0) "errors=$($s1errors.Count)"

    # ---------- session 2: color survives Word restart ----------
    $logStart = 0
    if (Test-Path $logFile) { $logStart = (Get-Item $logFile).Length }

    Start-Process 'winword.exe' -ArgumentList "`"$alphaFile`""
    $deadline = (Get-Date).AddSeconds(45)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 2
        try {
            $word = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
            if ($word -and $word.Documents.Count -ge 1) { break }
        } catch { $word = $null }
    }
    if (-not $word) { throw 'Could not attach to restarted Word' }
    Start-Sleep -Seconds 4

    $log4 = Read-NewLog
    Step 'restart: stored colors loaded' ($log4 -match 'Tab colors loaded: 1') ''

    $activeTop2 = [W]::GetAncestor([IntPtr]$word.ActiveWindow.Hwnd, 2)
    [void][W]::SetForegroundWindow($activeTop2)
    $word.ActiveWindow.WindowState = 1
    Start-Sleep -Seconds 2
    $s2 = (Get-HostSurface $activeTop2)[0]
    if ($s2) {
        $px2 = Get-ScreenPixel ($s2.L + 30) ($s2.T + [int]($s2.H * 0.30))
        $greenish = ($px2.G - $px2.R) -gt 10 -and $px2.G -gt 215
        Step 'restart: tab painted pastel green' $greenish "px=$($px2.R),$($px2.G),$($px2.B)"
        Snap 'e2e4-restart-colored'
    } else {
        Step 'restart: host present' $false 'no host surface'
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
    # restore user's config and color store
    if ($null -ne $cfgBackup) { Set-Content -Path $cfgPath -Value $cfgBackup -NoNewline }
    elseif (Test-Path $cfgPath) { Remove-Item $cfgPath -Force }   # cfg-fayla ne bylo - ne ostavlyaem svoy
    if ($null -ne $colorsBackup) { Set-Content -Path $colorsPath -Value $colorsBackup -NoNewline -Encoding UTF8 }
    else { Remove-Item $colorsPath -Force -ErrorAction SilentlyContinue }
}

Start-Sleep -Seconds 2
$still = Get-Process WINWORD -ErrorAction SilentlyContinue
Step 'WINWORD exited cleanly' (-not $still) ''

$tail = Read-NewLog
Write-Host ''
Write-Host '=== Add-in log (session 2) ==='
Write-Host $tail
$errors = @($tail -split "`n" | Where-Object { $_ -match '\[ERROR\]' })
Step 'no ERROR lines in add-in log' ($errors.Count -eq 0) "errors=$($errors.Count)"

Write-Host ''
Write-Host '=== SUMMARY ==='
$results | ForEach-Object { Write-Host $_ }
$failCount = @($results | Where-Object { $_ -like '`[FAIL*' }).Count
Write-Host "FAILURES: $failCount"
exit $failCount
