# E2E: tab pinning (context menu, reorders to front, persists across restart),
# tab order persistence across restart, and the Settings popup (opened from the
# footer of the "all tabs" menu) - tab size preset + Ctrl+Tab hotkey toggle.
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
        System.Threading.Thread.Sleep(180);
    }
    public static void CtrlTab() {
        keybd_event(0x11, 0, 0, UIntPtr.Zero);       // Ctrl down
        System.Threading.Thread.Sleep(60);
        keybd_event(0x09, 0, 0, UIntPtr.Zero);        // Tab down
        System.Threading.Thread.Sleep(60);
        keybd_event(0x09, 0, 2, UIntPtr.Zero);        // Tab up
        System.Threading.Thread.Sleep(60);
        keybd_event(0x11, 0, 2, UIntPtr.Zero);        // Ctrl up
        System.Threading.Thread.Sleep(250);
    }
    public static bool MenuOpen() { return FindWindowW("#32768", null) != IntPtr.Zero; }
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr FindWindowW(string cls, string title);
}
"@
[W]::SetProcessDpiAwareness(2) | Out-Null

function Wait-Menu([uint32]$wordPid, [bool]$shouldBeOpen, [int]$timeoutMs = 5000) {
    $deadline = [Environment]::TickCount + $timeoutMs
    while ([Environment]::TickCount -lt $deadline) {
        $isOpen = ([W]::FindWindowW('#32768', $null) -ne [IntPtr]::Zero)
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
                L = $r.L; T = $r.T; R = $r.R; B = $r.B; W = ($r.R - $r.L); H = ($r.B - $r.T)
                Visible = [W]::IsWindowVisible($h)
            })
        }
        return $true
    }
    [void][W]::EnumChildWindows($wordTop, $cb, [IntPtr]::Zero)
    return @($list | Where-Object { $_.Parent -eq $wordTop -and $_.Visible })
}

# Popups (TabListPopup / SettingsForm) are TOP-LEVEL WinForms windows (not
# children of Word), so EnumWindows (not EnumChildWindows) finds them. Small
# WindowsForms10-classed windows (tooltips) are excluded by a minimum size.
function Get-TopPopup([uint32]$targetPid) {
    $list = New-Object System.Collections.ArrayList
    $cb = [W+EnumProc]{
        param($h, $l)
        if ([W]::IsWindowVisible($h)) {
            $cls = [W]::Cls($h)
            if ($cls.StartsWith('WindowsForms10')) {
                $p = [uint32]0
                [void][W]::GetWindowThreadProcessId($h, [ref]$p)
                if ($p -eq $targetPid) {
                    $r = New-Object W+RECT
                    [void][W]::GetWindowRect($h, [ref]$r)
                    $w = $r.R - $r.L; $ht = $r.B - $r.T
                    if ($w -gt 100 -and $ht -gt 40) {
                        [void]$list.Add([pscustomobject]@{ Hwnd = $h; L = $r.L; T = $r.T; R = $r.R; B = $r.B; W = $w; H = $ht })
                    }
                }
            }
        }
        return $true
    }
    [void][W]::EnumWindows($cb, [IntPtr]::Zero)
    return @($list | Select-Object -First 1)
}

function Wait-TopPopup([uint32]$targetPid, [int]$timeoutMs = 4000) {
    $deadline = [Environment]::TickCount + $timeoutMs
    while ([Environment]::TickCount -lt $deadline) {
        $p = Get-TopPopup $targetPid
        if ($p) { return $p }
        Start-Sleep -Milliseconds 100
    }
    return $null
}

function Snap($name) {
    try {
        $vs = [System.Windows.Forms.SystemInformation]::VirtualScreen
        $bmp = New-Object System.Drawing.Bitmap($vs.Width, $vs.Height)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.CopyFromScreen($vs.X, $vs.Y, 0, 0, $bmp.Size)
        $bmp.Save((Join-Path $dir "$name.png"), [System.Drawing.Imaging.ImageFormat]::Png)
        $g.Dispose(); $bmp.Dispose()
    } catch { }
}

$p = Get-Process WINWORD -ErrorAction SilentlyContinue
if ($p) { Write-Host 'ABORT: WINWORD is running.'; exit 2 }

$installDir = Join-Path $env:LOCALAPPDATA 'TabsForWord'
Copy-Item (Join-Path $repo 'src\TabsForWord\bin\Release\TabsForWord.dll') (Join-Path $installDir 'TabsForWord.dll') -Force

# keep the user's settings intact: back up everything this test touches, restore at the end
$cfgPath = Join-Path $installDir 'native-host.cfg'
$cfgBackup = $null
if (Test-Path $cfgPath) { $cfgBackup = Get-Content $cfgPath -Raw }
Set-Content -Path $cfgPath -Value "mode=native`r`ndump=0`r`nreserve=1" -Encoding Ascii

$orderPath = Join-Path $installDir 'tab-order.cfg'
$orderBackup = $null
if (Test-Path $orderPath) { $orderBackup = Get-Content $orderPath -Raw; Remove-Item $orderPath -Force }

$sizePath = Join-Path $installDir 'tab-size.cfg'
$sizeBackup = $null
if (Test-Path $sizePath) { $sizeBackup = Get-Content $sizePath -Raw; Remove-Item $sizePath -Force }

$hotkeyPath = Join-Path $installDir 'hotkey.cfg'
$hotkeyBackup = $null
if (Test-Path $hotkeyPath) { $hotkeyBackup = Get-Content $hotkeyPath -Raw; Remove-Item $hotkeyPath -Force }

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

function Attach-Word {
    $deadline = (Get-Date).AddSeconds(45)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 2
        try {
            $w = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
            if ($w -and $w.Documents.Count -ge 1) { return $w }
        } catch { }
    }
    throw 'Could not attach to Word'
}

$alphaFile = Join-Path $dir 'v8-order-alpha.docx'
$bravoFile = Join-Path $dir 'v8-order-bravo.docx'
$charlieFile = Join-Path $dir 'v8-order-charlie.docx'
Remove-Item $alphaFile, $bravoFile, $charlieFile -Force -ErrorAction SilentlyContinue

$word = $null
try {
    # ---------- session 1: pin/unpin reorders, settings popup (size + hotkey) ----------
    Start-Process 'winword.exe' -ArgumentList '/w'
    $word = Attach-Word
    Start-Sleep -Seconds 3
    Step 'add-in started (native mode)' ((Read-NewLog) -match 'Native host mode enabled') ''

    foreach ($doc in @($word.Documents)) { if (-not $doc.Path) { $doc.Close([ref]0) } }
    Start-Sleep -Seconds 1

    $d1 = $word.Documents.Add(); $d1.Content.Text = 'Alpha'; $d1.SaveAs([string]$alphaFile)
    Start-Sleep -Milliseconds 700
    $d2 = $word.Documents.Add(); $d2.Content.Text = 'Bravo'; $d2.SaveAs([string]$bravoFile)
    Start-Sleep -Milliseconds 700
    $d3 = $word.Documents.Add(); $d3.Content.Text = 'Charlie'; $d3.SaveAs([string]$charlieFile)
    Start-Sleep -Seconds 2

    $wordTop = [W]::GetAncestor([IntPtr]$word.ActiveWindow.Hwnd, 2)
    [void][W]::SetForegroundWindow($wordTop)
    $word.ActiveWindow.WindowState = 1
    Start-Sleep -Seconds 2
    $wordPid = [uint32](Get-Process WINWORD | Select-Object -First 1).Id

    $s = (Get-HostSurface $wordTop)[0]
    Step 'host present (3 tabs: alpha, bravo, charlie)' ($null -ne $s) "rect=$($s.L),$($s.T) $($s.W)x$($s.H)"
    $scale = $s.H / 37.0   # panel height / spec base height (37px @ 100%/normal) -> DPI*size scale

    $tabY = $s.T + [int]($s.H * 0.55)
    $firstTabX = $s.L + 30

    # --- pin the 3rd tab (charlie): right-click it, then Down x1 -> Enter (Pin is first item)
    # The 3 tabs are laid out left-to-right starting at $s.L+14 (PadLeft); charlie is
    # 3rd. Click 2.5 tab-widths in (typical short-caption tab ~155px at scale=1) to
    # land safely within tab #3's body, well clear of tab #2's boundary.
    $tabW = [int](155 * $scale)
    $charlieX = $s.L + 14 + [int](2.5 * $tabW)
    [W]::RightClick($charlieX, $tabY)
    $opened = Wait-Menu $wordPid $true
    Step 'context menu opened on 3rd tab (charlie)' $opened ''
    if (-not $opened) { throw 'context menu did not open on 3rd tab' }
    Snap 'e2e8-menu-charlie'
    [W]::Key(0x28)   # Down x1 -> "Pin tab" (first item)
    [W]::Key(0x0D)   # Enter
    Start-Sleep -Seconds 2
    $null = Wait-Menu $wordPid $false 3000

    $log1 = Read-NewLog
    Step 'pin set logged' ($log1 -match 'Tab pin set.*Charlie') $log1
    $orderHasPinCharlie = (Test-Path $orderPath) -and ((Get-Content $orderPath -Raw) -match 'PIN\|.*[Cc]harlie')
    Step 'tab-order.cfg has PIN|...charlie' $orderHasPinCharlie ''

    # charlie should now be FIRST: clicking the first-tab position must activate it
    [W]::Click($firstTabX, $tabY)
    Start-Sleep -Seconds 2
    $activeAfterPin = $word.ActiveDocument.Name
    Step 'pinned tab moved to front (first-tab click activates charlie)' ($activeAfterPin -like '*harlie*') "active=$activeAfterPin"
    Snap 'e2e8-charlie-pinned-front'

    # --- unpin it (still the first tab): right-click first tab, Down x1 -> Enter.
    # NOTE: unpinning does NOT teleport a tab back to its pre-pin position (same
    # as Chrome/Edge) - the pinned-first partition is a STABLE sort, so with no
    # other pinned tab to sort against, charlie simply stays wherever it already
    # is (first). We only assert the log here; position is re-verified below.
    [W]::RightClick($firstTabX, $tabY)
    $opened2 = Wait-Menu $wordPid $true
    Step 'context menu opened on pinned (now first) tab' $opened2 ''
    if (-not $opened2) { throw 'context menu did not open on pinned tab' }
    [W]::Key(0x28)
    [W]::Key(0x0D)
    Start-Sleep -Seconds 2
    $null = Wait-Menu $wordPid $false 3000

    $log2 = Read-NewLog
    Step 'pin cleared logged' ($log2 -match 'Tab pin cleared.*harlie') ''

    # --- pin it again (still first tab), for the restart-persistence check below
    [W]::RightClick($firstTabX, $tabY)
    if (-not (Wait-Menu $wordPid $true)) { throw 'context menu did not open on charlie' }
    [W]::Key(0x28)
    [W]::Key(0x0D)
    Start-Sleep -Seconds 2
    $null = Wait-Menu $wordPid $false 3000
    $log3 = Read-NewLog
    Step 'charlie re-pinned (for restart check)' ($log3 -match 'Tab pin set.*harlie') ''

    # ---------- settings popup: size preset + Ctrl+Tab hotkey toggle ----------
    # Menu (Right block): ... [Plus] [sep] [Menu-btn v] [Collapse-btn] with a fixed
    # 12px right padding; Menu-btn is the 2nd button from the panel's right edge.
    $menuBtnX = $s.R - [int](53 * $scale)
    $menuBtnY = $s.T + [int](19 * $scale)
    [W]::Click($menuBtnX, $menuBtnY)
    $popup = Wait-TopPopup $wordPid
    Step 'all-tabs popup opened' ($null -ne $popup) $(if ($popup) { "rect=$($popup.L),$($popup.T) $($popup.W)x$($popup.H)" } else { '' })
    if (-not $popup) { throw 'all-tabs popup (TabListPopup) did not appear' }
    Snap 'e2e8-alltabs-popup'

    # footer "Settings..." row: last ~26px row before the bottom padding
    $footerX = $popup.L + 40
    $footerY = $popup.B - [int](19 * $scale)
    [W]::Click($footerX, $footerY)
    $settingsForm = Wait-TopPopup $wordPid
    Step 'settings window opened' ($null -ne $settingsForm) $(if ($settingsForm) { "rect=$($settingsForm.L),$($settingsForm.T) $($settingsForm.W)x$($settingsForm.H)" } else { '' })
    if (-not $settingsForm) { throw 'SettingsForm did not appear' }
    Snap 'e2e8-settings-open'

    # "Large" (3rd of 4) size radio: header(28)+section(24)+2 rows(28 each)+half-row(14)
    $largeRadioY = $settingsForm.T + [int]((14 + 28 + 24 + 2 * 28 + 14) * $scale)
    [W]::Click(($settingsForm.L + 30), $largeRadioY)
    Start-Sleep -Seconds 2
    $log4 = Read-NewLog
    Step 'settings: size preset applied (saved: 1.15)' ($log4 -match 'Tab size scale saved: 1\.15') ''
    $sAfterSize = (Get-HostSurface $wordTop)[0]
    Step 'settings: panel grew with Large preset' ($sAfterSize.H -gt $s.H) "H=$($s.H) -> $($sAfterSize.H)"

    # Ctrl+Tab checkbox: header(28)+size-section(24)+4 rows(28 each)+gap(8)+kb-section(24)+half of (28+12)
    $hotkeyCheckY = $settingsForm.T + [int]((14 + 28 + 24 + 4 * 28 + 8 + 24 + 20) * $scale)
    [W]::Click(($settingsForm.L + 30), $hotkeyCheckY)
    Start-Sleep -Seconds 1
    $log5 = Read-NewLog
    Step 'settings: hotkey disabled logged' ($log5 -match 'Ctrl\+Tab hotkey disabled from settings') $log5
    Step 'hotkey.cfg written (enabled=0)' ((Test-Path $hotkeyPath) -and ((Get-Content $hotkeyPath -Raw) -match 'enabled=0')) ''

    # click away to close the settings popup, then focus the document
    [W]::Click(($s.L + [int]($s.W * 0.5)), ($s.T + $sAfterSize.H + 60))
    Start-Sleep -Milliseconds 800
    [void][W]::SetForegroundWindow($wordTop)
    Start-Sleep -Milliseconds 500

    # Ctrl+Tab must now be inert (hotkey disabled)
    $beforeCT = $word.ActiveDocument.Name
    [W]::CtrlTab()
    Start-Sleep -Seconds 1
    $afterCT = $word.ActiveDocument.Name
    Step 'Ctrl+Tab does nothing while hotkey disabled' ($afterCT -eq $beforeCT) "active stayed: $afterCT"

    # re-enable via settings (popup re-opened fresh). The "Large" preset applied
    # above changed the panel's own scale, so recompute it from the CURRENT
    # panel height rather than reusing the pre-resize $scale.
    $sNow = (Get-HostSurface $wordTop)[0]
    $scale2 = $sNow.H / 37.0
    [W]::Click(($sNow.R - [int](53 * $scale2)), ($sNow.T + [int](19 * $scale2)))
    $popup2 = Wait-TopPopup $wordPid
    Step 'all-tabs popup re-opened' ($null -ne $popup2) ''
    if (-not $popup2) { throw 'all-tabs popup did not reappear for hotkey re-enable' }

    [W]::Click(($popup2.L + 40), ($popup2.B - [int](19 * $scale2)))
    $settingsForm2 = Wait-TopPopup $wordPid
    Step 'settings window re-opened' ($null -ne $settingsForm2) ''
    if (-not $settingsForm2) { throw 'SettingsForm did not reappear for hotkey re-enable' }

    $hotkeyCheckY2 = $settingsForm2.T + [int]((14 + 28 + 24 + 4 * 28 + 8 + 24 + 20) * $scale2)
    [W]::Click(($settingsForm2.L + 30), $hotkeyCheckY2)
    Start-Sleep -Seconds 1
    [W]::Click(($sNow.L + [int]($sNow.W * 0.5)), ($sNow.T + $sNow.H + 60))
    Start-Sleep -Milliseconds 500

    $log6 = Read-NewLog
    Step 're-enabled hotkey logged' ($log6 -match 'Ctrl\+Tab hotkey enabled from settings') $log6

    [void][W]::SetForegroundWindow($wordTop)
    Start-Sleep -Milliseconds 500
    $beforeCT2 = $word.ActiveDocument.Name
    [W]::CtrlTab()
    Start-Sleep -Seconds 1
    $afterCT2 = $word.ActiveDocument.Name
    Step 'Ctrl+Tab switches tabs again after re-enabling' ($afterCT2 -ne $beforeCT2) "$beforeCT2 -> $afterCT2"

    foreach ($doc in @($word.Documents)) { $doc.Close([ref]0) }
    Start-Sleep -Seconds 2
    $word.Quit()
    Start-Sleep -Seconds 3
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch {}
    $word = $null

    $s1errors = @((Read-NewLog) -split "`n" | Where-Object { $_ -match '\[ERROR\]' })
    Step 'no ERROR lines in add-in log (session 1)' ($s1errors.Count -eq 0) "errors=$($s1errors.Count)"

    # ---------- session 2: order + pin survive restart, even opened out of order ----------
    $logStart = 0
    if (Test-Path $logFile) { $logStart = (Get-Item $logFile).Length }

    # Reopen out of saved order, pinned doc (charlie) opened LAST - persisted
    # order (and its pinned-first placement) must still win.
    Start-Process 'winword.exe' -ArgumentList "`"$bravoFile`""
    $word = Attach-Word
    Start-Sleep -Seconds 3
    $word.Documents.Open([string]$alphaFile) | Out-Null
    Start-Sleep -Milliseconds 700
    $word.Documents.Open([string]$charlieFile) | Out-Null
    Start-Sleep -Seconds 2

    $log7 = Read-NewLog
    # Detail = the actual "Tab order loaded" line, not the whole log slice: when a
    # document left over from an earlier run sneaks into tab-order.cfg the count is
    # off by one, and the failure message has to say so by itself.
    $loadedLine = (($log7 -split "`n") | Where-Object { $_ -match 'Tab order loaded:' } | Select-Object -Last 1)
    Step 'restart: stored order loaded' ($log7 -match 'Tab order loaded: 3') $loadedLine

    $wordTop2 = [W]::GetAncestor([IntPtr]$word.ActiveWindow.Hwnd, 2)
    [void][W]::SetForegroundWindow($wordTop2)
    $word.ActiveWindow.WindowState = 1
    Start-Sleep -Seconds 2
    $s2 = (Get-HostSurface $wordTop2)[0]
    if ($s2) {
        $tabY2 = $s2.T + [int]($s2.H * 0.55)
        [W]::Click(($s2.L + 30), $tabY2)
        Start-Sleep -Seconds 2
        $activeRestart = $word.ActiveDocument.Name
        Step 'restart: pinned charlie is first tab despite opened last' ($activeRestart -like '*harlie*') "active=$activeRestart"
        Snap 'e2e8-restart-order'
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
    if ($null -ne $cfgBackup) { Set-Content -Path $cfgPath -Value $cfgBackup -NoNewline }
    elseif (Test-Path $cfgPath) { Remove-Item $cfgPath -Force }   # cfg-fayla ne bylo - ne ostavlyaem svoy
    if ($null -ne $orderBackup) { Set-Content -Path $orderPath -Value $orderBackup -NoNewline -Encoding UTF8 } else { Remove-Item $orderPath -Force -ErrorAction SilentlyContinue }
    if ($null -ne $sizeBackup) { Set-Content -Path $sizePath -Value $sizeBackup -NoNewline -Encoding UTF8 } else { Remove-Item $sizePath -Force -ErrorAction SilentlyContinue }
    if ($null -ne $hotkeyBackup) { Set-Content -Path $hotkeyPath -Value $hotkeyBackup -NoNewline -Encoding UTF8 } else { Remove-Item $hotkeyPath -Force -ErrorAction SilentlyContinue }
}

Start-Sleep -Seconds 2
$still = Get-Process WINWORD -ErrorAction SilentlyContinue
Step 'WINWORD exited cleanly' (-not $still) ''

$tail = Read-NewLog
$errors = @($tail -split "`n" | Where-Object { $_ -match '\[ERROR\]' })
Step 'no ERROR lines in add-in log (session 2)' ($errors.Count -eq 0) "errors=$($errors.Count)"

Write-Host ''
Write-Host '=== SUMMARY ==='
$results | ForEach-Object { Write-Host $_ }
$failCount = @($results | Where-Object { $_ -like '`[FAIL*' }).Count
Write-Host "FAILURES: $failCount"
exit $failCount
