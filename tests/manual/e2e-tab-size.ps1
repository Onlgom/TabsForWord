# E2E: manual tab size via context menu (Large -> restart persists -> Normal).
# Menu path: right-click tab -> Down x5 (Pin/Close/CloseOthers/Color/Size) -> Right ->
# submenu (Compact/Normal/Large/XLarge). ASCII only. PowerShell 5.1.
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
    [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L; public int T; public int R; public int B; }
    public static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
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
[W]::SetProcessDpiAwareness(2) | Out-Null

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
    return @($list | Where-Object { $_.Parent -eq $wordTop -and $_.Visible })
}

function Snap($name, $s) {
    try {
        $bmp = New-Object System.Drawing.Bitmap($s.W, ($s.H + 8))
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.CopyFromScreen($s.L, $s.T, 0, 0, $bmp.Size)
        $bmp.Save((Join-Path $dir "$name.png"), [System.Drawing.Imaging.ImageFormat]::Png)
        $g.Dispose(); $bmp.Dispose()
    } catch { }
}

# Menu: right-click -> Down x5 -> Right -> Down x(i) -> Enter (i: 0=Compact..3=Extra large)
function Choose-Size($s, [int]$downsInSubmenu) {
    [W]::RightClick($s.L + 30, $s.T + [int]($s.H * 0.55))
    if (-not (Wait-Menu $script:wordPid $true)) { throw 'context menu did not open' }
    [W]::Key(0x28); [W]::Key(0x28); [W]::Key(0x28); [W]::Key(0x28); [W]::Key(0x28)   # Down x5 -> Tab size
    [W]::Key(0x27)                                                    # Right -> submenu
    for ($i = 0; $i -lt $downsInSubmenu; $i++) { [W]::Key(0x28) }
    [W]::Key(0x0D)
    $null = Wait-Menu $script:wordPid $false 3000
    Start-Sleep -Seconds 2
}

$p = Get-Process WINWORD -ErrorAction SilentlyContinue
if ($p) { Write-Host 'ABORT: WINWORD is running.'; exit 2 }

$installDir = Join-Path $env:LOCALAPPDATA 'TabsForWord'
Copy-Item (Join-Path $repo 'src\TabsForWord\bin\Release\TabsForWord.dll') (Join-Path $installDir 'TabsForWord.dll') -Force

# save the user size setting and start from a clean state
$sizePath = Join-Path $installDir 'tab-size.cfg'
$sizeBackup = $null
if (Test-Path $sizePath) { $sizeBackup = Get-Content $sizePath -Raw; Remove-Item $sizePath -Force }

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

$alphaFile = Join-Path $dir 'v7-size-alpha.docx'
Remove-Item $alphaFile -Force -ErrorAction SilentlyContinue

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

$word = $null
try {
    # ---------- session 1: Large through the menu ----------
    Start-Process 'winword.exe' -ArgumentList '/w'
    $word = Attach-Word
    Start-Sleep -Seconds 3

    foreach ($doc in @($word.Documents)) { if (-not $doc.Path) { $doc.Close([ref]0) } }
    Start-Sleep -Seconds 1
    $d1 = $word.Documents.Add(); $d1.Content.Text = 'Alpha'
    $d1.SaveAs([string]$alphaFile)
    Start-Sleep -Seconds 2

    $wordTop = [W]::GetAncestor([IntPtr]$word.ActiveWindow.Hwnd, 2)
    [void][W]::SetForegroundWindow($wordTop)
    $word.ActiveWindow.WindowState = 1
    Start-Sleep -Seconds 2

    $script:wordPid = [uint32](Get-Process WINWORD | Select-Object -First 1).Id
    $s0 = (Get-HostSurface $wordTop)[0]
    Step 'host present' ($null -ne $s0) "H=$($s0.H)"

    # Every expectation below is computed from the baseline, so the baseline has
    # to BE the Normal size, not merely whatever the strip happens to be.
    # Deleting tab-size.cfg turned out not to be enough: in the stage 34
    # rehearsal one run started at 1.3 with the file gone (never reproduced),
    # and all three size expectations shifted with it. The log line is the
    # authority here - it says which multiplier was actually applied.
    Choose-Size $s0 1
    $s0 = (Get-HostSurface $wordTop)[0]
    $baseOk = @((Read-NewLog) -split "`n" | Where-Object { $_ -match 'Tab size scale saved: 1\s*$' }).Count -ge 1
    Step 'baseline forced to the Normal size (x1.0)' ($baseOk -and $null -ne $s0) "H=$($s0.H)"
    $baseH = $s0.H
    Snap 'e2e6-size-normal' $s0

    # "Large" = the 3rd submenu item (Down x2 from the first)
    Choose-Size $s0 2
    $s1 = (Get-HostSurface $wordTop)[0]
    $expLarge = [int][Math]::Round($baseH * 1.15)
    Step "Large: host height ~x1.15 (expect ~$expLarge)" ([Math]::Abs($s1.H - $expLarge) -le 2) "H=$($s0.H) -> $($s1.H)"
    Snap 'e2e6-size-large' $s1

    $log1 = Read-NewLog
    Step 'scale saved to log (1.15)' ($log1 -match 'Tab size scale saved: 1\.15') ''
    Step 'tab-size.cfg written' ((Test-Path $sizePath) -and ((Get-Content $sizePath -Raw) -match 'scale=1\.15')) ''

    foreach ($doc in @($word.Documents)) { $doc.Close([ref]0) }
    Start-Sleep -Seconds 2
    $word.Quit()
    Start-Sleep -Seconds 3
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch {}
    $word = $null

    # ---------- session 2: the size survived the restart; back to Normal ----------
    $logStart = 0
    if (Test-Path $logFile) { $logStart = (Get-Item $logFile).Length }

    Start-Process 'winword.exe' -ArgumentList "`"$alphaFile`""
    $word = Attach-Word
    Start-Sleep -Seconds 4

    $wordTop2 = [W]::GetAncestor([IntPtr]$word.ActiveWindow.Hwnd, 2)
    [void][W]::SetForegroundWindow($wordTop2)
    $word.ActiveWindow.WindowState = 1
    Start-Sleep -Seconds 2

    $script:wordPid = [uint32](Get-Process WINWORD | Select-Object -First 1).Id
    $s2 = (Get-HostSurface $wordTop2)[0]
    Step "restart: Large size persisted (expect ~$expLarge)" ([Math]::Abs($s2.H - $expLarge) -le 2) "H=$($s2.H)"
    Snap 'e2e6-size-restart' $s2

    # back to "Normal" (the 2nd item, Down x1)
    Choose-Size $s2 1
    $s3 = (Get-HostSurface $wordTop2)[0]
    Step "back to Normal (expect ~$baseH)" ([Math]::Abs($s3.H - $baseH) -le 2) "H=$($s2.H) -> $($s3.H)"

    foreach ($doc in @($word.Documents)) { $doc.Close([ref]0) }
    Start-Sleep -Seconds 2
    $word.Quit()
    Start-Sleep -Seconds 3
} catch {
    Step 'scenario exception' $false $_.Exception.Message
    try { if ($word) { $word.Quit() } } catch {}
} finally {
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch {}
    if ($null -ne $sizeBackup) { Set-Content -Path $sizePath -Value $sizeBackup -NoNewline -Encoding UTF8 }
    else { Remove-Item $sizePath -Force -ErrorAction SilentlyContinue }
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
