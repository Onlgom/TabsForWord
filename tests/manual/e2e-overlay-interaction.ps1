# E2E phase 3: tab-body activation click, ribbon collapse via double-click, Ctrl+Tab.
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
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L; public int T; public int R; public int B; }
    public static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
    public static void Click(int x, int y) {
        SetCursorPos(x, y);
        System.Threading.Thread.Sleep(150);
        mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero);
        System.Threading.Thread.Sleep(80);
        mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero);
    }
    public static void DoubleClick(int x, int y) {
        Click(x, y);
        System.Threading.Thread.Sleep(120);
        mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero);
        System.Threading.Thread.Sleep(60);
        mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero);
    }
    public static void CtrlTab(bool shift) {
        keybd_event(0x11, 0, 0, UIntPtr.Zero);
        if (shift) keybd_event(0x10, 0, 0, UIntPtr.Zero);
        System.Threading.Thread.Sleep(60);
        keybd_event(0x09, 0, 0, UIntPtr.Zero);
        System.Threading.Thread.Sleep(60);
        keybd_event(0x09, 0, 2, UIntPtr.Zero);
        System.Threading.Thread.Sleep(60);
        if (shift) keybd_event(0x10, 0, 2, UIntPtr.Zero);
        keybd_event(0x11, 0, 2, UIntPtr.Zero);
    }
}
"@
[W]::SetProcessDPIAware() | Out-Null

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
# The user mode is saved and restored in finally (see e2e-smoke.ps1): the script
# puts the add-in into overlay, while the user normally has reserve=1.
$cfgPath = Join-Path $installDir 'native-host.cfg'
$cfgBackup = $null
if (Test-Path $cfgPath) { $cfgBackup = Get-Content $cfgPath -Raw }
Set-Content -Path $cfgPath -Value "mode=native`r`ndump=0`r`nreserve=0" -Encoding Ascii

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
    Step 'add-in started (native mode)' ((Read-NewLog) -match 'Native host mode enabled') ''

    # add 2 saved docs (safe to close silently later)
    $d1 = $word.Documents.Add(); $d1.Content.Text = 'Alpha'
    $d1.SaveAs([string](Join-Path $dir 'v3-alpha.docx'))
    Start-Sleep -Milliseconds 900
    $d2 = $word.Documents.Add(); $d2.Content.Text = 'Bravo'
    $d2.SaveAs([string](Join-Path $dir 'v3-bravo.docx'))
    Start-Sleep -Seconds 2

    $activeTop = [W]::GetAncestor([IntPtr]$word.ActiveWindow.Hwnd, 2)
    [void][W]::SetForegroundWindow($activeTop)
    $word.ActiveWindow.WindowState = 1
    Start-Sleep -Seconds 2

    $s = (Get-HostSurface $activeTop)[0]
    Step 'host present' ($null -ne $s -and $s.Visible) "rect=$($s.L),$($s.T) $($s.W)x$($s.H)"

    # --- activation click on the BODY of the first tab (x=30 avoids hover close box)
    $activeBefore = $word.ActiveDocument.Name
    [W]::Click($s.L + 30, $s.T + [int]($s.H * 0.55))
    Start-Sleep -Seconds 2
    $activeAfter = $word.ActiveDocument.Name
    $log1 = Read-NewLog
    Step 'tab body click activates first document' (($activeAfter -ne $activeBefore) -and ($log1 -match 'Window activated')) "before=$activeBefore after=$activeAfter"
    Snap 'e2e3-tab-activated'

    # --- ribbon collapse via double-click on ribbon tab (approx 93,63 at 100% DPI)
    $curTop = [W]::GetAncestor([IntPtr]$word.ActiveWindow.Hwnd, 2)
    $sB = (Get-HostSurface $curTop)[0]
    $tBefore = $sB.T
    [W]::DoubleClick(93, 63)
    Start-Sleep -Seconds 2
    $sC = (Get-HostSurface $curTop)[0]
    Step 'ribbon collapse moves panel up' ($sC.T -lt $tBefore) "t=$tBefore -> $($sC.T)"
    Snap 'e2e3-ribbon-collapsed'
    [W]::DoubleClick(93, 63)
    Start-Sleep -Seconds 2
    $sD = (Get-HostSurface $curTop)[0]
    Step 'ribbon expand moves panel back' ($sD.T -eq $tBefore) "t=$($sC.T) -> $($sD.T)"

    # --- Ctrl+Tab cycles documents (keyboard hook still works in native mode)
    $before = $word.ActiveDocument.Name
    [W]::CtrlTab($false)
    Start-Sleep -Seconds 2
    $after1 = $word.ActiveDocument.Name
    Step 'Ctrl+Tab switches to next tab' ($after1 -ne $before) "before=$before after=$after1"
    [W]::CtrlTab($true)
    Start-Sleep -Seconds 2
    $after2 = $word.ActiveDocument.Name
    Step 'Ctrl+Shift+Tab returns back' ($after2 -eq $before) "now=$after2"

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
Write-Host ''
Write-Host '=== Add-in log (this run) ==='
Write-Host $tail
$errors = @($tail -split "`n" | Where-Object { $_ -match '\[ERROR\]' })
Step 'no ERROR lines in add-in log' ($errors.Count -eq 0) "errors=$($errors.Count)"

Write-Host ''
Write-Host '=== SUMMARY ==='
$results | ForEach-Object { Write-Host $_ }
$failCount = @($results | Where-Object { $_ -like '`[FAIL*' }).Count
Write-Host "FAILURES: $failCount"
exit $failCount
