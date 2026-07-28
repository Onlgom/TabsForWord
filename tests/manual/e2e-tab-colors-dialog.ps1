# E2E addon: "Other color..." opens the system ChooseColor dialog inside Word,
# OK applies the color, Word stays alive. ASCII only. PowerShell 5.1.
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
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] public static extern IntPtr GetDlgItem(IntPtr h, int id);
    [DllImport("user32.dll")] public static extern IntPtr SendMessageW(IntPtr h, uint msg, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L; public int T; public int R; public int B; }
    public static IntPtr FindDialogOfProcess(uint pid) {
        IntPtr found = IntPtr.Zero;
        EnumWindows(delegate(IntPtr h, IntPtr l) {
            if (!IsWindowVisible(h)) return true;
            if (Cls(h) != "#32770") return true;
            uint p; GetWindowThreadProcessId(h, out p);
            if (p != pid) return true;
            found = h; return false;
        }, IntPtr.Zero);
        return found;
    }
    public static bool ClickOk(IntPtr dlg) {
        IntPtr ok = GetDlgItem(dlg, 1); // IDOK
        if (ok == IntPtr.Zero) return false;
        SendMessageW(ok, 0x00F5, IntPtr.Zero, IntPtr.Zero); // BM_CLICK
        return true;
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
    public static bool DialogOpen() { return FindWindowW("#32770", null) != IntPtr.Zero; }
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

$p = Get-Process WINWORD -ErrorAction SilentlyContinue
if ($p) { Write-Host 'ABORT: WINWORD is running.'; exit 2 }

$installDir = Join-Path $env:LOCALAPPDATA 'TabsForWord'
Copy-Item (Join-Path $repo 'src\TabsForWord\bin\Release\TabsForWord.dll') (Join-Path $installDir 'TabsForWord.dll') -Force

$cfgPath = Join-Path $installDir 'native-host.cfg'
$cfgBackup = $null
if (Test-Path $cfgPath) { $cfgBackup = Get-Content $cfgPath -Raw }
Set-Content -Path $cfgPath -Value "mode=native`r`ndump=0`r`nreserve=1" -Encoding Ascii

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

$alphaFile = Join-Path $dir 'v5-dialog-alpha.docx'
Remove-Item $alphaFile -Force -ErrorAction SilentlyContinue

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

    foreach ($doc in @($word.Documents)) {
        if (-not $doc.Path) { $doc.Close([ref]0) }
    }
    Start-Sleep -Seconds 1
    $d1 = $word.Documents.Add(); $d1.Content.Text = 'Alpha'
    $d1.SaveAs([string]$alphaFile)
    Start-Sleep -Seconds 2

    $activeTop = [W]::GetAncestor([IntPtr]$word.ActiveWindow.Hwnd, 2)
    [void][W]::SetForegroundWindow($activeTop)
    $word.ActiveWindow.WindowState = 1
    Start-Sleep -Seconds 2

    $s = (Get-HostSurface $activeTop)[0]
    Step 'host present' ($null -ne $s -and $s.Visible) "rect=$($s.L),$($s.T) $($s.W)x$($s.H)"

    # right-click tab -> Down x4 (Pin/Close/CloseOthers/Color) -> Right (submenu, first=Red) ->
    # Down x8 walks 7 presets + separator to "Other color..." -> Enter opens dialog
    # (arrows do NOT skip disabled items, so Up would stop on grayed "Reset")
    $wordPid2 = [uint32](Get-Process WINWORD | Select-Object -First 1).Id
    [W]::RightClick($s.L + 30, $s.T + [int]($s.H * 0.55))
    $menuDeadline = [Environment]::TickCount + 5000
    while ([Environment]::TickCount -lt $menuDeadline -and
           [W]::FindMenuOfProcess($wordPid2) -eq [IntPtr]::Zero) { Start-Sleep -Milliseconds 100 }
    [W]::Key(0x28); [W]::Key(0x28); [W]::Key(0x28); [W]::Key(0x28)
    [W]::Key(0x27)
    for ($i = 0; $i -lt 8; $i++) { [W]::Key(0x28) }
    [W]::Key(0x0D)
    Start-Sleep -Seconds 2
    $wordPid = (Get-Process WINWORD).Id
    $dlg = [W]::FindDialogOfProcess([uint32]$wordPid)
    Step 'ChooseColor dialog opened inside Word' ($dlg -ne [IntPtr]::Zero) "dlg=$dlg"

    $clicked = [W]::ClickOk($dlg)   # IDOK via BM_CLICK (no focus dependency)
    Start-Sleep -Seconds 2
    $dlgAfter = [W]::FindDialogOfProcess([uint32]$wordPid)
    Step 'dialog closed by OK' ($clicked -and $dlgAfter -eq [IntPtr]::Zero) ''

    $log1 = Read-NewLog
    Step 'custom color applied and logged' ($log1 -match 'Tab color set: ') ''
    $cfgHasColor = (Test-Path $colorsPath) -and ((Get-Content $colorsPath -Raw) -match '\|.*alpha')
    Step 'custom color persisted to tab-colors.cfg' $cfgHasColor ''

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
    if ($null -ne $colorsBackup) { Set-Content -Path $colorsPath -Value $colorsBackup -NoNewline -Encoding UTF8 }
    else { Remove-Item $colorsPath -Force -ErrorAction SilentlyContinue }
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
