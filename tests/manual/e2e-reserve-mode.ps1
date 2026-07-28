# E2E phase 4: reserve mode (shift document area below the panel instead of overlaying).
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
    [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr h, int x, int y, int w, int hh, bool repaint);
    [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr h, uint f);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint flags, int dx, int dy, uint data, UIntPtr extra);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L; public int T; public int R; public int B; }
    public static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
    public static void DoubleClick(int x, int y) {
        SetCursorPos(x, y);
        System.Threading.Thread.Sleep(150);
        mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero);
        System.Threading.Thread.Sleep(60);
        mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero);
        System.Threading.Thread.Sleep(120);
        mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero);
        System.Threading.Thread.Sleep(60);
        mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero);
    }
}
"@
[W]::SetProcessDPIAware() | Out-Null

function Get-ChildInfo([IntPtr]$wordTop, [string]$prefix) {
    $list = New-Object System.Collections.ArrayList
    $cb = [W+EnumProc]{
        param($h, $l)
        $cls = [W]::Cls($h)
        if ($cls.StartsWith($prefix)) {
            $r = New-Object W+RECT
            [void][W]::GetWindowRect($h, [ref]$r)
            [void]$list.Add([pscustomobject]@{
                Hwnd = $h; Class = $cls; Parent = [W]::GetParent($h)
                L = $r.L; T = $r.T; W = ($r.R - $r.L); H = ($r.B - $r.T)
                Visible = [W]::IsWindowVisible($h)
            })
        }
        return $true
    }
    [void][W]::EnumChildWindows($wordTop, $cb, [IntPtr]::Zero)
    return $list
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
# The user mode is saved and restored in finally (see e2e-smoke.ps1).
$cfgPath = Join-Path $installDir 'native-host.cfg'
$cfgBackup = $null
if (Test-Path $cfgPath) { $cfgBackup = Get-Content $cfgPath -Raw }
Set-Content -Path $cfgPath -Value "mode=native`r`ndump=0`r`nreserve=1" -Encoding Ascii
Step 'install dll + cfg (native, RESERVE)' $true ''

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
    Step 'add-in started (native+reserve)' ((Read-NewLog) -match 'reserve=1') ''

    # The ribbon state survives a Word restart and leaks between runs: if a previous
    # script left it collapsed, the "ribbon collapsed" step would be checking the
    # EXPANSION instead and would fail. Bring it to a known state.
    function Set-RibbonMinimized([bool]$want) {
        try {
            if ($word.CommandBars.GetPressedMso('MinimizeRibbon') -ne $want) {
                $word.CommandBars.ExecuteMso('MinimizeRibbon')
                Start-Sleep -Seconds 2
            }
        } catch { Write-Host "Ribbon toggle failed: $($_.Exception.Message)" }
    }
    Set-RibbonMinimized $false

    $d1 = $word.Documents.Add(); $d1.Content.Text = 'Reserve mode test'
    Start-Sleep -Seconds 3

    $activeTop = [W]::GetAncestor([IntPtr]$word.ActiveWindow.Hwnd, 2)
    [void][W]::SetForegroundWindow($activeTop)
    $word.ActiveWindow.WindowState = 1
    Start-Sleep -Seconds 3

    $panel = @(Get-ChildInfo $activeTop 'WindowsForms10' | Where-Object { $_.Parent -eq $activeTop })[0]
    $wwf = @(Get-ChildInfo $activeTop '_WwF' | Where-Object Visible)[0]
    Write-Host ("panel: T={0} B={1}; _WwF: T={2}" -f $panel.T, ($panel.T + $panel.H), $wwf.T)
    Step 'document area shifted below panel (no overlap)' ($wwf.T -ge ($panel.T + $panel.H)) "wwfTop=$($wwf.T) panelBottom=$($panel.T + $panel.H)"
    Snap 'e2e4-reserved'

    # stability: 8 seconds of idle, then check reserve still holds and no oscillation give-up
    Start-Sleep -Seconds 8
    $panel2 = @(Get-ChildInfo $activeTop 'WindowsForms10' | Where-Object { $_.Parent -eq $activeTop })[0]
    $wwf2 = @(Get-ChildInfo $activeTop '_WwF' | Where-Object Visible)[0]
    $log = Read-NewLog
    Step 'reserve stable after idle' (($wwf2.T -ge ($panel2.T + $panel2.H))) "wwfTop=$($wwf2.T) panelBottom=$($panel2.T + $panel2.H)"
    Step 'no oscillation give-up' (-not ($log -match 'Reserve mode disabled')) ''
    $shiftCount = @($log -split "`n" | Where-Object { $_ -match 'Reserve: anchor shifted' }).Count
    Write-Host "anchor shift operations so far: $shiftCount"

    # resize window: reserve must re-apply without runaway
    $word.ActiveWindow.WindowState = 0
    Start-Sleep -Seconds 1
    [void][W]::MoveWindow($activeTop, 40, 40, 1200, 700, $true)
    Start-Sleep -Seconds 3
    $panel3 = @(Get-ChildInfo $activeTop 'WindowsForms10' | Where-Object { $_.Parent -eq $activeTop })[0]
    $wwf3 = @(Get-ChildInfo $activeTop '_WwF' | Where-Object Visible)[0]
    Step 'reserve holds after resize' ($wwf3.T -ge ($panel3.T + $panel3.H)) "wwfTop=$($wwf3.T) panelBottom=$($panel3.T + $panel3.H)"
    Snap 'e2e4-after-resize'

    # ribbon collapse in reserve mode.
    # Collapse through the object model rather than a double click at (120,103):
    # that coordinate depended on whether the ribbon was expanded and where the window
    # sat, and a miss looked exactly like "the strip did not move".
    [void][W]::SetForegroundWindow($activeTop)
    Start-Sleep -Milliseconds 500
    Set-RibbonMinimized $true
    Start-Sleep -Seconds 3
    $panel4 = @(Get-ChildInfo $activeTop 'WindowsForms10' | Where-Object { $_.Parent -eq $activeTop })[0]
    $wwf4 = @(Get-ChildInfo $activeTop '_WwF' | Where-Object Visible)[0]
    Step 'reserve holds after ribbon collapse' (($panel4.T -lt $panel3.T) -and ($wwf4.T -ge ($panel4.T + $panel4.H))) "panelTop=$($panel3.T)->$($panel4.T) wwfTop=$($wwf4.T)"
    Snap 'e2e4-ribbon-collapsed'
    Set-RibbonMinimized $false
    Start-Sleep -Seconds 2

    $log2 = Read-NewLog
    $shiftTotal = @($log2 -split "`n" | Where-Object { $_ -match 'Reserve: anchor shifted' }).Count
    Write-Host "total anchor shifts: $shiftTotal"
    Step 'shift count is bounded (no fight loop)' ($shiftTotal -lt 25) "shifts=$shiftTotal"
    Step 'no oscillation give-up (final)' (-not ($log2 -match 'Reserve mode disabled')) ''

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
