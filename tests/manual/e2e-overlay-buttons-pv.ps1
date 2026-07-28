# E2E phase 2: real-input clicks (SendInput), ribbon collapse, panel collapse, Protected View.
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
    [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L; public int T; public int R; public int B; }
    public static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
    public static void Click(int x, int y) {
        SetCursorPos(x, y);
        System.Threading.Thread.Sleep(150);
        mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero); // LEFTDOWN
        System.Threading.Thread.Sleep(80);
        mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero); // LEFTUP
    }
    public static void CtrlF1() {
        keybd_event(0x11, 0, 0, UIntPtr.Zero);       // CTRL down
        System.Threading.Thread.Sleep(60);
        keybd_event(0x70, 0, 0, UIntPtr.Zero);       // F1 down
        System.Threading.Thread.Sleep(60);
        keybd_event(0x70, 0, 2, UIntPtr.Zero);       // F1 up
        System.Threading.Thread.Sleep(60);
        keybd_event(0x11, 0, 2, UIntPtr.Zero);       // CTRL up
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

# ---------------------------------------------------------------- install
$p = Get-Process WINWORD -ErrorAction SilentlyContinue
if ($p) { Write-Host 'ABORT: WINWORD is running.'; exit 2 }

$installDir = Join-Path $env:LOCALAPPDATA 'TabsForWord'
Copy-Item (Join-Path $repo 'src\TabsForWord\bin\Release\TabsForWord.dll') (Join-Path $installDir 'TabsForWord.dll') -Force
# The user mode is saved and restored in finally: the script puts the add-in into
# overlay (reserve=0) while the user normally has reserve=1 - without this the strip
# would be left lying over the ruler after the run.
$cfgPath = Join-Path $installDir 'native-host.cfg'
$cfgBackup = $null
if (Test-Path $cfgPath) { $cfgBackup = Get-Content $cfgPath -Raw }
Set-Content -Path $cfgPath -Value "mode=native`r`ndump=0`r`nreserve=0" -Encoding Ascii
Step 'install dll + cfg (native, overlay)' $true ''

# Protected View test file: docx with mark-of-the-web.
# The fixture is created by the script itself: it used to rely on a test1.docx left
# by earlier runs, and after a %TEMP% cleanup the test failed during preparation.
$pvFile = Join-Path $dir 'pv-test.docx'
if (-not (Test-Path $pvFile)) {
    $seed = Join-Path $dir 'test1.docx'
    if (-not (Test-Path $seed)) {
        $seedWord = New-Object -ComObject Word.Application
        try {
            $seedWord.Visible = $false
            $seedDoc = $seedWord.Documents.Add()
            $seedDoc.Content.Text = 'TabsForWord Protected View fixture'
            $seedDoc.SaveAs([string]$seed)
            $seedDoc.Close([ref]0)
        } finally {
            $seedWord.Quit()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($seedWord)
        }
        # wait for WINWORD to exit, otherwise the main session starts on a foreign process
        $deadline = (Get-Date).AddSeconds(30)
        while ((Get-Date) -lt $deadline -and (Get-Process WINWORD -ErrorAction SilentlyContinue)) {
            Start-Sleep -Milliseconds 500
        }
    }
    Copy-Item $seed $pvFile -Force
}
Set-Content -Path $pvFile -Stream Zone.Identifier -Value "[ZoneTransfer]`r`nZoneId=3"
Step 'pv test file marked with MotW' $true ''

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

# ---------------------------------------------------------------- start word
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
    $boot = Read-NewLog
    Step 'add-in started (native mode)' (($boot -match 'Add-in started') -and ($boot -match 'Native host mode enabled')) ''

    # two more documents
    $d1 = $word.Documents.Add(); $d1.Content.Text = 'Alpha'
    Start-Sleep -Milliseconds 900
    $d2 = $word.Documents.Add(); $d2.Content.Text = 'Bravo'
    Start-Sleep -Seconds 2

    # focus + maximize active window for stable coordinates
    $activeTop = [W]::GetAncestor([IntPtr]$word.ActiveWindow.Hwnd, 2)
    [void][W]::SetForegroundWindow($activeTop)
    $word.ActiveWindow.WindowState = 1
    Start-Sleep -Seconds 2

    $surface = (Get-HostSurface $activeTop)[0]
    Step 'host on maximized active window' ($null -ne $surface -and $surface.Visible) "rect=$($surface.L),$($surface.T) $($surface.W)x$($surface.H)"

    # --- REAL click on first tab (x=80 inside first tab, y=center) ---
    $activeBefore = $word.ActiveDocument.Name
    [W]::Click($surface.L + 80, $surface.T + [int]($surface.H * 0.55))
    Start-Sleep -Seconds 2
    $activeAfter = $word.ActiveDocument.Name
    $log1 = Read-NewLog
    Step 'real click on first tab switches document' (($activeAfter -ne $activeBefore) -and ($log1 -match 'Window activated')) "before=$activeBefore after=$activeAfter"
    Snap 'e2e2-after-tab-click'

    # --- typing goes to the document (focus not stuck on panel) ---
    $selBefore = $word.Selection.Text
    [W]::keybd_event(0x58, 0, 0, [UIntPtr]::Zero)   # X down
    Start-Sleep -Milliseconds 60
    [W]::keybd_event(0x58, 0, 2, [UIntPtr]::Zero)   # X up
    Start-Sleep -Milliseconds 800
    # The check is NOT for a particular character: with a Russian keyboard layout VK "X"
    # produces another letter and comparing with "x" failed falsely. All that matters is
    $typedOk = $false
    $typedText = ''
    try { $typedText = $word.ActiveDocument.Content.Text } catch {}
    $typedOk = ($typedText -replace "[`r`n]", '').Length -gt ($selBefore -replace "[`r`n]", '').Length
    Step 'typing after tab click goes into document' $typedOk "text='$($typedText -replace "[`r`n]", '')'"

    # --- collapse panel via its button (rightmost), then expand ---
    $activeTop2 = [W]::GetAncestor([IntPtr]$word.ActiveWindow.Hwnd, 2)
    $s = (Get-HostSurface $activeTop2)[0]
    $hBefore = $s.H
    [W]::Click($s.L + $s.W - 22, $s.T + [int]($s.H * 0.55))
    Start-Sleep -Seconds 2
    $s2 = (Get-HostSurface $activeTop2)[0]
    $log2 = Read-NewLog
    Step 'collapse button shrinks panel' (($s2.H -lt $hBefore) -and ($log2 -match 'Tab panel collapsed')) "h=$hBefore -> $($s2.H)"
    Snap 'e2e2-collapsed'
    [W]::Click($s2.L + $s2.W - 22, $s2.T + [int]($s2.H * 0.6))
    Start-Sleep -Seconds 2
    $s3 = (Get-HostSurface $activeTop2)[0]
    Step 'expand button restores panel' ($s3.H -eq $hBefore) "h=$($s2.H) -> $($s3.H)"

    # --- ribbon collapse (Ctrl+F1): panel must move up, then back ---
    $tBefore = $s3.T
    [W]::CtrlF1()
    Start-Sleep -Seconds 2
    $s4 = (Get-HostSurface $activeTop2)[0]
    Step 'ribbon collapse moves panel up' ($s4.T -lt $tBefore) "t=$tBefore -> $($s4.T)"
    Snap 'e2e2-ribbon-collapsed'
    [W]::CtrlF1()
    Start-Sleep -Seconds 2
    $s5 = (Get-HostSurface $activeTop2)[0]
    Step 'ribbon expand moves panel back' ($s5.T -eq $tBefore) "t=$($s4.T) -> $($s5.T)"

    # --- Protected View: no crash, PV tab appears in strips, PV window w/o panel ---
    $pvCountBefore = $word.ProtectedViewWindows.Count
    $word.ProtectedViewWindows.Open([string]$pvFile) | Out-Null
    Start-Sleep -Seconds 4
    $pvOk = $word.ProtectedViewWindows.Count -gt $pvCountBefore
    $log3 = Read-NewLog
    Step 'protected view opened without crash' $pvOk "pvCount=$($word.ProtectedViewWindows.Count)"
    Step 'PV tab appears in model' ($log3 -match '\[PV\]') ''
    Snap 'e2e2-protected-view'
    try { $word.ProtectedViewWindows.Item(1).Close() } catch {}
    Start-Sleep -Seconds 2

    # --- close everything gracefully
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

Write-Host ''
Write-Host '=== Add-in log (this run) ==='
$tail = Read-NewLog
Write-Host $tail
$errors = @($tail -split "`n" | Where-Object { $_ -match '\[ERROR\]' })
Step 'no ERROR lines in add-in log' ($errors.Count -eq 0) "errors=$($errors.Count)"
Step 'no CTP fallback' (-not ($tail -match 'falling back')) ''

Write-Host ''
Write-Host '=== SUMMARY ==='
$results | ForEach-Object { Write-Host $_ }
$failCount = @($results | Where-Object { $_ -like '`[FAIL*' }).Count
Write-Host "FAILURES: $failCount"
exit $failCount
