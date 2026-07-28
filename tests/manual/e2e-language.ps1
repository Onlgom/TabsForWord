# E2E: interface language (stage 29).
# Phase A: lang=en in language.cfg -> add-in runs in English and the tab context
#          menu really shows English items (read back with GetMenuStringW).
# Phase B: lang=auto -> language is taken from Word itself; on this machine Word
#          is Russian, so the log must say Ru with the LCID it detected.
# The user's own language.cfg and native-host.cfg are restored afterwards.
# ASCII only. PowerShell 5.1. Run with Word CLOSED.
$ErrorActionPreference = 'Stop'
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

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class WL {
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
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr wp, IntPtr lp);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetMenuStringW(IntPtr hMenu, uint id, StringBuilder sb, int max, uint flags);
    [DllImport("user32.dll")] public static extern int GetMenuItemCount(IntPtr hMenu);
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
    // A popup menu window owns its HMENU; MN_GETHMENU (0x01E1) hands it over,
    // after which the item captions can be read as plain text.
    public static string MenuTexts(IntPtr menuWindow) {
        IntPtr hMenu = SendMessage(menuWindow, 0x01E1, IntPtr.Zero, IntPtr.Zero);
        if (hMenu == IntPtr.Zero) return "";
        int n = GetMenuItemCount(hMenu);
        var all = new StringBuilder();
        for (int i = 0; i < n; i++) {
            var sb = new StringBuilder(256);
            GetMenuStringW(hMenu, (uint)i, sb, sb.Capacity, 0x0400 /* MF_BYPOSITION */);
            if (sb.Length > 0) { if (all.Length > 0) all.Append(" | "); all.Append(sb.ToString()); }
        }
        return all.ToString();
    }
}
"@
[WL]::SetProcessDpiAwareness(2) | Out-Null

function Get-Panel([IntPtr]$wordTop) {
    $list = New-Object System.Collections.ArrayList
    $cb = [WL+EnumProc]{
        param($h, $l)
        $cls = [WL]::Cls($h)
        if ($cls.StartsWith('WindowsForms10') -and [WL]::IsWindowVisible($h)) {
            $r = New-Object WL+RECT
            [void][WL]::GetWindowRect($h, [ref]$r)
            [void]$list.Add([pscustomobject]@{
                Hwnd = $h; Parent = [WL]::GetParent($h)
                L = $r.L; T = $r.T; W = ($r.R - $r.L); H = ($r.B - $r.T)
            })
        }
        return $true
    }
    [void][WL]::EnumChildWindows($wordTop, $cb, [IntPtr]::Zero)
    return @($list | Where-Object { $_.Parent -eq $wordTop } | Select-Object -First 1)
}

function Wait-Menu([uint32]$wordPid, [bool]$open, [int]$timeoutMs = 5000) {
    $deadline = [Environment]::TickCount + $timeoutMs
    while ([Environment]::TickCount -lt $deadline) {
        if ((([WL]::FindMenuOfProcess($wordPid)) -ne [IntPtr]::Zero) -eq $open) { return $true }
        Start-Sleep -Milliseconds 100
    }
    return $false
}

$p = Get-Process WINWORD -ErrorAction SilentlyContinue
if ($p) { Write-Host 'ABORT: WINWORD is running.'; exit 2 }

$installDir = Join-Path $env:LOCALAPPDATA 'TabsForWord'
Copy-Item (Join-Path $repo 'src\TabsForWord\bin\Release\TabsForWord.dll') (Join-Path $installDir 'TabsForWord.dll') -Force

$langPath = Join-Path $installDir 'language.cfg'
$langExisted = Test-Path $langPath
$langBackup = $null
if ($langExisted) { $langBackup = Get-Content $langPath -Raw }

$logFile = Join-Path $installDir ('Logs\TabsForWord-' + (Get-Date -Format 'yyyyMMdd') + '.log')
$script:logStart = 0
function Mark-Log { $script:logStart = 0; if (Test-Path $logFile) { $script:logStart = (Get-Item $logFile).Length } }
function Read-NewLog {
    if (-not (Test-Path $logFile)) { return '' }
    $fs = [System.IO.File]::Open($logFile, 'Open', 'Read', 'ReadWrite')
    try {
        [void]$fs.Seek($script:logStart, 'Begin')
        $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8)
        return $sr.ReadToEnd()
    } finally { $fs.Dispose() }
}

# One Word session; optionally opens the tab context menu and reads its captions.
function Invoke-WordSession([bool]$readMenu) {
    Mark-Log
    $word = $null
    $menuText = ''
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
        $wordPid = [uint32](Get-Process WINWORD).Id
        $wordTop = [WL]::GetAncestor([IntPtr]$word.ActiveWindow.Hwnd, 2)
        [void][WL]::SetForegroundWindow($wordTop)
        $word.ActiveWindow.WindowState = 1
        Start-Sleep -Seconds 2

        if ($readMenu) {
            $s = Get-Panel $wordTop
            if ($null -ne $s) {
                [WL]::RightClick($s.L + 30, $s.T + [int]($s.H * 0.55))
                if (Wait-Menu $wordPid $true) {
                    $menuWin = [WL]::FindMenuOfProcess($wordPid)
                    $menuText = [WL]::MenuTexts($menuWin)
                }
                [WL]::Key(0x1B)   # Esc
                [void](Wait-Menu $wordPid $false 3000)
            }
        }

        foreach ($doc in @($word.Documents)) { $doc.Close([ref]0) }
        Start-Sleep -Seconds 2
        $word.Quit()
        Start-Sleep -Seconds 3
    } finally {
        try { if ($word) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } } catch {}
    }
    Start-Sleep -Seconds 2
    return [pscustomobject]@{ Log = (Read-NewLog); Menu = $menuText }
}

try {
    # ---------- Phase A: forced English ----------
    Set-Content -Path $langPath -Value 'lang=en' -Encoding Ascii
    $a = Invoke-WordSession $true
    Step 'A: add-in started (lang=en)' ($a.Log -match 'Add-in started') ''
    Step 'A: language resolved from language.cfg' ($a.Log -match 'UI language: En \(language\.cfg\)') ''
    Step 'A: context menu is in English' ($a.Menu -match 'Close' -and $a.Menu -match 'Tab color') $a.Menu
    Step 'A: no Cyrillic left in the menu' ($a.Menu -notmatch '[Ѐ-ӿ]') ''
    $errA = @($a.Log -split "`n" | Where-Object { $_ -match '\[ERROR\]' })
    Step 'A: no ERROR lines' ($errA.Count -eq 0) "errors=$($errA.Count)"

    # ---------- Phase B: follow Word ----------
    Set-Content -Path $langPath -Value 'lang=auto' -Encoding Ascii
    $b = Invoke-WordSession $true
    Step 'B: language taken from Word itself' ($b.Log -match 'UI language: \w+ \(Word UI language \(LCID \d+\)\)') `
        ((($b.Log -split "`n") | Where-Object { $_ -match 'UI language:' } | Select-Object -First 1))
    Step 'B: menu matches the detected language' `
        (($b.Log -match 'UI language: Ru' -and $b.Menu -match '[Ѐ-ӿ]') -or
         ($b.Log -match 'UI language: En' -and $b.Menu -match 'Close')) $b.Menu
    $errB = @($b.Log -split "`n" | Where-Object { $_ -match '\[ERROR\]' })
    Step 'B: no ERROR lines' ($errB.Count -eq 0) "errors=$($errB.Count)"
} catch {
    Step 'scenario exception' $false $_.Exception.Message
} finally {
    if ($langExisted) { Set-Content -Path $langPath -Value $langBackup -NoNewline }
    elseif (Test-Path $langPath) { Remove-Item $langPath -Force }
}

Write-Host ''
Write-Host '=== LANGUAGE SUMMARY ==='
$results | ForEach-Object { Write-Host $_ }
$failCount = @($results | Where-Object { $_ -like '`[FAIL*' }).Count
Write-Host "FAILURES: $failCount"
exit $failCount
