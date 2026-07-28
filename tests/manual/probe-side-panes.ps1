# Probe: how Word lays out side task panes (Navigation/Editor) vs _WwF anchor.
# Prints direct children of OpusApp with class+rect before/after opening panes.
# ASCII only. PowerShell 5.1.
$ErrorActionPreference = 'Stop'
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
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L; public int T; public int R; public int B; }
    public static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
}
"@
[W]::SetProcessDpiAwareness(2) | Out-Null

function Dump-Children([IntPtr]$top, [string]$label) {
    Write-Host "===== $label ====="
    $list = New-Object System.Collections.ArrayList
    $cb = [W+EnumProc]{
        param($h, $l)
        $r = New-Object W+RECT
        [void][W]::GetWindowRect($h, [ref]$r)
        $vis = [W]::IsWindowVisible($h)
        if ($vis -and ($r.B - $r.T) -gt 20 -and ($r.R - $r.L) -gt 20) {
            [void]$list.Add([pscustomobject]@{
                Hwnd = $h; Parent = [W]::GetParent($h); Cls = [W]::Cls($h)
                L=$r.L; T=$r.T; W=($r.R-$r.L); H=($r.B-$r.T)
            })
        }
        return $true
    }
    [void][W]::EnumChildWindows($top, $cb, [IntPtr]::Zero)
    # direct children of OpusApp plus children of _WwF (level 2)
    $direct = @($list | Where-Object { $_.Parent -eq $top })
    foreach ($c in $direct | Sort-Object T, L) {
        Write-Host ("  {0,-28} {1},{2} {3}x{4}" -f $c.Cls, $c.L, $c.T, $c.W, $c.H)
        if ($c.Cls -eq '_WwF') {
            foreach ($g in ($list | Where-Object { $_.Parent -eq $c.Hwnd } | Sort-Object T, L)) {
                Write-Host ("      -> {0,-24} {1},{2} {3}x{4}" -f $g.Cls, $g.L, $g.T, $g.W, $g.H)
            }
        }
    }
}

$p = Get-Process WINWORD -ErrorAction SilentlyContinue
if ($p) { Write-Host 'ABORT: WINWORD is running.'; exit 2 }

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
    $word.ActiveWindow.WindowState = 1
    Start-Sleep -Seconds 2

    Dump-Children $top 'BASELINE (no panes)'

    # Navigation (the left pane) - a reliable programmatic way
    $word.ActiveWindow.DocumentMap = $true
    Start-Sleep -Seconds 3
    Dump-Children $top 'NAVIGATION PANE (left)'
    $word.ActiveWindow.DocumentMap = $false
    Start-Sleep -Seconds 2

    # Editor (the right pane) - trying idMso candidates
    $opened = $null
    foreach ($id in @('EditorTool', 'Editor', 'ProofingTool', 'SpellingAndGrammar')) {
        try {
            $word.CommandBars.ExecuteMso($id)
            $opened = $id
            break
        } catch { }
    }
    Start-Sleep -Seconds 4
    Write-Host "Editor idMso opened: $opened"
    Dump-Children $top 'EDITOR PANE (right)'

    foreach ($doc in @($word.Documents)) { $doc.Close([ref]0) }
    Start-Sleep -Seconds 2
    $word.Quit()
    Start-Sleep -Seconds 3
} catch {
    Write-Host "EXCEPTION: $($_.Exception.Message)"
    try { if ($word) { $word.Quit() } } catch {}
} finally {
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch {}
}
exit 0
