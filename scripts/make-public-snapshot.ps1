# make-public-snapshot.ps1 - builds the folder that becomes the PUBLIC repository.
#
# The private repository stays the working archive: its history, its Russian
# working notes and its build artifacts are not published. This script derives a
# clean snapshot from it - tracked source only, one fresh commit, no history - so
# the two never have to be maintained by hand.
#
# What it does:
#   1. takes the list of TRACKED files (so nothing untracked or ignored sneaks in);
#   2. drops the working notes and the release/ artifacts;
#   3. resolves the private/public markers in the Markdown files (see below);
#   4. refuses to continue if anything personal survived the copy;
#   5. initialises a git repository with a single commit.
#
# It never creates a GitHub repository and never pushes: the snapshot only
# becomes public when someone adds a remote and pushes it by hand.
#
# Markers inside .md files (invisible when the file is rendered):
#   ... text ... <!--private-->        the line is dropped from the snapshot
#   <!--private-start--> ... <!--private-end-->    the block is dropped
#   <!--public:TEXT-->                 becomes TEXT in the snapshot
#
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File scripts\make-public-snapshot.ps1
param(
    [string]$OutDir = (Join-Path $env:TEMP 'TabsForWord-public'),
    [string]$AuthorName = 'Onlgom',
    [string]$AuthorEmail = 'onlgom@users.noreply.github.com',
    [switch]$NoCommit
)
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot

# Files that stay in the private archive. Everything else that git tracks is published.
$excludeFiles = @(
    'HANDOFF.md', 'PROJECT_STATE.md', 'TASKS.md', 'CONTINUE_PROMPT.md',
    'CLAUDE.md', 'FINAL_REPORT.md'
)
# Build outputs: the installer package and the ZIP belong on the Releases page,
# not in the source tree (they are rebuilt by scripts\build-release.ps1).
$excludePrefixes = @('release/')

function Get-TrackedFiles {
    Push-Location $repoRoot
    $prev = [Console]::OutputEncoding
    try {
        # quotepath=false + UTF-8 so the Cyrillic file names in installer/ survive.
        [Console]::OutputEncoding = [Text.Encoding]::UTF8
        $out = & git -c core.quotepath=false ls-files
        if ($LASTEXITCODE -ne 0) { throw 'git ls-files failed (is this a git repository?)' }
        return @($out)
    }
    finally {
        [Console]::OutputEncoding = $prev
        Pop-Location
    }
}

function Resolve-Markers([string]$text) {
    $lines = $text -split "`r?`n"
    $result = New-Object System.Collections.ArrayList
    $skipping = $false
    foreach ($line in $lines) {
        if ($line -match '<!--private-start-->') { $skipping = $true; continue }
        if ($line -match '<!--private-end-->')   { $skipping = $false; continue }
        if ($skipping) { continue }
        if ($line -match '<!--private-->') { continue }
        $m = [regex]::Match($line, '^\s*<!--public:(.*)-->\s*$')
        if ($m.Success) { [void]$result.Add($m.Groups[1].Value); continue }
        [void]$result.Add($line)
    }
    return ($result -join "`r`n")
}

# ---------------------------------------------------------------- collect
$tracked = Get-TrackedFiles
$publish = @()
foreach ($f in $tracked) {
    if ($excludeFiles -contains $f) { continue }
    $skip = $false
    foreach ($p in $excludePrefixes) { if ($f.StartsWith($p)) { $skip = $true; break } }
    if ($skip) { continue }
    $publish += $f
}
Write-Host ("Tracked: {0}; publishing: {1}; kept private: {2}" -f `
    $tracked.Count, $publish.Count, ($tracked.Count - $publish.Count))

# ---------------------------------------------------------------- copy
if (Test-Path -LiteralPath $OutDir) { Remove-Item -LiteralPath $OutDir -Recurse -Force }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

foreach ($rel in $publish) {
    $src = Join-Path $repoRoot ($rel -replace '/', '\')
    $dst = Join-Path $OutDir ($rel -replace '/', '\')
    $dstDir = Split-Path -Parent $dst
    if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Force -Path $dstDir | Out-Null }

    if ($rel -like '*.md') {
        # Markdown carries the markers; keep the BOM the original had.
        $bytes = [System.IO.File]::ReadAllBytes($src)
        $hadBom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
        $text = [System.IO.File]::ReadAllText($src, [Text.Encoding]::UTF8)
        $text = Resolve-Markers $text
        [System.IO.File]::WriteAllText($dst, $text, (New-Object System.Text.UTF8Encoding $hadBom))
    }
    else {
        Copy-Item -LiteralPath $src -Destination $dst -Force
    }
}

# .gitignore is the one file whose content genuinely differs: the private archive
# commits the built package at checkpoints, the published repository never does.
# It cannot carry the HTML markers (git would read them as patterns), so the tail
# is replaced here.
$ignorePath = Join-Path $OutDir '.gitignore'
if (Test-Path -LiteralPath $ignorePath) {
    $ignore = [System.IO.File]::ReadAllText($ignorePath, [Text.Encoding]::UTF8)
    $cut = $ignore.IndexOf('# Release archives')
    if ($cut -ge 0) { $ignore = $ignore.Substring(0, $cut) }
    $ignore = $ignore.TrimEnd() + "`r`n`r`n" +
        "# Build artifacts are not published: release/ is produced by`r`n" +
        "# scripts\build-release.ps1, and the ready-to-install ZIP lives on the`r`n" +
        "# Releases page of this repository.`r`n" +
        "release/`r`n"
    [System.IO.File]::WriteAllText($ignorePath, $ignore, (New-Object System.Text.UTF8Encoding $false))
}

# ---------------------------------------------------------------- verify
# The snapshot is about to become public, so the check is a hard gate rather than
# a warning: absolute paths, the account name and the machine name must not be in it.
$badPatterns = @(
    @{ Name = 'absolute Windows path'; Pattern = '[A-Za-z]:\\Users\\' },
    @{ Name = 'account name';          Pattern = [regex]::Escape($env:USERNAME) },
    @{ Name = 'computer name';         Pattern = [regex]::Escape($env:COMPUTERNAME) },
    @{ Name = 'machine name pattern';  Pattern = '(DESKTOP|LAPTOP)-[A-Z0-9]{7}' }
)
$problems = @()
Get-ChildItem -LiteralPath $OutDir -Recurse -File | ForEach-Object {
    if ($_.Extension -in @('.dll', '.zip', '.png', '.jpg', '.pdf', '.ico')) { return }
    $content = [System.IO.File]::ReadAllText($_.FullName, [Text.Encoding]::UTF8)
    foreach ($b in $badPatterns) {
        foreach ($m in [regex]::Matches($content, $b.Pattern, 'IgnoreCase')) {
            $rel = $_.FullName.Substring($OutDir.Length + 1)
            # A blog URL in RESEARCH.md happens to contain the author's login as a
            # substring; the same goes for the method name CollapsedExpandRect.
            $line = ($content -split "`r?`n" | Where-Object { $_ -match [regex]::Escape($m.Value) } | Select-Object -First 1)
            if ($line -match 'learn\.microsoft\.com' -or $line -match 'CollapsedExpandRect') { continue }
            $problems += ("{0}: {1} ({2})" -f $rel, $b.Name, $m.Value)
        }
    }
}
$problems = @($problems | Select-Object -Unique)
if ($problems.Count -gt 0) {
    Write-Host ''
    Write-Host 'STOPPED: the snapshot still contains personal data:' -ForegroundColor Red
    $problems | ForEach-Object { Write-Host ("  " + $_) -ForegroundColor Red }
    exit 1
}
Write-Host 'Check: no absolute paths, no account or machine names.' -ForegroundColor Green

foreach ($f in $excludeFiles) {
    if (Test-Path -LiteralPath (Join-Path $OutDir $f)) {
        Write-Host "STOPPED: $f must not be in the snapshot." -ForegroundColor Red
        exit 1
    }
}
if (Test-Path -LiteralPath (Join-Path $OutDir 'release')) {
    Write-Host 'STOPPED: release/ must not be in the snapshot.' -ForegroundColor Red
    exit 1
}

# ---------------------------------------------------------------- git
if ($NoCommit) {
    Write-Host ''
    Write-Host "Snapshot ready (not committed): $OutDir"
    exit 0
}

Push-Location $OutDir
try {
    & git init -q
    & git config user.name $AuthorName
    & git config user.email $AuthorEmail
    & git add -A
    & git commit -q -m 'Tabs for Microsoft Word: a tab bar inside the Word window'
    if ($LASTEXITCODE -ne 0) { throw 'the initial commit failed' }
    $count = (& git ls-files).Count
}
finally { Pop-Location }

Write-Host ''
Write-Host "Snapshot ready: $OutDir  ($count files, one commit)" -ForegroundColor Green
Write-Host 'Next: create an EMPTY public repository on GitHub (no README, no licence), then:'
Write-Host "  cd `"$OutDir`""
Write-Host '  git remote add origin https://github.com/<owner>/<name>.git'
Write-Host '  git branch -M main'
Write-Host '  git push -u origin main'
Write-Host 'Then attach release\TabsForWord.zip to a GitHub Release (the README links there).'
exit 0
