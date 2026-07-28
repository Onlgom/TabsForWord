# build-release.ps1 - the full cycle: dependency check -> Release build ->
# unit tests -> installation package. Stops at the first critical error.
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File scripts\build-release.ps1
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

function Invoke-Step([string]$title, [string]$scriptName) {
    Write-Host ''
    Write-Host "=== $title ===" -ForegroundColor Cyan
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repoRoot "scripts\$scriptName")
    if ($LASTEXITCODE -ne 0) {
        Write-Host "ABORTED: step '$title' failed (exit=$LASTEXITCODE)." -ForegroundColor Red
        exit $LASTEXITCODE
    }
}

# 1. Dependencies
Write-Host '=== Checking dependencies ===' -ForegroundColor Cyan
$dotnet = Join-Path $env:LOCALAPPDATA 'Microsoft\dotnet\dotnet.exe'
if (-not (Test-Path $dotnet)) {
    try { $null = & dotnet --version 2>$null; $dotnet = 'dotnet' } catch { }
}
try { $v = & $dotnet --version 2>$null } catch { $v = $null }
if (-not $v) {
    Write-Host 'ERROR: dotnet SDK not found (see docs/ENVIRONMENT.md).' -ForegroundColor Red
    exit 1
}
Write-Host "dotnet SDK: $v"
$pia = Get-ChildItem "$env:WINDIR\assembly\GAC_MSIL\office" -Recurse -Filter 'office.dll' -ErrorAction SilentlyContinue
if (-not $pia) {
    Write-Host 'ERROR: the office.dll PIA was not found (is Microsoft Office installed?).' -ForegroundColor Red
    exit 1
}
Write-Host 'Office PIA: found'

# 2. Release build (the log is kept)
Write-Host ''
Write-Host '=== Release build ===' -ForegroundColor Cyan
$logDir = Join-Path $repoRoot 'release'
New-Item -ItemType Directory -Force $logDir | Out-Null
$buildLog = Join-Path $logDir 'build-log.txt'
$buildOutput = & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repoRoot 'scripts\build.ps1') *>&1
$buildExit = $LASTEXITCODE
# This log is kept in the repository, so the absolute path of the machine that built
# the DLL must not travel with it - the same rule that PathMap enforces inside the
# binary itself (see the .csproj and ADR in docs/DECISIONS.md).
$buildText = ($buildOutput | Out-String) -replace [regex]::Escape($repoRoot), '<repo>'
[System.IO.File]::WriteAllText($buildLog, $buildText, (New-Object System.Text.UTF8Encoding $true))
Write-Host $buildText
if ($buildExit -ne 0) {
    Write-Host "ABORTED: the build failed (log: $buildLog)." -ForegroundColor Red
    exit 1
}

# 3. Unit tests
Invoke-Step 'Unit tests' 'test.ps1'

# 4. Package
Invoke-Step 'Installation package' 'package.ps1'

Write-Host ''
Write-Host 'DONE: release/TabsForWord-Installer/ and release/TabsForWord.zip' -ForegroundColor Green
exit 0
