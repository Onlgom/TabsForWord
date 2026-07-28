# build.ps1 - builds TabsForWord (Release).
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File scripts\build.ps1
$ErrorActionPreference = 'Stop'

function Find-DotNet {
    $candidates = @(
        (Join-Path $env:LOCALAPPDATA 'Microsoft\dotnet\dotnet.exe'),
        'dotnet'
    )
    foreach ($c in $candidates) {
        try {
            $null = & $c --version 2>$null
            if ($LASTEXITCODE -eq 0) { return $c }
        } catch { }
    }
    return $null
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$proj = Join-Path $repoRoot 'src\TabsForWord\TabsForWord.csproj'

$dotnet = Find-DotNet
if (-not $dotnet) {
    Write-Host 'ERROR: dotnet SDK not found. Install .NET SDK 8 (see docs/ENVIRONMENT.md).' -ForegroundColor Red
    exit 1
}

Write-Host "Building (Release): $proj"
& $dotnet build $proj -c Release -v minimal --nologo
if ($LASTEXITCODE -ne 0) {
    Write-Host 'ERROR: the build failed.' -ForegroundColor Red
    exit 1
}

$dll = Join-Path $repoRoot 'src\TabsForWord\bin\Release\TabsForWord.dll'
if (-not (Test-Path $dll)) {
    Write-Host "ERROR: build output not found: $dll" -ForegroundColor Red
    exit 1
}
Write-Host "Done: $dll" -ForegroundColor Green
exit 0
