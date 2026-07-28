# test.ps1 - builds and runs the unit tests (no Word involved).
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File scripts\test.ps1
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
$testProj = Join-Path $repoRoot 'tests\TabsForWord.Tests\TabsForWord.Tests.csproj'

$dotnet = Find-DotNet
if (-not $dotnet) {
    Write-Host 'ERROR: dotnet SDK not found.' -ForegroundColor Red
    exit 1
}

Write-Host 'Building the tests...'
& $dotnet build $testProj -c Release -v minimal --nologo
if ($LASTEXITCODE -ne 0) {
    Write-Host 'ERROR: the test build failed.' -ForegroundColor Red
    exit 1
}

$exe = Join-Path $repoRoot 'tests\TabsForWord.Tests\bin\Release\TabsForWord.Tests.exe'
Write-Host 'Running the tests...'
& $exe
$code = $LASTEXITCODE
if ($code -ne 0) {
    Write-Host "TESTS FAILED (exit=$code)." -ForegroundColor Red
} else {
    Write-Host 'ALL TESTS PASSED.' -ForegroundColor Green
}
exit $code
