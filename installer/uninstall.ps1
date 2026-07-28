# uninstall.ps1 - completely removes the "Tabs for Word" add-in.
# No administrator rights required. Word must be closed.
# Usage: double-click Uninstall.cmd, or:
#   powershell -NoProfile -ExecutionPolicy Bypass -File uninstall.ps1 [-Lang en|ru]
param([ValidateSet('auto', 'en', 'ru')][string]$Lang = 'auto')
$ErrorActionPreference = 'Stop'

if ($Lang -eq 'auto') {
    $Lang = if ((Get-UICulture).TwoLetterISOLanguageName -eq 'ru') { 'ru' } else { 'en' }
}
$script:Lang = $Lang
function T([string]$en, [string]$ru) { if ($script:Lang -eq 'ru') { return $ru } return $en }

$ClsidConnect  = '{A3B7C9D1-5E2F-4A6B-8C0D-1F2E3D4C5B6A}'
$ClsidControl  = '{D4E5F6A7-8B9C-4D0E-A1B2-C3D4E5F6A7B8}'
$ProgIdConnect = 'TabsForWord.Connect'
$ProgIdControl = 'TabsForWord.TabStripControl'

if (-not [Environment]::Is64BitProcess) {
    $sysnative = Join-Path $env:windir 'sysnative\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path $sysnative) {
        & $sysnative -NoProfile -ExecutionPolicy Bypass -File $MyInvocation.MyCommand.Path -Lang $script:Lang
        exit $LASTEXITCODE
    }
    Write-Host (T 'ERROR: 64-bit PowerShell is required.' 'ОШИБКА: требуется 64-битный PowerShell.') -ForegroundColor Red
    exit 1
}

Write-Host (T '=== Uninstalling: Tabs for Word ===' '=== Удаление: Вкладки для Word ===') -ForegroundColor Cyan

$winword = Get-Process WINWORD -ErrorAction SilentlyContinue
if ($winword) {
    Write-Host ''
    Write-Host (T 'Microsoft Word is running right now (possibly in the background - Office Startup Boost).' `
                  'Microsoft Word сейчас запущен (возможно, в фоне - Office Startup Boost).') -ForegroundColor Yellow
    Write-Host (T 'Please:' 'Пожалуйста:')
    Write-Host (T '  1. Save your documents and close every Word window.' '  1. Сохраните документы и закройте все окна Word.')
    Write-Host (T '  2. Run the uninstaller again.' '  2. Запустите удаление ещё раз.')
    Write-Host ''
    Write-Host (T 'NOTHING WAS REMOVED.' 'Удаление НЕ выполнено.') -ForegroundColor Red
    exit 1
}

# 1. Ключ надстройки Word
Remove-Item -Path "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgIdConnect" -Recurse -Force -ErrorAction SilentlyContinue

# 2. COM-регистрация (оба вида реестра: 64-битный и WOW6432Node для 32-битного Word)
foreach ($key in @(
    "HKCU:\Software\Classes\CLSID\$ClsidConnect",
    "HKCU:\Software\Classes\CLSID\$ClsidControl",
    "HKCU:\Software\Classes\WOW6432Node\CLSID\$ClsidConnect",
    "HKCU:\Software\Classes\WOW6432Node\CLSID\$ClsidControl",
    "HKCU:\Software\Classes\$ProgIdConnect",
    "HKCU:\Software\Classes\$ProgIdControl"
)) {
    Remove-Item -Path $key -Recurse -Force -ErrorAction SilentlyContinue
}

# 3. Файлы (логи и настройки удаляем вместе с папкой)
$installDir = Join-Path $env:LOCALAPPDATA 'TabsForWord'
if (Test-Path $installDir) {
    Remove-Item -Path $installDir -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ''
Write-Host (T 'UNINSTALL COMPLETE. The add-in has been removed.' `
              'УДАЛЕНИЕ ЗАВЕРШЕНО. Надстройка полностью удалена.') -ForegroundColor Green
exit 0
