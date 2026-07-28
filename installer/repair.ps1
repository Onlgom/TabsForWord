# repair.ps1 - a fix-up for when the tab bar stops appearing: resets
# LoadBehavior to 3 and clears Word's ban on OUR add-in (entries that belong to
# other add-ins are left alone). Run with Word closed.
# Usage: double-click Repair.cmd, or:
#   powershell -NoProfile -ExecutionPolicy Bypass -File repair.ps1 [-Lang en|ru]
param([ValidateSet('auto', 'en', 'ru')][string]$Lang = 'auto')
$ErrorActionPreference = 'Stop'

if ($Lang -eq 'auto') {
    $Lang = if ((Get-UICulture).TwoLetterISOLanguageName -eq 'ru') { 'ru' } else { 'en' }
}
$script:Lang = $Lang
function T([string]$en, [string]$ru) { if ($script:Lang -eq 'ru') { return $ru } return $en }

$ProgIdConnect = 'TabsForWord.Connect'

Write-Host (T '=== Repairing: Tabs for Word ===' '=== Починка: Вкладки для Word ===') -ForegroundColor Cyan

$winword = Get-Process WINWORD -ErrorAction SilentlyContinue
if ($winword) {
    Write-Host ''
    Write-Host (T 'Microsoft Word is running. Save your documents, close Word' `
                  'Microsoft Word сейчас запущен. Сохраните документы, закройте Word') -ForegroundColor Yellow
    Write-Host (T 'and run the repair again.' 'и запустите починку ещё раз.')
    Write-Host (T 'NOTHING WAS REPAIRED.' 'Починка НЕ выполнена.') -ForegroundColor Red
    exit 1
}

$addinKey = "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgIdConnect"
if (Test-Path $addinKey) {
    $lb = (Get-ItemProperty -Path $addinKey -Name LoadBehavior -ErrorAction SilentlyContinue).LoadBehavior
    if ($lb -ne 3) {
        Write-Host ((T 'LoadBehavior was {0} (Word disabled the add-in after a failure) - setting it back to 3.' `
                       'LoadBehavior был {0} (Word отключил надстройку после сбоя) - возвращаю 3.') -f $lb)
    }
    Set-ItemProperty -Path $addinKey -Name 'LoadBehavior' -Value 3 -Type DWord
} else {
    Write-Host (T 'The add-in registry key was not found - run Install.cmd first.' `
                  'Ключ надстройки не найден - запустите install.ps1.') -ForegroundColor Yellow
    exit 1
}

# Записи DisabledItems двоичные; внутри в UTF-16 лежат путь DLL и имя надстройки.
# Удаляем только записи с упоминанием TabsForWord: чужие надстройки не включаем.
$unbanned = $false
foreach ($officeVer in @('14.0', '15.0', '16.0')) {
    $resiliency = "HKCU:\Software\Microsoft\Office\$officeVer\Word\Resiliency\DisabledItems"
    if (-not (Test-Path $resiliency)) { continue }
    $key = Get-Item $resiliency
    try {
        foreach ($valueName in $key.GetValueNames()) {
            if (-not $valueName) { continue }
            $bytes = $key.GetValue($valueName)
            if ($bytes -is [byte[]] -and
                [System.Text.Encoding]::Unicode.GetString($bytes) -match 'TabsForWord') {
                Remove-ItemProperty -Path $resiliency -Name $valueName -ErrorAction SilentlyContinue
                Write-Host ((T 'Cleared the Word ban on our add-in ({0}, entry {1}).' `
                               'Снят бан Word с нашей надстройки ({0}, запись {1}).') -f $officeVer, $valueName)
                $unbanned = $true
            }
        }
    } finally {
        $key.Close()
    }
}
if (-not $unbanned) {
    Write-Host (T 'Our add-in is not in Word''s list of disabled items.' `
                  'В списке отключённых элементов Word нашей надстройки нет.')
}

Write-Host (T 'DONE. Start Word and check the tab bar.' `
              'ГОТОВО. Запустите Word и проверьте панель вкладок.') -ForegroundColor Green
exit 0
