# diagnose.ps1 - collects diagnostics for the "Tabs for Word" add-in.
# Changes nothing on the computer: it only reads and assembles a report.
# Result: a ZIP file on the desktop, to be sent to the developer.
# No administrator rights required. Works in both 32- and 64-bit PowerShell.
#
# Язык отчёта: -Lang, иначе по языку интерфейса Windows. Отчёт и то, что видно
# на экране, всегда на одном языке: человек должен иметь возможность прочитать
# то, что он отправляет (см. раздел о приватности в INSTALL_RU.md). Для любой
# нерусской локали это английский, поэтому отчёты «из мира» приходят читаемыми.
# Строки самого лога надстройки английские всегда - их этот скрипт и разбирает.
param([ValidateSet('auto', 'en', 'ru')][string]$Lang = 'auto')
$ErrorActionPreference = 'Continue'

if ($Lang -eq 'auto') {
    $Lang = if ((Get-UICulture).TwoLetterISOLanguageName -eq 'ru') { 'ru' } else { 'en' }
}
$script:Lang = $Lang
function T([string]$en, [string]$ru) { if ($script:Lang -eq 'ru') { return $ru } return $en }

$ClsidConnect  = '{A3B7C9D1-5E2F-4A6B-8C0D-1F2E3D4C5B6A}'
$ClsidControl  = '{D4E5F6A7-8B9C-4D0E-A1B2-C3D4E5F6A7B8}'
$ProgIdConnect = 'TabsForWord.Connect'

# Страница загрузки .NET Framework 4.8 (ДОЛЖНО СОВПАДАТЬ с install.ps1)
$NetDownloadUrl = T 'https://dotnet.microsoft.com/download/dotnet-framework/net48' `
                    'https://dotnet.microsoft.com/ru-ru/download/dotnet-framework/net48'

$report = New-Object System.Text.StringBuilder
$problems = New-Object System.Collections.ArrayList
function W([string]$s) { [void]$report.AppendLine($s); Write-Host $s }
function Section([string]$name) { W ''; W ('=' * 70); W ("== $name"); W ('=' * 70) }
function Problem([string]$s) { [void]$problems.Add($s) }

function Get-ExeBitness([string]$path) {
    try {
        $fs = [System.IO.File]::Open($path, 'Open', 'Read', 'ReadWrite')
        try {
            $br = New-Object System.IO.BinaryReader($fs)
            [void]$fs.Seek(0x3C, 'Begin')
            $peOffset = $br.ReadInt32()
            [void]$fs.Seek($peOffset + 4, 'Begin')
            switch ($br.ReadUInt16()) {
                0x8664 { return 'x64' }
                0x014C { return 'x86' }
                0xAA64 { return 'ARM64' }
                default { return 'unknown' }
            }
        } finally { $fs.Dispose() }
    } catch { return "error: $($_.Exception.Message)" }
}

# Чтение реестра в явном виде (64/32) - работает из PowerShell любой битности
function Open-RegKey([string]$hive, [string]$view, [string]$subKey) {
    try {
        $h = [Microsoft.Win32.RegistryHive]::$hive
        $v = [Microsoft.Win32.RegistryView]::$view
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($h, $v)
        return $base.OpenSubKey($subKey)
    } catch { return $null }
}
function Dump-RegKey([string]$hive, [string]$view, [string]$subKey) {
    $k = Open-RegKey $hive $view $subKey
    if (-not $k) { W ("  [$hive/$view] $subKey : " + (T 'NO SUCH KEY' 'НЕТ КЛЮЧА')); return $false }
    W "  [$hive/$view] $subKey"
    foreach ($name in $k.GetValueNames()) {
        $shown = if ($name) { $name } else { '(default)' }
        $val = $k.GetValue($name)
        if ($val -is [byte[]]) { $val = (($val | ForEach-Object { $_.ToString('X2') }) -join ' ') }
        W ("    {0} = {1}" -f $shown, $val)
    }
    foreach ($sub in $k.GetSubKeyNames()) { W ((T '    <subkey> {0}' '    <подраздел> {0}') -f $sub) }
    $k.Close()
    return $true
}

$ts = Get-Date -Format 'yyyyMMdd-HHmmss'
W (T 'TabsForWord - diagnostic report' 'TabsForWord - диагностический отчёт')
W ((T 'Date: {0}, computer: {1}, user: {2}' 'Дата: {0}, компьютер: {1}, пользователь: {2}') -f `
    (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $env:COMPUTERNAME, $env:USERNAME)
W ((T 'Diagnostics script version: {0}' 'Версия скрипта диагностики: {0}') -f '1.6')

# ---------------------------------------------------------------- система
Section (T 'System' 'Система')
try {
    $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    # The registry ProductName still says "Windows 10" on Windows 11; the build number
    # tells them apart (Windows 11 starts at build 22000).
    $productName = [string]$cv.ProductName
    try { if ([int]$cv.CurrentBuild -ge 22000) { $productName = $productName -replace 'Windows 10', 'Windows 11' } } catch { }
    W ((T 'Windows: {0} {1} (build {2}.{3})' 'Windows: {0} {1} (сборка {2}.{3})') -f `
        $productName, $cv.DisplayVersion, $cv.CurrentBuild, $cv.UBR)
} catch {
    W ((T 'Windows: could not read the version: {0}' 'Windows: ошибка чтения версии: {0}') -f $_.Exception.Message)
}
W ((T '64-bit OS: {0}; PowerShell: {1} ({2}-bit process)' '64-битная ОС: {0}; PowerShell: {1} ({2}-бит процесс)') -f `
    [Environment]::Is64BitOperatingSystem, $PSVersionTable.PSVersion, $(if ([Environment]::Is64BitProcess) { 64 } else { 32 }))
W ((T 'System locale: {0}; UI language: {1}' 'Язык системы: {0}; язык пользователя: {1}') -f `
    (Get-Culture).Name, (Get-UICulture).Name)

# .NET Framework (порог 528040 = 4.8; должен совпадать с install.ps1)
$netRelease = 0
$netVersion = $null
foreach ($view in @('Registry64', 'Registry32')) {
    $k = Open-RegKey 'LocalMachine' $view 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full'
    if ($k) {
        $r = $k.GetValue('Release')
        $v = $k.GetValue('Version')
        W ((T '.NET Framework 4.x [{0}]: version {1} (Release={2})' `
               '.NET Framework 4.x [{0}]: версия {1} (Release={2})') -f $view, $v, $r)
        if ($r -gt $netRelease) { $netRelease = $r; $netVersion = $v }
        $k.Close()
    } else {
        W ((T '.NET Framework 4.x [{0}]: key not found' '.NET Framework 4.x [{0}]: ключ не найден') -f $view)
    }
}
if ($netRelease -eq 0) {
    Problem ((T '.NET Framework 4 was not found - the add-in cannot load. Install .NET Framework 4.8: {0} (administrator rights are needed for that; it is a Microsoft requirement, the add-in itself does not need them).' `
                '.NET Framework 4 не обнаружен - надстройка не сможет загрузиться. Установите .NET Framework 4.8: {0} (нужны права администратора; это требование Microsoft, самой надстройке они не нужны).') -f $NetDownloadUrl)
} elseif ($netRelease -lt 528040) {
    $shown = if ($netVersion) { $netVersion } else { "Release=$netRelease" }
    Problem ((T '.NET Framework {0} is not enough, the add-in needs 4.8 or newer. Install it: {1} (administrator rights are needed for that; it is a Microsoft requirement, the add-in itself does not need them).' `
                '.NET Framework {0} - недостаточно, надстройке нужен 4.8 или новее. Установите: {1} (нужны права администратора; это требование Microsoft, самой надстройке они не нужны).') -f $shown, $NetDownloadUrl)
}

# ---------------------------------------------------------------- Word
Section 'Microsoft Word'
$wordPath = $null
foreach ($view in @('Registry64', 'Registry32')) {
    foreach ($hive in @('LocalMachine', 'CurrentUser')) {
        $k = Open-RegKey $hive $view 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe'
        if ($k) {
            $p = $k.GetValue('')
            W "App Paths [$hive/$view]: $p"
            if ($p -and (Test-Path $p) -and -not $wordPath) { $wordPath = $p }
            $k.Close()
        }
    }
}
# Если App Paths пуст - это ещё не значит, что Word нет: проверяем стандартные
# папки Office и COM-регистрацию Word.Application (как в install.ps1).
if (-not $wordPath) {
    W (T 'winword.exe was NOT found in App Paths - checking the standard Office folders.' `
          'winword.exe НЕ найден в App Paths - проверяем стандартные папки Office.')
    foreach ($progFiles in @($env:ProgramFiles, ${env:ProgramFiles(x86)})) {
        if (-not $progFiles -or $wordPath) { continue }
        foreach ($rel in @(
            'Microsoft Office\root\Office16\WINWORD.EXE',
            'Microsoft Office\Office16\WINWORD.EXE',
            'Microsoft Office\root\Office15\WINWORD.EXE',
            'Microsoft Office\Office15\WINWORD.EXE',
            'Microsoft Office\Office14\WINWORD.EXE'
        )) {
            $candidate = Join-Path $progFiles $rel
            if (Test-Path $candidate) {
                $wordPath = $candidate
                W ((T 'Found on disk: {0}' 'Найден на диске: {0}') -f $candidate)
                break
            }
        }
    }
}
$wordProgId = $false
foreach ($view in @('Registry64', 'Registry32')) {
    foreach ($hive in @('LocalMachine', 'CurrentUser')) {
        $k = Open-RegKey $hive $view 'SOFTWARE\Classes\Word.Application'
        if ($k) { $wordProgId = $true; $k.Close() }
    }
}
W ((T 'Word.Application COM registration: {0}' 'COM-регистрация Word.Application: {0}') -f `
    $(if ($wordProgId) { T 'present' 'есть' } else { T 'MISSING' 'НЕТ' }))

$wordBitness = $null
if ($wordPath) {
    $wordBitness = Get-ExeBitness $wordPath
    $fv = $null
    try { $fv = (Get-Item $wordPath).VersionInfo.FileVersion } catch { }
    W "winword.exe: $wordPath"
    W ((T 'Word bitness: {0}; file version: {1}' 'Битность Word: {0}; файловая версия: {1}') -f $wordBitness, $fv)
} elseif ($wordProgId) {
    W (T 'winword.exe was not found, but Word is COM-registered (non-standard Office layout).' `
          'winword.exe не найден, но COM-регистрация Word есть (нестандартная установка Office).')
} else {
    W (T 'Microsoft Word was NOT FOUND by any of the three methods.' `
          'Microsoft Word НЕ НАЙДЕН ни одним из трёх способов.')
    Problem (T ('Desktop Microsoft Word was not found on this computer (checked App Paths, the Office ' +
                'folders and the Word.Application COM registration). The add-in only works with desktop ' +
                'Word for Windows - not with Word Online, Word for macOS or Word on mobile devices.') `
               ('Настольный Microsoft Word на компьютере не обнаружен (проверены App Paths, ' +
                'папки Office и COM-регистрация Word.Application). Надстройка работает только ' +
                'с настольным Word для Windows - не с Word Online, Word для macOS и мобильным Word.'))
}
foreach ($view in @('Registry64', 'Registry32')) {
    $k = Open-RegKey 'LocalMachine' $view 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    if ($k) {
        W ((T 'Click-to-Run [{0}]: version {1}, platform {2}, channel {3}' `
               'Click-to-Run [{0}]: версия {1}, платформа {2}, канал {3}') -f `
            $view, $k.GetValue('VersionToReport'), $k.GetValue('Platform'), $k.GetValue('UpdateChannel'))
        W ((T 'Click-to-Run [{0}]: products {1}' 'Click-to-Run [{0}]: продукты {1}') -f $view, $k.GetValue('ProductReleaseIds'))
        $k.Close()
    }
}
foreach ($officeVer in @('16.0', '15.0', '14.0')) {
    foreach ($view in @('Registry64', 'Registry32')) {
        $k = Open-RegKey 'LocalMachine' $view "SOFTWARE\Microsoft\Office\$officeVer\Word\InstallRoot"
        if ($k) { W ("MSI Office {0} [{1}]: {2}" -f $officeVer, $view, $k.GetValue('Path')); $k.Close() }
    }
}
$wwProc = Get-Process WINWORD -ErrorAction SilentlyContinue
W ((T 'WINWORD.EXE running right now: {0}' 'WINWORD.EXE сейчас запущен: {0}') -f `
    $(if ($wwProc) { (T 'yes (PID ' 'да (PID ') + (($wwProc | ForEach-Object Id) -join ',') + ')' } else { T 'no' 'нет' }))

# ---------------------------------------------------------------- наши файлы
Section (T 'Add-in files' 'Файлы надстройки')
$installDir = Join-Path $env:LOCALAPPDATA 'TabsForWord'
$dllPath = Join-Path $installDir 'TabsForWord.dll'
if (Test-Path $installDir) {
    Get-ChildItem $installDir -File -ErrorAction SilentlyContinue | ForEach-Object {
        W ((T '  {0}  {1,10} bytes  {2:yyyy-MM-dd HH:mm}' '  {0}  {1,10} байт  {2:yyyy-MM-dd HH:mm}') -f `
            $_.Name, $_.Length, $_.LastWriteTime)
    }
    if (Test-Path $dllPath) {
        try {
            $vi = (Get-Item $dllPath).VersionInfo
            $asm = [System.Reflection.AssemblyName]::GetAssemblyName($dllPath)
            W ((T 'DLL: file version {0}; assembly {1}' 'DLL: файловая версия {0}; сборка {1}') -f $vi.FileVersion, $asm.FullName)
        } catch {
            W ((T 'DLL: could not read the version / assembly name: {0}' `
                   'DLL: ошибка чтения версии/имени сборки: {0}') -f $_.Exception.Message)
        }
        try { W ("DLL: SHA256 {0}" -f (Get-FileHash $dllPath -Algorithm SHA256).Hash) } catch { }
        $zone = Get-Item -Path $dllPath -Stream Zone.Identifier -ErrorAction SilentlyContinue
        if ($zone) {
            W (T 'DLL: Mark-of-the-Web IS PRESENT (the file is flagged as downloaded from the internet)!' `
                  'DLL: ЕСТЬ Mark-of-the-Web (файл помечен как скачанный из интернета)!')
            Problem (T 'The DLL carries a Mark-of-the-Web flag - .NET may refuse to load it. Reinstalling fixes it (Install.cmd removes the flag).' `
                       'На DLL стоит метка Mark-of-the-Web - .NET может отказаться её загружать. Лечится переустановкой (Install.cmd снимает метку).')
        } else {
            W (T 'DLL: no Mark-of-the-Web flag (this is correct).' 'DLL: метки Mark-of-the-Web нет (это правильно).')
        }
    } else {
        W (T 'TabsForWord.dll WAS NOT FOUND!' 'TabsForWord.dll НЕ НАЙДЕНА!')
        Problem (T 'The DLL is missing from the installation folder - the installation never ran or was interrupted.' `
                   'Нет файла DLL в папке установки - установка не выполнялась или прервана.')
    }
} else {
    W ((T 'Installation folder not found: {0}' 'Папка установки не найдена: {0}') -f $installDir)
    Problem (T 'The installation folder is missing - Install.cmd was never run (or Uninstall removed everything).' `
               'Папка установки отсутствует - Install.cmd не запускался (или Uninstall удалил всё).')
}

# ---------------------------------------------------------------- реестр надстройки
Section (T 'Registry' 'Регистрация в реестре')
W (T 'Word add-in key (HKCU, not redirected):' 'Ключ надстройки Word (HKCU, не перенаправляется):')
$addinFound = Dump-RegKey 'CurrentUser' 'Registry64' "Software\Microsoft\Office\Word\Addins\$ProgIdConnect"
if (-not $addinFound) {
    Problem (T 'There is no Addins key - Word does not know about the add-in at all. Run Install.cmd.' `
               'Нет ключа Addins - Word вообще не знает о надстройке. Запустите Install.cmd.')
}
$lb = $null
$k = Open-RegKey 'CurrentUser' 'Registry64' "Software\Microsoft\Office\Word\Addins\$ProgIdConnect"
if ($k) { $lb = $k.GetValue('LoadBehavior'); $k.Close() }
if ($null -ne $lb -and $lb -ne 3) {
    Problem ((T 'LoadBehavior={0} (not 3): Word disabled the add-in after a failure on a previous start. Run Repair.cmd and start Word again.' `
                'LoadBehavior={0} (не 3): Word отключил надстройку после сбоя при прошлом запуске. Запустите Repair.cmd и Word заново.') -f $lb)
}

W ''
W (T 'Connect COM registration (CLSID, both registry views):' 'COM-регистрация Connect (CLSID, оба вида реестра):')
$clsid64 = Dump-RegKey 'CurrentUser' 'Registry64' "Software\Classes\CLSID\$ClsidConnect\InprocServer32"
$clsid32 = Dump-RegKey 'CurrentUser' 'Registry32' "Software\Classes\CLSID\$ClsidConnect\InprocServer32"
W ''
W (T 'TabStripControl COM registration (CLSID, both registry views):' `
      'COM-регистрация TabStripControl (CLSID, оба вида реестра):')
$null = Dump-RegKey 'CurrentUser' 'Registry64' "Software\Classes\CLSID\$ClsidControl\InprocServer32"
$null = Dump-RegKey 'CurrentUser' 'Registry32' "Software\Classes\CLSID\$ClsidControl\InprocServer32"

if ($wordBitness -eq 'x64' -and -not $clsid64) {
    Problem (T '64-bit Word, but there is NO 64-bit COM registration - the add-in cannot load. Reinstall with a current Install.cmd.' `
               '64-битный Word, но НЕТ 64-битной COM-регистрации - надстройка не может загрузиться. Переустановите (Install.cmd новой версии).')
}
if ($wordBitness -eq 'x86' -and -not $clsid32) {
    Problem (T '32-bit Word, but there is NO 32-bit COM registration (WOW6432Node) - a known cause of "it does not work". Reinstall with a current Install.cmd: it registers both views.' `
               '32-битный Word, но НЕТ 32-битной COM-регистрации (WOW6432Node) - это известная причина «не работает». Переустановите новой версией Install.cmd: она регистрирует оба вида.')
}

# ---------------------------------------------------------------- бан-листы Word
Section (T 'Word disabled items (Resiliency)' 'Отключённые элементы Word (Resiliency)')
# Подробно раскрываем только записи про TabsForWord; записи ЧУЖИХ надстроек
# не декодируем - их имена и пути не наше дело и в отчёт не попадают.
$foundResiliency = $false
$foundOurs = $false
foreach ($officeVer in @('16.0', '15.0', '14.0')) {
    foreach ($sub in @('DisabledItems', 'CrashingAddinList')) {
        $k = Open-RegKey 'CurrentUser' 'Registry64' "Software\Microsoft\Office\$officeVer\Word\Resiliency\$sub"
        if ($k) {
            $names = $k.GetValueNames()
            if ($names.Count -gt 0) {
                $foundResiliency = $true
                W "[$officeVer\$sub]:"
                foreach ($n in $names) {
                    $v = $k.GetValue($n)
                    $text = ''
                    if ($v -is [byte[]]) {
                        try { $text = [System.Text.Encoding]::Unicode.GetString($v) } catch { }
                    } else {
                        $text = "$v"
                    }
                    if ($text -match 'TabsForWord') {
                        $foundOurs = $true
                        if ($v -is [byte[]]) {
                            $hex = (($v | ForEach-Object { $_.ToString('X2') }) -join ' ')
                            W ("  {0} = {1}" -f $n, $hex)
                            $printable = ($text -replace '[^ -~Ѐ-ӿ\\.:]+', ' ').Trim()
                            if ($printable) { W ((T '       as text: {0}' '       как текст: {0}') -f $printable) }
                        } else {
                            W ("  {0} = {1}" -f $n, $v)
                        }
                    } else {
                        W ((T '  {0} = [another add-in''s entry - its content is not included in the report]' `
                               '  {0} = [запись другой надстройки - содержимое не включается в отчёт]') -f $n)
                    }
                }
            }
            $k.Close()
        }
    }
}
if ($foundOurs) {
    Problem (T 'Word has disabled the TabsForWord add-in (a Resiliency entry) - run Repair.cmd.' `
               'Word отключал надстройку TabsForWord (запись в Resiliency) - запустите Repair.cmd.')
} elseif ($foundResiliency) {
    W (T 'There are entries, but they belong to other add-ins - not related to TabsForWord.' `
          'Записи есть, но только о других надстройках - к TabsForWord не относятся.')
} else {
    W (T 'Empty (Word has not disabled anything).' 'Пусто (Word ничего не отключал).')
}

# ---------------------------------------------------------------- журнал событий
Section (T 'Windows event log (last 14 days: WINWORD / .NET)' 'Журнал событий Windows (последние 14 дней: WINWORD / .NET)')
try {
    $events = Get-WinEvent -FilterHashtable @{ LogName = 'Application'; StartTime = (Get-Date).AddDays(-14) } `
        -MaxEvents 1500 -ErrorAction Stop |
        Where-Object { $_.ProviderName -in @('Application Error', '.NET Runtime', 'Windows Error Reporting') } |
        Where-Object { ($_.Message -match 'WINWORD|TabsForWord') } |
        Select-Object -First 15
    if ($events) {
        foreach ($e in $events) {
            W ('--- {0:yyyy-MM-dd HH:mm:ss}  [{1}] id={2}' -f $e.TimeCreated, $e.ProviderName, $e.Id)
            $msg = $e.Message
            if ($msg.Length -gt 1800) { $msg = $msg.Substring(0, 1800) + (T ' …(truncated)' ' …(обрезано)') }
            W $msg
        }
        Problem (T 'The Windows event log has WINWORD/.NET errors from the last 14 days - see the event log section of this report.' `
                   'В журнале Windows есть ошибки WINWORD/.NET за 14 дней - см. раздел «Журнал событий» в отчёте.')
    } else {
        W (T 'No WINWORD / .NET Runtime errors in the last 14 days.' `
              'Ошибок WINWORD / .NET Runtime за 14 дней не найдено.')
    }
} catch {
    W ((T 'Could not read the event log: {0}' 'Не удалось прочитать журнал событий: {0}') -f $_.Exception.Message)
}

# ---------------------------------------------------------------- логи надстройки
Section (T 'Add-in and installation logs' 'Логи надстройки и установки')
$logsDir = Join-Path $installDir 'Logs'
$latestLogs = @()
if (Test-Path $logsDir) {
    $all = Get-ChildItem $logsDir -File | Sort-Object LastWriteTime -Descending
    W ((T 'Log files in total: {0}' 'Всего файлов логов: {0}') -f $all.Count)
    foreach ($f in $all | Select-Object -First 12) {
        W ((T '  {0}  {1,9} bytes  {2:yyyy-MM-dd HH:mm}' '  {0}  {1,9} байт  {2:yyyy-MM-dd HH:mm}') -f `
            $f.Name, $f.Length, $f.LastWriteTime)
    }
    $latestLogs = @($all | Where-Object { $_.Name -like 'TabsForWord-*.log' } | Select-Object -First 3) +
                  @($all | Where-Object { $_.Name -like 'install-*.log' } | Select-Object -First 2)
    foreach ($f in $latestLogs) {
        W ''
        W ((T '---------- {0} (last 400 lines) ----------' '---------- {0} (последние 400 строк) ----------') -f $f.Name)
        try { Get-Content $f.FullName -Tail 400 -Encoding UTF8 | ForEach-Object { [void]$report.AppendLine($_) } }
        catch { W ((T 'could not be read: {0}' 'не прочитан: {0}') -f $_.Exception.Message) }
    }
    if (-not ($all | Where-Object { $_.Name -like 'TabsForWord-*.log' })) {
        Problem (T 'There is no add-in log at all: Word has never loaded our DLL (registration / CLR / policies) - or Word was never started after the installation.' `
                   'Нет ни одного лога надстройки: Word ни разу не загружал нашу DLL (регистрация/CLR/политики) - либо Word не запускали после установки.')
    }
} else {
    W ((T 'No logs folder: {0}' 'Папки логов нет: {0}') -f $logsDir)
    Problem (T 'There is no logs folder: the add-in has never run and no logged installation took place.' `
               'Папки логов нет: надстройка ни разу не запускалась и установка с логом не выполнялась.')
}

# ---------------------------------------------------------------- разбор логов надстройки
# Раньше автоитог смотрел только реестр, наличие логов и журнал событий Windows,
# а САМ лог надстройки не читал: на машине с предупреждениями отчёт всё равно
# писал «явных проблем не найдено». Теперь WARN/ERROR разбираются по существу.
Section (T 'Add-in log analysis (WARN / ERROR)' 'Разбор логов надстройки (WARN / ERROR)')

# Ключ сообщения: убрать метку времени, уровень и переменные части (hwnd, числа),
# чтобы одинаковые по смыслу строки схлопнулись в одну группу со счётчиком.
# Заодно из итога не попадают имена документов - они и так есть в дампе лога выше.
function Get-LogKind([string]$line) {
    $t = $line -replace '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d+\s*\[[A-Za-z ]+\]\s*', ''
    $t = $t -replace '0[xX][0-9A-Fa-f]+', 'HWND'   # до маскировки цифр, иначе '0' съест следующее правило
    $t = $t -replace '\d+', 'N'
    $t = $t -replace '(name|путь|path)=.*$', '$1=...'
    return $t.Trim()
}

# Предупреждения, которые означают реальную деградацию для пользователя.
$SeriousPatterns = @(
    'Reserve mode disabled',
    'Reserve mode suspended',
    'native-host.cfg read failed',
    'Locator: GetClientRect failed',
    'Locator snapshot failed',
    # 1.7.3: переключение вкладкой не удержалось (окно не вышло вперёд или его вернули назад)
    'Activate check: target'
)
# Предупреждения, штатные для переходных состояний (окно строится/закрывается).
$TransientPatterns = @(
    'Locator: no content zone found',
    'Activate (COM) failed',
    'Activate: window not found',
    'PV activate failed',
    'PV close failed',
    'Close failed for hwnd',
    'Initialize deferred: CTP factory not yet available',
    'EnumChildWindows failed',
    'LogEnvironment failed'
)
function Test-AnyPattern([string]$text, [string[]]$patterns) {
    foreach ($p in $patterns) { if ($text.Contains($p)) { return $true } }
    return $false
}

$addinLogs = @()
if (Test-Path $logsDir) {
    # Только два последних файла: лог дневной, и счётчики за неделю смешали бы
    # свежий сеанс со старыми (в т.ч. с прежними версиями надстройки).
    $addinLogs = @(Get-ChildItem $logsDir -File -Filter 'TabsForWord-*.log' -ErrorAction SilentlyContinue |
                   Sort-Object LastWriteTime -Descending | Select-Object -First 2)
}

if ($addinLogs.Count -eq 0) {
    W (T 'There are no add-in logs - nothing to analyse (see the problem above).' `
          'Логов надстройки нет - разбирать нечего (см. проблему выше).')
} else {
    W (T 'Analysed in full (not just the 400 lines shown above):' `
          'Разобрано целиком (а не только показанные выше 400 строк):')
    foreach ($f in $addinLogs) { W ("  {0}  ({1:yyyy-MM-dd HH:mm})" -f $f.Name, $f.LastWriteTime) }

    $errKinds  = @{}
    $warnKinds = @{}
    $startCount = 0
    $lastStart = ''
    $fallbackToCtp = $false
    $hostMode = ''         # Native (основной, с 1.6.0) или CustomTaskPane (классический)
    $panelAtTop = 0        # панель встала в самый верх окна - так быть не должно
    $logVersion = ''       # версия надстройки, которая ФАКТИЧЕСКИ работала
    # Файлы разбираем от старого к новому, чтобы «последний запуск» был правда последним.
    foreach ($f in @($addinLogs | Sort-Object LastWriteTime)) {
        $lines = @()
        try { $lines = Get-Content $f.FullName -Encoding UTF8 -ErrorAction Stop } catch { $lines = @() }
        foreach ($ln in $lines) {
            if ($ln -match '\[ERROR') {
                $k = Get-LogKind $ln
                if ($errKinds.ContainsKey($k)) { $errKinds[$k]++ } else { $errKinds[$k] = 1 }
                if ($ln -match 'falling back to CustomTaskPane' -or $ln -match 'Native host init failed') { $fallbackToCtp = $true }
            }
            elseif ($ln -match '\[WARN') {
                $k = Get-LogKind $ln
                if ($warnKinds.ContainsKey($k)) { $warnKinds[$k]++ } else { $warnKinds[$k] = 1 }
            }
            elseif ($ln -match 'Add-in started') {
                $startCount++
                if ($ln.Length -ge 19) { $lastStart = $ln.Substring(0, 19) }
            }
            elseif ($ln -match 'Add-in connecting.*version=([0-9.]+)') {
                $logVersion = $matches[1]
            }
            elseif ($ln -match 'Tab host mode: (\w+)') {
                # С версии 1.6.0 режим пишется в лог всегда. Классический режим
                # сам по себе не поломка (его можно включить файлом), но в отчёте
                # он должен быть виден - иначе вопрос «почему полоса сверху?»
                # разбирается вслепую.
                $hostMode = $matches[1]
            }
            elseif ($ln -match 'repositioned \(shown\).*rect=(-?\d+),(-?\d+),') {
                # Верх клиентской области Word занят заголовком и лентой, поэтому
                # y=0 - всегда ошибка: панель легла поверх них. Строка уровня INFO,
                # обычным поиском ошибок такое не ловится (случай машины A).
                if ([int]$matches[2] -le 0) { $panelAtTop++ }
            }
        }
    }

    $errTotal  = 0; foreach ($v in $errKinds.Values)  { $errTotal  += $v }
    $warnTotal = 0; foreach ($v in $warnKinds.Values) { $warnTotal += $v }

    if ($logVersion) {
        W ((T 'Add-in version according to the log (what actually ran): {0}' `
               'Версия надстройки по логу (что реально работало): {0}') -f $logVersion)
    }
    if ($hostMode -eq 'Native') {
        W (T 'Tab bar mode: main (tabs inside the Word window).' `
              'Режим панели: основной (вкладки внутри окна Word).')
    } elseif ($hostMode -eq 'CustomTaskPane') {
        W (T 'Tab bar mode: CLASSIC (a Custom Task Pane strip with its own title bar).' `
              'Режим панели: КЛАССИЧЕСКИЙ (полоса Custom Task Pane со служебным заголовком).')
        W (T '  That happens after running "Classic-mode.cmd" or with mode=ctp set' `
              '  Так бывает, если запускали «Classic-mode.cmd» или задан mode=ctp')
        W ((T '  in {0}\TabsForWord\native-host.cfg.' '  в файле {0}\TabsForWord\native-host.cfg.') -f $env:LOCALAPPDATA)
    } elseif ($logVersion) {
        W (T 'Tab bar mode: not stated in the log (add-in older than version 1.6.0).' `
              'Режим панели: в логе не указан (надстройка старее версии 1.6.0).')
    }
    W ((T 'Successful add-in starts ("Add-in started"): {0}{1}' `
           'Успешных запусков надстройки («Add-in started»): {0}{1}') -f $startCount,
        $(if ($lastStart) { (T '; last one ' '; последний ') + $lastStart } else { '' }))
    W ((T 'ERROR lines: {0}; WARN lines: {1}' 'Строк ERROR: {0}; строк WARN: {1}') -f $errTotal, $warnTotal)

    if ($startCount -eq 0) {
        Problem (T 'The log has no "Add-in started" line at all: Word loaded the DLL, but the add-in never started - look at the ERROR lines in the log dump.' `
                   'В логе нет ни одной строки «Add-in started»: Word загружал DLL, но надстройка не запустилась - смотрите строки ERROR в дампе лога.')
    }

    if ($panelAtTop -gt 0) {
        W ''
        W ((T 'The tab bar landed at the very top of the window: {0} time(s)' `
               'Панель вставала в самый верх окна: {0} раз(а)') -f $panelAtTop)
        Problem ((T 'The tab bar landed at the very top of the Word window {0} time(s) - over the title bar and the ribbon (usually briefly, while Word starts). Fixed in version 1.5.1: if this report shows an older version, update the add-in.' `
                    'Панель вкладок {0} раз(а) вставала в самый верх окна Word - поверх заголовка и ленты (обычно кратко, на старте Word). Исправлено в версии 1.5.1: если в отчёте версия старее - обновите надстройку.') -f $panelAtTop)
    }

    if ($errTotal -gt 0) {
        W ''
        W (T 'ERRORS, by kind:' 'ОШИБКИ (ERROR), по видам:')
        foreach ($e in ($errKinds.GetEnumerator() | Sort-Object Value -Descending)) {
            W ("  x{0,-4} {1}" -f $e.Value, $e.Key)
        }
        Problem ((T 'The add-in log has {0} ERROR line(s) - see the "Add-in log analysis" section of this report.' `
                    'В логе надстройки {0} строк(и) ERROR - раздел «Разбор логов надстройки» в этом отчёте.') -f $errTotal)
    } else {
        W (T 'No ERROR lines.' 'Строк ERROR нет.')
    }

    if ($fallbackToCtp) {
        Problem (T 'The main tab bar mode did not work and the add-in switched to the fallback (Custom Task Pane strip).' `
                   'Native-режим панели не заработал, надстройка перешла на запасной режим (полоса CustomTaskPane).')
    }

    if ($warnTotal -gt 0) {
        $serious = @(); $transient = @(); $other = @()
        foreach ($w in ($warnKinds.GetEnumerator() | Sort-Object Value -Descending)) {
            if (Test-AnyPattern $w.Key $SeriousPatterns)        { $serious   += $w }
            elseif (Test-AnyPattern $w.Key $TransientPatterns)  { $transient += $w }
            else                                                { $other     += $w }
        }

        if ($serious.Count -gt 0) {
            W ''
            W (T 'WARNINGS that affect how the tab bar works:' 'ПРЕДУПРЕЖДЕНИЯ, влияющие на работу:')
            foreach ($w in $serious) { W ("  x{0,-4} {1}" -f $w.Value, $w.Key) }
            Problem (T 'There are warnings that affect how the tab bar works (the list is in the "Add-in log analysis" section).' `
                       'Есть предупреждения, влияющие на работу панели (список - в разделе «Разбор логов надстройки»).')
        }

        if ($transient.Count -gt 0) {
            W ''
            W (T 'Normal transient warnings (expected while a Word window is being built or closed):' `
                  'Штатные переходные предупреждения (норма: окно Word строится или закрывается):')
            foreach ($w in $transient) { W ("  x{0,-4} {1}" -f $w.Value, $w.Key) }
            # Счётчик здесь НЕ повод для тревоги: «no content zone» надстройка
            # выдаёт штатно - пока Word не разложил новое окно и когда документ
            # закрывается, то есть по 1-2 раза на каждое окно за сеанс. Признак
            # настоящего отказа - дампы дерева окон ниже: их надстройка пишет
            # только после шести неудач подряд по одному и тому же окну.
        }

        if ($other.Count -gt 0) {
            W ''
            W (T 'Other warnings:' 'Прочие предупреждения:')
            foreach ($w in $other) { W ("  x{0,-4} {1}" -f $w.Value, $w.Key) }
            Problem (T 'The log contains warnings of an unknown kind (the list is in the "Add-in log analysis" section).' `
                       'В логе есть предупреждения неизвестного вида (список - в разделе «Разбор логов надстройки»).')
        }
    } else {
        W (T 'No WARN lines.' 'Строк WARN нет.')
    }

    # Дампы дерева окон надстройка пишет сама, когда локатор отказывает подряд.
    # Исключение - дампы, снятые при открытом меню «Файл» (полностраничный UI
    # Word): области документа в этот момент действительно нет, панель штатно
    # прячется, ломаться нечему. Надстройка с 1.7.1 такие дампы вовсе не пишет,
    # но у розданных сборок 1.6.0-1.7.0 они встречаются (случай машины C), и по
    # ним итог поднимал ложную тревогу. Признак в дампе - видимое окно класса
    # FullpageUIHost поверх клиентской области.
    $dumps = @(Get-ChildItem $logsDir -File -Filter 'WindowTree-*.txt' -ErrorAction SilentlyContinue)
    if ($dumps.Count -gt 0) {
        $benign = 0
        foreach ($d in $dumps) {
            $body = ''
            try { $body = Get-Content $d.FullName -Raw -Encoding UTF8 -ErrorAction Stop } catch { $body = '' }
            if ($body -match '"FullpageUIHost"\s+visible') { $benign++ }
        }
        W ''
        W ((T 'Window-tree dumps (WindowTree-*.txt): {0}' 'Дампов дерева окон (WindowTree-*.txt): {0}') -f $dumps.Count)
        if ($benign -gt 0) {
            W ((T '  of them {0} were taken while the full-page Word UI (the File menu) was open - that is normal: there is no document area then and the tab bar hides on purpose.' `
                  '  из них {0} сняты при открытом полностраничном экране Word (меню «Файл») - это норма: области документа в этот момент нет, панель прячется намеренно.') -f $benign)
        }
        if ($dumps.Count -gt $benign) {
            Problem ((T 'The add-in saved {0} window-tree dump(s): the locator failed to find the document area several times in a row.' `
                        'Надстройка сохранила {0} дамп(ов) дерева окон: локатор несколько раз подряд не смог найти область документа.') -f ($dumps.Count - $benign))
        }
    }
}

# ---------------------------------------------------------------- резюме
Section (T 'SUMMARY (automatic analysis)' 'РЕЗЮМЕ (автоматический разбор)')
if ($problems.Count -eq 0) {
    W (T 'No obvious problems found. If the tabs still do not appear, send this report to the developer.' `
          'Явных проблем не найдено. Если вкладок всё равно нет - пришлите этот отчёт разработчику.')
} else {
    $i = 0
    foreach ($p in $problems) { $i++; W ("{0}. {1}" -f $i, $p) }
}

# ---------------------------------------------------------------- сохранение
$desktop = [Environment]::GetFolderPath('Desktop')
$outDir = Join-Path $env:TEMP "TabsForWord-diag-$ts"
New-Item -ItemType Directory -Force $outDir | Out-Null
# Имя файла - ASCII: на чужой локали кириллица в имени читается плохо, а файл
# ещё и пересылают почтой (этап 30).
$txtPath = Join-Path $outDir "TabsForWord-diagnostics-$env:COMPUTERNAME-$ts.txt"
[System.IO.File]::WriteAllText($txtPath, $report.ToString(), (New-Object System.Text.UTF8Encoding $true))
if (Test-Path $logsDir) {
    Copy-Item $logsDir (Join-Path $outDir 'Logs') -Recurse -Force -ErrorAction SilentlyContinue
}
$zipPath = Join-Path $desktop "TabsForWord-diagnostics-$env:COMPUTERNAME-$ts.zip"
try {
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
    Compress-Archive -Path (Join-Path $outDir '*') -DestinationPath $zipPath -Force
    Remove-Item $outDir -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host ''
    Write-Host (T 'DONE. The report has been saved to your desktop:' `
                  'ГОТОВО. Отчёт сохранён на рабочем столе:') -ForegroundColor Green
    Write-Host "  $zipPath" -ForegroundColor Green
    Write-Host ''
    Write-Host (T "What is inside: Word and Windows versions, the state of the installation, the" `
                  'Что внутри: версии Word и Windows, состояние установки, журналы')
    Write-Host (T "add-in's logs and errors from the Windows event log. The logs contain the" `
                  'надстройки, ошибки из журнала событий Windows. В журналах встречаются')
    Write-Host (T 'NAMES and PATHS of your documents (but not their content), plus the computer' `
                  'ИМЕНА и ПУТИ ваших документов (но не их содержимое), а также имя')
    Write-Host (T 'name and the user name.' 'компьютера и имя пользователя.')
    Write-Host (T 'The archive holds plain text files: you can open and read them before sending.' `
                  'Внутри архива - обычные текстовые файлы: перед отправкой их можно')
    Write-Host (T '' 'открыть и просмотреть.')
    Write-Host ''
    Write-Host (T 'If you are fine with that, send this file to the developer.' `
                  'Если всё устраивает - пришлите этот файл разработчику.') -ForegroundColor Green
} catch {
    Write-Host ((T 'Could not create the ZIP ({0}). The report folder is: {1}' `
                   'Не удалось создать ZIP ({0}). Папка с отчётом: {1}') -f $_.Exception.Message, $outDir) -ForegroundColor Yellow
}
exit 0
