# install.ps1 - installs the "Tabs for Word" add-in for the current user.
# No administrator rights required. Word must be closed.
# Usage: double-click Install.cmd, or:
#   powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1 [-Lang en|ru]
# The full installation log is written to %LOCALAPPDATA%\TabsForWord\Logs\install-*.log
#
# Язык сообщений: -Lang, иначе по языку интерфейса Windows (всё, кроме русского,
# показывается по-английски). Сама надстройка определяет язык иначе - по языку
# интерфейса Word, - потому что в момент установки Word закрыт и спросить некого.
param([ValidateSet('auto', 'en', 'ru')][string]$Lang = 'auto')
$ErrorActionPreference = 'Stop'

if ($Lang -eq 'auto') {
    $Lang = if ((Get-UICulture).TwoLetterISOLanguageName -eq 'ru') { 'ru' } else { 'en' }
}
$script:Lang = $Lang

# Английский - базовый, русский - перевод (тот же приём, что Strings.T в коде).
function T([string]$en, [string]$ru) { if ($script:Lang -eq 'ru') { return $ru } return $en }

# --- Константы (должны совпадать с атрибутами классов в исходном коде) ---
$ClsidConnect   = '{A3B7C9D1-5E2F-4A6B-8C0D-1F2E3D4C5B6A}'
$ClsidControl   = '{D4E5F6A7-8B9C-4D0E-A1B2-C3D4E5F6A7B8}'
$ProgIdConnect  = 'TabsForWord.Connect'
$ProgIdControl  = 'TabsForWord.TabStripControl'
$ClassConnect   = 'TabsForWord.Connect'
$ClassControl   = 'TabsForWord.TabStripControl'
$RuntimeVersion = 'v4.0.30319'
# Имя в списке «Надстройки COM» самого Word - на языке установки. «MVP» убрано в 1.6.0.
$FriendlyName = T 'Tabs for Word' 'Вкладки для Word'
$Description  = T 'A tab bar for the documents open in Word' 'Панель вкладок открытых документов Word'

# Страница загрузки .NET Framework 4.8 (ДОЛЖНО СОВПАДАТЬ с diagnose.ps1)
$NetDownloadUrl = T 'https://dotnet.microsoft.com/download/dotnet-framework/net48' `
                    'https://dotnet.microsoft.com/ru-ru/download/dotnet-framework/net48'

# --- Битность: ключи CLSID в HKCU перенаправляются WOW64 - нужен 64-битный процесс ---
if (-not [Environment]::Is64BitProcess) {
    $sysnative = Join-Path $env:windir 'sysnative\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path $sysnative) {
        # -Lang передаём явно: у 64-битного процесса та же локаль, но если язык
        # выбрали параметром, перезапуск обязан его сохранить.
        & $sysnative -NoProfile -ExecutionPolicy Bypass -File $MyInvocation.MyCommand.Path -Lang $script:Lang
        exit $LASTEXITCODE
    }
    Write-Host (T 'ERROR: 64-bit PowerShell is required (32-bit Windows is not supported).' `
                  'ОШИБКА: требуется 64-битный PowerShell (Windows 32-bit не поддерживается).') -ForegroundColor Red
    exit 1
}

# --- Лог установки: всё, что видит пользователь, попадает и в файл ---
$script:LogLines = New-Object System.Collections.ArrayList
function Log([string]$msg, [string]$color) {
    if ($color) { Write-Host $msg -ForegroundColor $color } else { Write-Host $msg }
    [void]$script:LogLines.Add(('{0:yyyy-MM-dd HH:mm:ss.fff} {1}' -f (Get-Date), $msg))
}
function Save-InstallLog([bool]$failed) {
    try {
        $logDir = Join-Path (Join-Path $env:LOCALAPPDATA 'TabsForWord') 'Logs'
        New-Item -ItemType Directory -Force $logDir | Out-Null
        $logPath = Join-Path $logDir ('install-{0:yyyyMMdd-HHmmss}.log' -f (Get-Date))
        [System.IO.File]::WriteAllLines($logPath, $script:LogLines, (New-Object System.Text.UTF8Encoding $true))
        if ($failed) {
            # На сбое путь к логу нужен чаще всего - показываем его заметно.
            Write-Host ''
            Write-Host ((T 'Installation log: {0}' 'Лог установки: {0}') -f $logPath) -ForegroundColor Yellow
            Write-Host (T 'Send this file if you need help.' 'Пришлите этот файл, если понадобится помощь.') -ForegroundColor Yellow
        } else {
            Write-Host ((T 'Installation log: {0}' 'Лог установки: {0}') -f $logPath) -ForegroundColor DarkGray
        }
    } catch { }
}

# --- Сообщение «нужен .NET Framework 4.8» (общее для двух аварийных случаев) ---
function Show-NetRequirement([string]$problem, [int]$winBuild) {
    Log ''
    Log ((T 'ERROR: {0}' 'ОШИБКА: {0}') -f $problem) 'Red'
    Log (T 'The "Tabs for Word" add-in requires Microsoft .NET Framework 4.8 or newer.' `
           'Надстройке «Вкладки для Word» требуется Microsoft .NET Framework 4.8 или новее.')
    # 18362 = Windows 10 версии 1903 - первая, где .NET 4.8 входит в состав системы.
    if ($winBuild -gt 0 -and $winBuild -lt 18362) {
        Log ''
        Log ((T 'Your Windows version (build {0}) does not include .NET Framework 4.8,' `
                'В вашей версии Windows (сборка {0}) .NET Framework 4.8 не входит в состав') -f $winBuild)
        Log (T 'so it has to be installed separately. Windows 10 version 1903 (May 2019)' `
               'системы, его нужно доустановить. Начиная с Windows 10 версии 1903 (май 2019)')
        Log (T 'and every Windows 11 already have it built in.' `
               'и во всех Windows 11 он уже встроен.')
    }
    Log ''
    Log (T 'What to do:' 'Что делать:')
    Log (T '  1. Install .NET Framework 4.8 from this page:' '  1. Установите .NET Framework 4.8 по ссылке:')
    Log "     $script:NetDownloadUrl" 'Cyan'
    Log (T '  2. Restart the computer if Windows asks for it.' '  2. Если система попросит - перезагрузите компьютер.')
    Log (T '  3. Close every Microsoft Word window.' '  3. Закройте все окна Microsoft Word.')
    Log (T '  4. Run the installation again.' '  4. Запустите установку ещё раз.')
    Log ''
    Log (T 'Note: installing .NET Framework is a one-time, system-wide operation and it' `
           'Обратите внимание: установка .NET Framework - разовая общесистемная операция,')
    Log (T 'does require administrator rights. That is a Microsoft requirement, not ours:' `
           'и она требует прав администратора. Это требование Microsoft, а не нашей')
    Log (T 'the "Tabs for Word" add-in itself never needs administrator rights.' `
           'надстройки. Сама надстройка «Вкладки для Word» прав администратора НЕ требует.')
    Log ''
    # Обе операции ниже необязательные: ссылка уже напечатана и попала в лог.
    try {
        Set-Clipboard -Value $script:NetDownloadUrl -ErrorAction Stop
        Log (T 'The link has been copied to the clipboard.' 'Ссылка скопирована в буфер обмена.')
    } catch { }
    try {
        $answer = Read-Host (T 'Open the .NET download page in the browser? (Y/n)' `
                               'Открыть страницу загрузки .NET в браузере? (Д/н)')
        if ($answer -notmatch '^\s*[nNнН]') { Start-Process $script:NetDownloadUrl | Out-Null }
    } catch { }
    Log ''
    Log (T 'NOTHING WAS INSTALLED.' 'Установка НЕ выполнена.') 'Red'
}

# --- Сведения об окружении (для диагностики проблем на разных машинах) ---
function Get-ExeBitness([string]$path) {
    # PE-заголовок: смещение по 0x3C, поле Machine: 0x8664=x64, 0x14C=x86, 0xAA64=ARM64
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

function Find-WordInfo {
    # Возвращает объект: Path, Bitness, FileVersion, C2R, Detected, DetectedBy.
    # Word ищется ТРЕМЯ независимыми способами: одного признака мало - у части
    # установок Office ключ App Paths отсутствует, и это ещё не значит, что Word нет.
    $info = [pscustomobject]@{
        Path = $null; Bitness = $null; FileVersion = $null; C2R = $null
        Detected = $false; DetectedBy = $null
    }

    # (1) Реестр: App Paths + попутно сведения Click-to-Run
    foreach ($view in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
        foreach ($hive in @([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryHive]::CurrentUser)) {
            try {
                $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
                $k = $base.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe')
                if ($k) {
                    $p = $k.GetValue('')
                    if ($p -and (Test-Path $p) -and -not $info.Path) {
                        $info.Path = $p
                        $info.DetectedBy = (T 'App Paths registry key' 'реестр App Paths')
                    }
                    $k.Close()
                }
                $c = $base.OpenSubKey('SOFTWARE\Microsoft\Office\ClickToRun\Configuration')
                if ($c) {
                    if (-not $info.C2R) {
                        $info.C2R = (T 'version {0}, platform {1}, products {2}' 'версия {0}, платформа {1}, продукты {2}') -f `
                            $c.GetValue('VersionToReport'), $c.GetValue('Platform'), $c.GetValue('ProductReleaseIds')
                    }
                    $c.Close()
                }
                $base.Close()
            } catch { }
        }
    }

    # (2) winword.exe в стандартных папках Office (если в реестре не нашлось)
    if (-not $info.Path) {
        foreach ($progFiles in @($env:ProgramFiles, ${env:ProgramFiles(x86)})) {
            if (-not $progFiles) { continue }
            foreach ($rel in @(
                'Microsoft Office\root\Office16\WINWORD.EXE',
                'Microsoft Office\Office16\WINWORD.EXE',
                'Microsoft Office\root\Office15\WINWORD.EXE',
                'Microsoft Office\Office15\WINWORD.EXE',
                'Microsoft Office\Office14\WINWORD.EXE'
            )) {
                $candidate = Join-Path $progFiles $rel
                if (Test-Path $candidate) {
                    $info.Path = $candidate
                    $info.DetectedBy = (T 'winword.exe file in an Office folder' 'файл winword.exe в папке Office')
                    break
                }
            }
            if ($info.Path) { break }
        }
    }

    # (3) COM-регистрация Word.Application - есть у любой установки настольного Word
    if (-not $info.Path) {
        foreach ($view in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
            foreach ($hive in @([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryHive]::CurrentUser)) {
                if ($info.Detected) { continue }
                try {
                    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
                    $k = $base.OpenSubKey('SOFTWARE\Classes\Word.Application')
                    if ($k) {
                        $info.Detected = $true
                        $info.DetectedBy = (T 'Word.Application COM registration' 'COM-регистрация Word.Application')
                        $k.Close()
                    }
                    $base.Close()
                } catch { }
            }
        }
    }

    if ($info.Path) {
        $info.Detected = $true
        $info.Bitness = Get-ExeBitness $info.Path
        try { $info.FileVersion = (Get-Item $info.Path).VersionInfo.FileVersion } catch { }
    }
    return $info
}

# Тело установки - в функции: return корректно возвращает код выхода,
# а finally снаружи гарантированно сохраняет лог даже при сбое.
function Invoke-Install {
    Log (T '=== Installing: Tabs for Word ===' '=== Установка: Вкладки для Word ===') 'Cyan'
    Log (T 'Requires: Windows 10 version 1903 (May 2019) or newer, or Windows 11' `
           'Требуется: Windows 10 версии 1903 (май 2019) или новее, либо Windows 11')
    Log (T '(both already include .NET Framework 4.8), and desktop Microsoft Word.' `
           '(на них .NET Framework 4.8 уже встроен), и настольный Microsoft Word.')
    Log (T 'Administrator rights are NOT required.' 'Права администратора НЕ требуются.')
    Log ''

    # --- Окружение ---
    $winBuild = 0
    try {
        $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
        try { $winBuild = [int]$cv.CurrentBuild } catch { }
        Log ((T 'Windows: {0} {1} (build {2}.{3}), 64-bit OS: {4}' `
                'Windows: {0} {1} (сборка {2}.{3}), 64-bit ОС: {4}') -f `
            $cv.ProductName, $cv.DisplayVersion, $cv.CurrentBuild, $cv.UBR, [Environment]::Is64BitOperatingSystem)
    } catch {
        Log ((T 'Windows: could not read the version ({0})' `
                'Windows: не удалось прочитать версию ({0})') -f $_.Exception.Message)
    }
    Log ((T 'PowerShell: {0}, 64-bit process: {1}, user: {2}' `
            'PowerShell: {0}, процесс 64-bit: {1}, пользователь: {2}') -f `
        $PSVersionTable.PSVersion, [Environment]::Is64BitProcess, $env:USERNAME)

    # --- .NET Framework 4.8 - обязателен ---
    # Сборка net48 + NoPIA: без 4.8 Word не сможет активировать COM-компонент.
    # Три случая, а не два:
    #   1) ключ прочитан, версия ниже 4.8         -> ФАТАЛЬНО (ясная причина);
    #   2) ключа нет И файлов .NET 4 нет на диске -> ФАТАЛЬНО (.NET 4 вообще нет);
    #   3) ключа нет, но файлы .NET 4 на месте    -> только предупреждение (сбой
    #      чтения реестра не должен блокировать заведомо рабочую машину).
    $netRelease = 0
    $netVersion = $null
    $netReadOk  = $false
    try {
        $ndp = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction Stop
        if ($null -eq $ndp.Release) { throw 'в ключе нет значения Release' }
        $netRelease = [int]$ndp.Release
        $netVersion = "$($ndp.Version)"
        $netReadOk  = $true
        Log ((T '.NET Framework: version {0} (Release={1}; 528040 or higher is required, that is 4.8)' `
                '.NET Framework: версия {0} (Release={1}; требуется >=528040, это 4.8)') -f $netVersion, $netRelease)
    } catch {
        Log ((T '.NET Framework: could not read the version ({0}).' `
                '.NET Framework: не удалось прочитать версию ({0}).') -f $_.Exception.Message) 'Yellow'
    }

    if ($netReadOk -and $netRelease -lt 528040) {
        $shown = if ($netVersion) { $netVersion } else { "Release=$netRelease" }
        Show-NetRequirement ((T 'this computer has .NET Framework {0}, which is not enough.' `
                                'на компьютере установлен .NET Framework {0} - этой версии недостаточно.') -f $shown) $winBuild
        return 1
    }
    if (-not $netReadOk) {
        $clrFound = $false
        foreach ($clr in @(
            (Join-Path $env:windir 'Microsoft.NET\Framework64\v4.0.30319\mscorlib.dll'),
            (Join-Path $env:windir 'Microsoft.NET\Framework\v4.0.30319\mscorlib.dll')
        )) {
            if (Test-Path $clr) { $clrFound = $true; break }
        }
        if (-not $clrFound) {
            Show-NetRequirement (T 'Microsoft .NET Framework 4 is not installed on this computer.' `
                                   'Microsoft .NET Framework 4 не установлен на этом компьютере.') $winBuild
            return 1
        }
        Log (T 'WARNING: could not verify the .NET version, but the .NET 4 files are there - continuing.' `
               'ВНИМАНИЕ: версию .NET проверить не удалось, но файлы .NET 4 на месте - продолжаем.') 'Yellow'
        Log (T 'If the tab bar does not appear in Word, install .NET Framework 4.8:' `
               'Если панель в Word не появится - установите .NET Framework 4.8:') 'Yellow'
        Log "  $NetDownloadUrl" 'Cyan'
    }

    # --- Microsoft Word должен быть установлен ---
    $word = Find-WordInfo
    if ($word.Path) {
        Log ((T 'Word: {0} (found by: {1})' 'Word: {0} (найден: {1})') -f $word.Path, $word.DetectedBy)
        Log ((T 'Word: {0}, file version {1}' 'Word: битность {0}, файловая версия {1}') -f $word.Bitness, $word.FileVersion)
        if ($word.C2R) { Log ("Office Click-to-Run: {0}" -f $word.C2R) }
    } elseif ($word.Detected) {
        # Word есть, но winword.exe не нашли - ставим, только предупреждаем.
        Log ((T 'Word: detected by "{0}", but winword.exe itself was not found.' `
                'Word: обнаружен по признаку «{0}», но сам winword.exe не найден.') -f $word.DetectedBy) 'Yellow'
        if ($word.C2R) { Log ("Office Click-to-Run: {0}" -f $word.C2R) }
    } else {
        Log ''
        Log (T 'ERROR: Microsoft Word was not found on this computer.' `
               'ОШИБКА: Microsoft Word на этом компьютере не найден.') 'Red'
        Log (T 'Checked the App Paths registry key, the standard Office folders and the' `
               'Проверены реестр App Paths, стандартные папки Office и COM-регистрация')
        Log (T 'Word.Application COM registration - no sign of a desktop Word installation.' `
               'Word.Application - ни одного признака установленного настольного Word.')
        Log ''
        Log (T 'The add-in works ONLY with desktop Word for Windows. It does not work with' `
               'Надстройка работает ТОЛЬКО с настольным Word для Windows. Она не работает')
        Log (T 'Word Online (in a browser), Word for macOS or Word on mobile devices.' `
               'с Word Online (в браузере), Word для macOS и мобильным Word.')
        Log ''
        Log (T 'What to do:' 'Что делать:')
        Log (T '  1. Install desktop Microsoft Word.' '  1. Установите настольный Microsoft Word.')
        Log (T '  2. Start it at least once, then close it.' '  2. Запустите его хотя бы один раз, затем закройте.')
        Log (T '  3. Run the installation again.' '  3. Запустите установку ещё раз.')
        Log ''
        Log (T 'NOTHING WAS INSTALLED.' 'Установка НЕ выполнена.') 'Red'
        return 1
    }

    # --- Word должен быть закрыт (НЕ завершаем принудительно: возможны несохранённые документы) ---
    $winword = Get-Process WINWORD -ErrorAction SilentlyContinue
    if ($winword) {
        Log ''
        Log (T 'Microsoft Word is running right now (possibly in the background - Office Startup Boost).' `
               'Microsoft Word сейчас запущен (возможно, в фоне - Office Startup Boost).') 'Yellow'
        Log (T 'Please:' 'Пожалуйста:')
        Log (T '  1. Save your documents and close every Word window.' '  1. Сохраните документы и закройте все окна Word.')
        Log (T '  2. Run the installation again.' '  2. Запустите установку ещё раз.')
        Log ''
        Log (T 'NOTHING WAS INSTALLED.' 'Установка НЕ выполнена.') 'Red'
        return 1
    }

    # --- Поиск DLL: рядом со скриптом (пакет установки) или в dev-сборке ---
    $here = $PSScriptRoot
    $dllSource = $null
    foreach ($candidate in @(
        (Join-Path $here 'TabsForWord.dll'),
        (Join-Path $here '..\src\TabsForWord\bin\Release\TabsForWord.dll')
    )) {
        if (Test-Path $candidate) { $dllSource = (Resolve-Path $candidate).Path; break }
    }
    if (-not $dllSource) {
        Log (T 'ERROR: TabsForWord.dll was not found next to the installation script.' `
               'ОШИБКА: не найден TabsForWord.dll рядом со скриптом установки.') 'Red'
        return 1
    }

    # --- Копирование в %LOCALAPPDATA%\TabsForWord ---
    $installDir = Join-Path $env:LOCALAPPDATA 'TabsForWord'
    New-Item -ItemType Directory -Force $installDir | Out-Null
    Copy-Item $dllSource (Join-Path $installDir 'TabsForWord.dll') -Force
    # Снять Mark-of-the-Web (иначе .NET откажется грузить DLL, скачанную из интернета)
    Get-ChildItem $installDir -Recurse -File | Unblock-File -ErrorAction SilentlyContinue
    $dllPath = Join-Path $installDir 'TabsForWord.dll'
    $codeBase = 'file:///' + $dllPath.Replace('\', '/')

    # Точное имя сборки читаем из самой DLL (версия меняется от релиза к релизу)
    $asmName = [System.Reflection.AssemblyName]::GetAssemblyName($dllPath)
    $script:AssemblyName = $asmName.FullName
    $script:AssemblyVersion = $asmName.Version.ToString()
    $script:codeBase = $codeBase
    $dllVer = (Get-Item $dllPath).VersionInfo.FileVersion
    $dllHash = (Get-FileHash $dllPath -Algorithm SHA256).Hash
    Log ((T 'Files copied to: {0}' 'Файлы скопированы: {0}') -f $installDir)
    Log ((T 'DLL: version {0}, assembly {1}' 'DLL: версия {0}, сборка {1}') -f $dllVer, $script:AssemblyName)
    Log "DLL: SHA256 $dllHash"

    # --- Регистрация COM-классов (HKCU\Software\Classes) ---
    # ВАЖНО: подраздел CLSID в HKCU перенаправляется WOW64, поэтому пишем в ОБА
    # вида реестра: 64-битный (x64 Word) и 32-битный (WOW6432Node, x86 Word).
    # ProgId-ключи не перенаправляются - их пишем один раз.
    function Register-ProgId([string]$progId, [string]$clsid, [string]$className) {
        $classes = 'HKCU:\Software\Classes'
        $null = New-Item -Path "$classes\$progId" -Value $className -Force
        $null = New-Item -Path "$classes\$progId\CLSID" -Value $clsid -Force
    }
    function Register-Clsid([string]$clsidRoot, [string]$clsid, [string]$progId, [string]$className) {
        $clsidKey = "$clsidRoot\$clsid"
        $null = New-Item -Path $clsidKey -Value $className -Force
        $inproc = "$clsidKey\InprocServer32"
        $null = New-Item -Path $inproc -Value 'mscoree.dll' -Force
        Set-ItemProperty -Path $inproc -Name 'ThreadingModel' -Value 'Both'
        Set-ItemProperty -Path $inproc -Name 'Class' -Value $className
        Set-ItemProperty -Path $inproc -Name 'Assembly' -Value $script:AssemblyName
        Set-ItemProperty -Path $inproc -Name 'RuntimeVersion' -Value $script:RuntimeVersion
        Set-ItemProperty -Path $inproc -Name 'CodeBase' -Value $script:codeBase

        $verKey = "$inproc\$script:AssemblyVersion"
        $null = New-Item -Path $verKey -Force
        Set-ItemProperty -Path $verKey -Name 'Class' -Value $className
        Set-ItemProperty -Path $verKey -Name 'Assembly' -Value $script:AssemblyName
        Set-ItemProperty -Path $verKey -Name 'RuntimeVersion' -Value $script:RuntimeVersion
        Set-ItemProperty -Path $verKey -Name 'CodeBase' -Value $script:codeBase

        $null = New-Item -Path "$clsidKey\ProgId" -Value $progId -Force
        $null = New-Item -Path "$clsidKey\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}" -Force
    }

    Register-ProgId $ProgIdConnect $ClsidConnect $ClassConnect
    Register-ProgId $ProgIdControl $ClsidControl $ClassControl
    foreach ($clsidRoot in @('HKCU:\Software\Classes\CLSID', 'HKCU:\Software\Classes\WOW6432Node\CLSID')) {
        Register-Clsid $clsidRoot $ClsidConnect $ProgIdConnect $ClassConnect
        Register-Clsid $clsidRoot $ClsidControl $ProgIdControl $ClassControl
    }
    Log (T 'COM classes registered (HKCU, both the 64-bit and the 32-bit registry views).' `
           'COM-классы зарегистрированы (HKCU, 64-битный и 32-битный виды реестра).')

    # --- Ключ надстройки Word (путь БЕЗ номера версии Office) ---
    $addinKey = "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgIdConnect"
    $null = New-Item -Path $addinKey -Force
    Set-ItemProperty -Path $addinKey -Name 'FriendlyName' -Value $FriendlyName
    Set-ItemProperty -Path $addinKey -Name 'Description' -Value $Description
    Set-ItemProperty -Path $addinKey -Name 'LoadBehavior' -Value 3 -Type DWord
    Log (T 'The add-in is registered with Word (LoadBehavior=3).' `
           'Надстройка зарегистрирована в Word (LoadBehavior=3).')

    # --- Снять возможный «бан» от предыдущих сбоев (Office 2013/2016+/365) ---
    # Записи DisabledItems двоичные; внутри в UTF-16 лежат путь DLL и имя надстройки.
    # Удаляем ТОЛЬКО записи с упоминанием TabsForWord: чужие отключённые надстройки
    # Word отключил не просто так, включать их обратно мы не вправе.
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
                    Log ((T 'Cleared the Word ban on our add-in ({0}, entry {1}).' `
                            'Снят бан Word с нашей надстройки ({0}, запись {1}).') -f $officeVer, $valueName)
                }
            }
        } finally {
            $key.Close()
        }
    }

    # --- Проверка: читаем обратно ключевые записи ---
    $ok = $true
    foreach ($check in @(
        @{ Name = (T 'CLSID (64-bit view)' 'CLSID (64-bit вид)'); Path = "HKCU:\Software\Classes\CLSID\$ClsidConnect\InprocServer32" },
        @{ Name = (T 'CLSID (32-bit view)' 'CLSID (32-bit вид)'); Path = "HKCU:\Software\Classes\WOW6432Node\CLSID\$ClsidConnect\InprocServer32" },
        @{ Name = (T 'Word add-in key' 'Ключ надстройки Word'); Path = $addinKey }
    )) {
        if (Test-Path $check.Path) {
            Log ((T 'Check: {0} - OK' 'Проверка: {0} - OK') -f $check.Name)
        } else {
            Log ((T 'Check: {0} - NOT FOUND ({1})' 'Проверка: {0} - НЕ НАЙДЕН ({1})') -f $check.Name, $check.Path) 'Red'
            $ok = $false
        }
    }
    $lb = (Get-ItemProperty -Path $addinKey -Name LoadBehavior -ErrorAction SilentlyContinue).LoadBehavior
    Log ((T 'Check: LoadBehavior = {0} (must be 3)' 'Проверка: LoadBehavior = {0} (должно быть 3)') -f $lb)
    if (-not $ok -or $lb -ne 3) {
        Log (T 'INSTALLATION FINISHED WITH WARNINGS (see above). Run Diagnostics.cmd and send the report.' `
               'УСТАНОВКА ЗАВЕРШИЛАСЬ С ЗАМЕЧАНИЯМИ (см. выше). Запустите Diagnostics.cmd и пришлите отчёт.') 'Yellow'
        return 1
    }

    Log ''
    Log (T 'INSTALLATION COMPLETED SUCCESSFULLY.' 'УСТАНОВКА ЗАВЕРШЕНА УСПЕШНО.') 'Green'
    Log (T 'Start Microsoft Word - a tab bar will appear at the top of every document window.' `
           'Запустите Microsoft Word - сверху каждого окна документа появится панель вкладок.')
    Log (T 'If it does not: run Diagnostics.cmd and send the report file from your desktop.' `
           'Если панель не появилась: запустите Diagnostics.cmd и пришлите файл отчёта с рабочего стола.')
    return 0
}

$exitCode = 1
try {
    $exitCode = Invoke-Install
} catch {
    Log ((T 'INSTALLATION ERROR: {0}' 'ОШИБКА УСТАНОВКИ: {0}') -f $_.Exception.Message) 'Red'
    Log ($_.ScriptStackTrace)
    Log (T 'Run Diagnostics.cmd and send the report file from your desktop.' `
           'Запустите Diagnostics.cmd и пришлите файл отчёта с рабочего стола.') 'Yellow'
    $exitCode = 1
} finally {
    Save-InstallLog ($exitCode -ne 0)
}
exit $exitCode
