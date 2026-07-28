# package.ps1 - builds the installation package into release/.
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File scripts\package.ps1
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$dll = Join-Path $repoRoot 'src\TabsForWord\bin\Release\TabsForWord.dll'
if (-not (Test-Path $dll)) {
    Write-Host 'ERROR: no Release build. Run scripts\build.ps1 first.' -ForegroundColor Red
    exit 1
}

$releaseDir = Join-Path $repoRoot 'release'
$pkgDir = Join-Path $releaseDir 'TabsForWord-Installer'

# Clean and assemble the installation package
if (Test-Path $pkgDir) { Remove-Item $pkgDir -Recurse -Force }
New-Item -ItemType Directory -Force $pkgDir | Out-Null

Copy-Item $dll $pkgDir
# The .cmd names are ASCII: on a non-Russian locale Cyrillic file names read badly
# and get in the way of sending the files by e-mail. The Russian-named
# "Диагностика.cmd" is kept as a wrapper: earlier instructions refer to it.
foreach ($f in @('install.ps1', 'uninstall.ps1', 'repair.ps1', 'diagnose.ps1',
                 'Install.cmd', 'Uninstall.cmd', 'Repair.cmd', 'Diagnostics.cmd',
                 'Classic-mode.cmd', 'In-window-tabs.cmd', 'Диагностика.cmd')) {
    Copy-Item (Join-Path $repoRoot "installer\$f") $pkgDir
}
foreach ($doc in @('INSTALL_EN.md', 'INSTALL_RU.md', 'UNINSTALL_EN.md', 'UNINSTALL_RU.md')) {
    Copy-Item (Join-Path $repoRoot "docs\$doc") $pkgDir -ErrorAction SilentlyContinue
}
# The licence travels with the binary: what a person downloads has to say what
# they are allowed to do with it. Missing it is an error, not something to skip.
Copy-Item (Join-Path $repoRoot 'LICENSE') (Join-Path $pkgDir 'LICENSE.txt')

# Package README: English (README.txt - the one everybody can read) and Russian
Set-Content -Path (Join-Path $pkgDir 'README.txt') -Encoding UTF8 -Value @'
Tabs for Microsoft Word
=======================
A tab bar for the documents you have open, inside the Word window itself.
(По-русски: файл ПРОЧТИ_МЕНЯ.txt рядом.)

WHAT YOU NEED:
  - Windows 10 version 1903 (May 2019) or newer, or Windows 11 (64-bit);
    .NET Framework 4.8 is already part of those - nothing extra to install;
  - desktop Microsoft Word for Windows (Microsoft 365 / Word 2016 or newer);
    it does not work with Word Online, Word for macOS or Word on mobile;
  - administrator rights are NOT required.
  If something is missing, the installer says so and explains what to do.

INSTALL:
  1. Close Microsoft Word.
  2. Double-click Install.cmd.
  3. Start Word.

UNINSTALL:     Uninstall.cmd
IF IT BREAKS:  Repair.cmd (then restart Word)

IF IT DOES NOT WORK:
  Run Diagnostics.cmd - a ZIP report appears on your desktop. Send it to
  whoever gave you this program. The report contains the names and paths of
  your documents, your computer name and your user name, but not the content
  of any document; the archive holds plain text files you can read first.

FALLBACK LOOK (rarely needed):
  Normally the tab bar is built into the Word window, below the ribbon.
  If it misbehaves on your computer (does not appear, overlaps the document),
  run Classic-mode.cmd - the bar moves into a separate Word strip with a grey
  service title. To go back: In-window-tabs.cmd. Restart Word after either.

LANGUAGE:
  The add-in follows Word's own interface language (Russian Word - Russian
  tabs, anything else - English). To choose it by hand: the small arrow on the
  right of the bar -> "Settings..." -> "Language". The installer and the
  diagnostics follow the Windows display language; to force one:
    powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1 -Lang en

Full guide: INSTALL_EN.md
No administrator rights are required.
Licence: MIT (see LICENSE.txt) - free to use and to pass on, no warranty.
'@

# The Russian package README
Set-Content -Path (Join-Path $pkgDir 'ПРОЧТИ_МЕНЯ.txt') -Encoding UTF8 -Value @'
Вкладки для Microsoft Word
==========================
Панель вкладок открытых документов прямо в окне Word.
(In English: see README.txt next to this file.)

ТРЕБОВАНИЯ К КОМПЬЮТЕРУ:
  - Windows 10 версии 1903 (май 2019) или новее, либо Windows 11 (64-bit);
    на них .NET Framework 4.8 уже встроен — доустанавливать ничего не нужно;
  - настольный Microsoft Word для Windows (Microsoft 365 / Word 2016 и новее);
    не работает с Word Online (в браузере), Word для macOS и мобильным Word;
  - права администратора НЕ требуются.
  Если чего-то не хватает, установщик прямо об этом скажет и подскажет, что делать.

УСТАНОВКА:
  1. Закройте Microsoft Word.
  2. Дважды щёлкните файл Install.cmd.
  3. Запустите Word.

УДАЛЕНИЕ:      Uninstall.cmd
ЕСЛИ СЛОМАЛОСЬ: Repair.cmd (потом перезапустите Word)

ЕСЛИ НЕ РАБОТАЕТ:
  Запустите Diagnostics.cmd (то же самое — файл «Диагностика.cmd») — на
  рабочем столе появится ZIP-файл с отчётом. Пришлите его тому, кто дал
  вам эту программу.

ЗАПАСНОЙ ВИД ПАНЕЛИ (нужен редко):
  Обычно панель вкладок встроена прямо в окно Word, под ленту.
  Если на вашем компьютере она ведёт себя странно (не появляется,
  налезает на документ), запустите Classic-mode.cmd — панель переедет
  в отдельную полосу Word со служебным заголовком. Вернуть обычный вид:
  In-window-tabs.cmd. После смены — перезапустите Word.

ЯЗЫК:
  Надстройка сама берёт язык у Word (русский Word — русские вкладки,
  любой другой — английские). Поменять вручную: кнопка ▾ на панели →
  «Настройки…» → «Язык». Установщик и Диагностика говорят на языке
  Windows; можно задать явно, например:
    powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1 -Lang en

Подробная инструкция: INSTALL_RU.md (English: INSTALL_EN.md)
Права администратора НЕ требуются.
Лицензия: MIT (файл LICENSE.txt) — можно свободно пользоваться и передавать
дальше, без каких-либо гарантий.
'@

# ZIP
$zip = Join-Path $releaseDir 'TabsForWord.zip'
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path $pkgDir -DestinationPath $zip

Write-Host "Package: $pkgDir" -ForegroundColor Green
Write-Host "ZIP:   $zip" -ForegroundColor Green
exit 0
