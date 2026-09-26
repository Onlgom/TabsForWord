@echo off
rem SLUZHEBNYY skript (dlya razrabotki i diagnostiki, v postavku ne vhodit):
rem native-rezhim BEZ rezervirovaniya mesta - panel lozhitsya POVERH verhnego
rem kraya oblasti dokumenta (zakryvaet lineyku). Obychnym polzovatelyam nuzhny
rem installer\In-window-tabs.cmd / installer\Classic-mode.cmd.
set CFG=%LOCALAPPDATA%\TabsForWord\native-host.cfg
if not exist "%LOCALAPPDATA%\TabsForWord" mkdir "%LOCALAPPDATA%\TabsForWord"
(
echo mode=native
echo reserve=0
) > "%CFG%"
echo.
echo [OK] mode=native, reserve=0 (overlay). Perezapustite Word.
echo.
pause
