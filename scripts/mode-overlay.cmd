@echo off
rem SLUZHEBNYY skript (dlya razrabotki i diagnostiki, v postavku ne vhodit):
rem native-rezhim BEZ rezervirovaniya mesta - panel lozhitsya POVERH verhnego
rem kraya oblasti dokumenta (zakryvaet lineyku). Obychnym polzovatelyam nuzhny
rem installer\"Vkladki v okne Word.cmd" / installer\"Klassicheskiy rezhim.cmd".
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
