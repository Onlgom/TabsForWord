@echo off
rem Pereklyuchit panel vkladok v KLASSICHESKIY rezhim (Custom Task Pane):
rem panel s sluzhebnoy seroy polosoy zagolovka, kak do versii 1.6.0.
rem Switch the tab strip back to the classic Custom Task Pane host.
set CFG=%LOCALAPPDATA%\TabsForWord\native-host.cfg
if not exist "%LOCALAPPDATA%\TabsForWord" mkdir "%LOCALAPPDATA%\TabsForWord"
(
echo mode=ctp
) > "%CFG%"
echo.
echo [OK] Vklyuchen KLASSICHESKIY rezhim (mode=ctp).
echo Perezapustite Microsoft Word.
echo.
echo Vernut obychnyy vid (vkladki vnutri okna Word):
echo   fayl "In-window-tabs.cmd" iz etoy zhe papki.
echo.
pause
