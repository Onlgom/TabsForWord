@echo off
rem Vernut OBYCHNYY rezhim: panel vkladok vnutri okna Word, bez sluzhebnoy
rem polosy Custom Task Pane. Eto rezhim po umolchaniyu s versii 1.6.0.
rem Restore the default native in-window tab strip.
set CFG=%LOCALAPPDATA%\TabsForWord\native-host.cfg
if not exist "%LOCALAPPDATA%\TabsForWord" mkdir "%LOCALAPPDATA%\TabsForWord"
(
echo mode=native
echo reserve=1
) > "%CFG%"
echo.
echo [OK] Vklyuchen obychnyy rezhim (mode=native, reserve=1).
echo Perezapustite Microsoft Word.
echo.
pause
