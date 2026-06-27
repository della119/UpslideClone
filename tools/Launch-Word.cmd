@echo off
REM Optional clean-restart helper. With the PRODUCTION install
REM (installer\Install-Production.ps1) the Upslide tab loads on a normal launch,
REM so you usually don't need this. Use it only to force a fully clean restart if
REM the tab is ever missing (Office "fast restart" reused a stale process).
echo Closing any running Word...
taskkill /IM WINWORD.EXE /F >nul 2>&1
ping -n 6 127.0.0.1 >nul
echo Starting Word...
start "" winword.exe
exit
