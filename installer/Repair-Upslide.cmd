@echo off
REM Double-click if the Upslide tab ever disappears.
REM Re-registers + clears Office's soft-disable + verifies — WITHOUT rebuilding.
echo Repairing UpslideClone registration (no rebuild)...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-Upslide.ps1" -SkipBuild %*
echo.
pause
