@echo off
REM Double-click to install the UpslideClone add-ins (Excel / PowerPoint / Word).
REM Builds Release, copies to a stable location, trusts the cert, registers, and
REM verifies each add-in actually loads. No admin rights needed.
echo Installing UpslideClone add-ins... this builds + verifies, give it a minute.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-Upslide.ps1" %*
echo.
pause
