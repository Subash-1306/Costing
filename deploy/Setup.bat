@echo off
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup_windows.ps1"
if errorlevel 1 (echo. & echo Setup did not finish. See the message above and logs\setup.log.)
pause
