@echo off
rem Update from GitHub now instead of waiting for 3 AM (restarts the dashboard only if there is something new).
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0update_windows.ps1"
pause
