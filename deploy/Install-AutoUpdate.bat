@echo off
rem One-time: this PC updates the costing dashboard from GitHub every night at 3 AM. Needs Git for Windows (git-scm.com).
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_autoupdate.ps1"
pause
