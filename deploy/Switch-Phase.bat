@echo off
rem Switch between Phase 1 (before Tailscale and photos) and Phase 2 (now).
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0switch_phase.ps1" %*
pause
