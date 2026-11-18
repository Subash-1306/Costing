@echo off
rem Phase 2: open the dashboard over Tailscale (phones and laptops anywhere). Needs the Tailscale app.
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup_tailscale.ps1"
pause
