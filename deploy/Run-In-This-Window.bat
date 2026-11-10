@echo off
rem Runs the dashboard in this window (for testing on a laptop). Close the window to stop it.
cd /d "%~dp0..\DASHBOARD"
if exist "..\venv\Scripts\python.exe" (set PY="..\venv\Scripts\python.exe") else (set PY=python)
%PY% ProjectDashboard.py
pause
