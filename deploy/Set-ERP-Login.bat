@echo off
cd /d "%~dp0"
if exist "..\venv\Scripts\python.exe" (set PY="..\venv\Scripts\python.exe") else (set PY=python)
%PY% set_erp_login.py
pause
