@echo off
rem Safely copies the .db files from another folder into this install (close the dashboard first).
cd /d "%~dp0"
if exist "..\venv\Scripts\python.exe" (set PY="..\venv\Scripts\python.exe") else (set PY=python)
%PY% copy_databases.py
pause
