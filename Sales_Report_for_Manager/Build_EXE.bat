@echo off
REM ============================================================
REM  Builds "Sales Report.exe" from Monthly_Sales_Report.py.
REM  Run this ONCE on a Windows PC that has Python installed.
REM  The exe it makes runs on any Windows PC - no Python needed.
REM ============================================================
cd /d "%~dp0"
set PY=python
where py >nul 2>nul && set PY=py
echo Installing the build tools (first time only)...
%PY% -m pip install --upgrade pyinstaller requests xlrd openpyxl
if errorlevel 1 goto fail
echo.
echo Building Sales Report.exe (takes a minute or two)...
%PY% -m PyInstaller --onefile --clean --noconfirm --name "Sales Report" --exclude-module pandas --exclude-module numpy Monthly_Sales_Report.py
if errorlevel 1 goto fail
copy /y "dist\Sales Report.exe" "Sales Report.exe" >nul
rmdir /s /q build 2>nul
del /q "Sales Report.spec" 2>nul
echo.
echo ============================================================
echo  Done: "Sales Report.exe" is in this folder.
echo  Give that one file to your manager.
echo ============================================================
pause
exit /b 0
:fail
echo.
echo Build failed - send Claude the messages above.
pause
exit /b 1
