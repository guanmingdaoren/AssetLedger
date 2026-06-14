@echo off
setlocal
cd /d "%~dp0"
python packaging\build_windows.py
if errorlevel 1 exit /b 1
endlocal
