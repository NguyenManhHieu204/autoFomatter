@echo off
setlocal
cd /d "%~dp0"
python --version >nul 2>&1
if %ERRORLEVEL% equ 0 (
    start "" pythonw main.py
    exit /b
)
if exist "%~dp0dist\AutoFormatter\AutoFormatter.exe" (
    start "" "%~dp0dist\AutoFormatter\AutoFormatter.exe"
    exit /b
)
start "" python main.py
