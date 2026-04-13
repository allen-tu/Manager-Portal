@echo off
chcp 65001 >nul
title Forecast System

cd /d "%~dp0"

echo ================================================
echo   Forecast System - Starting...
echo ================================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found.
    echo         Install Python 3.8+ from https://www.python.org
    echo         Check "Add Python to PATH" during install.
    echo.
    cmd /k
    exit /b
)

echo Python:
python --version
echo.

:: Try normal install first
echo Installing packages...
python -m pip install -r requirements.txt --disable-pip-version-check
if errorlevel 1 (
    echo.
    echo Normal install failed. Trying user-mode install...
    python -m pip install -r requirements.txt --user --disable-pip-version-check
    if errorlevel 1 (
        echo.
        echo [ERROR] Package install failed.
        echo         Try running this window as Administrator,
        echo         or manually run: python -m pip install flask openpyxl
        echo.
        cmd /k
        exit /b
    )
)
echo.
echo Packages OK.
echo.

:: Kill old process on port 8000
for /f "tokens=5" %%a in ('netstat -ano 2^>nul ^| findstr /R "[:.]8000 " ^| findstr "LISTENING"') do (
    echo Killing old process on port 8000 (PID: %%a)...
    taskkill /F /PID %%a >nul 2>&1
    timeout /t 1 /nobreak >nul
)

echo Server starting... Browser will open automatically.
echo If browser does not open, visit: http://localhost:8000
echo.
echo Do NOT close this window.
echo ================================================
echo.

python server.py

echo.
echo [Server stopped]
cmd /k
