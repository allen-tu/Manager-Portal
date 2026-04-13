@echo off
chcp 65001 >nul
title Forecast System

cd /d "%~dp0"

echo ================================================
echo   Forecast System - Starting...
echo ================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found.
    echo         Please install Python 3.8+ from https://www.python.org
    echo         Make sure to check "Add Python to PATH" during install.
    pause
    exit /b 1
)

echo Python version:
python --version
echo.

echo Installing required packages...
python -m pip install -r requirements.txt --quiet --disable-pip-version-check
if errorlevel 1 (
    echo [ERROR] Failed to install packages.
    echo         Please check your network connection, or run manually:
    echo         python -m pip install flask openpyxl
    pause
    exit /b 1
)
echo Packages OK
echo.

echo Checking port 8000...
for /f "tokens=5" %%a in ('netstat -ano 2^>nul ^| findstr /R "[:.]8000 " ^| findstr "LISTENING"') do (
    echo   Killing old process (PID: %%a)...
    taskkill /F /PID %%a >nul 2>&1
    timeout /t 1 /nobreak >nul
)

echo Starting server... Browser will open automatically.
echo If browser does not open, go to: http://localhost:8000
echo.
echo Do NOT close this window.
echo ================================================
echo.

python server.py

pause
