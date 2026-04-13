@echo off
chcp 65001 >nul
title Forecast System
cd /d "%~dp0"

echo ================================================
echo   Forecast System - Starting...
echo   Log: startup_log.txt
echo ================================================
echo.

:: Write log header
echo [%date% %time%] Forecast System starting > startup_log.txt
echo Working dir: %cd% >> startup_log.txt

:: ── Check Python ──────────────────────────────────
python --version >> startup_log.txt 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found.
    echo         Install Python 3.8+ from https://www.python.org
    echo         Check "Add Python to PATH" during install.
    echo [ERROR] Python not found >> startup_log.txt
    goto :end
)
for /f "tokens=*" %%v in ('python --version 2^>^&1') do echo Python: %%v
echo.

:: ── Install packages ──────────────────────────────
echo Installing / checking packages...
python -m pip install -r requirements.txt --disable-pip-version-check >> startup_log.txt 2>&1
if %errorlevel% neq 0 (
    echo Trying user-mode install...
    python -m pip install -r requirements.txt --user --disable-pip-version-check >> startup_log.txt 2>&1
    if %errorlevel% neq 0 (
        echo [ERROR] pip install failed. See startup_log.txt for details.
        echo         Or run manually: python -m pip install flask openpyxl
        goto :end
    )
)
echo Packages installed.
echo.

:: ── Verify flask is actually importable ───────────
python -c "import flask" >> startup_log.txt 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] flask installed but cannot be imported.
    echo         This usually means pip and python point to different versions.
    echo         Try: py -m pip install flask openpyxl
    echo [ERROR] flask import failed >> startup_log.txt
    goto :end
)
echo Flask OK.
echo.

:: ── Kill old process on port 8000 ─────────────────
for /f "tokens=5" %%a in ('netstat -ano 2^>nul ^| findstr /R "[:.]8000 " ^| findstr "LISTENING"') do (
    echo Killing old process on port 8000 ^(PID: %%a^)...
    taskkill /F /PID %%a >nul 2>&1
    timeout /t 1 /nobreak >nul
)

:: ── Start server ──────────────────────────────────
echo Server starting...
echo If browser does not open, visit: http://localhost:8000
echo Do NOT close this window.
echo ================================================
echo.

python server.py >> startup_log.txt 2>&1

echo.
echo [Server stopped. Check startup_log.txt for details.]

:end
echo.
echo Check startup_log.txt in this folder if there were errors.
echo.
pause
