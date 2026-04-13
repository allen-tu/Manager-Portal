@echo off
chcp 65001 >nul
title 銷售案 Forecast 分析系統

:: 第一步：切換到腳本所在目錄（必須最先執行）
cd /d "%~dp0"

echo ================================================
echo   銷售案 Forecast 分析系統  啟動中...
echo ================================================
echo.

:: 確認 Python 是否安裝
python --version >nul 2>&1
if errorlevel 1 (
    echo [錯誤] 找不到 Python，請先至 https://www.python.org 安裝 Python 3.8 以上版本
    echo        安裝時請勾選 "Add Python to PATH"
    pause
    exit /b 1
)

echo Python 版本：
python --version
echo.

:: 安裝必要套件（使用 python -m pip 確保對應同一個 Python）
echo 檢查並安裝必要套件...
python -m pip install -r requirements.txt --quiet --disable-pip-version-check
if errorlevel 1 (
    echo [錯誤] 套件安裝失敗，請確認網路連線或手動執行：
    echo        python -m pip install flask openpyxl
    pause
    exit /b 1
)
echo 套件檢查完成
echo.

:: 清除佔用 port 8000 的舊進程
echo 檢查 port 8000 是否已被佔用...
for /f "tokens=5" %%a in ('netstat -ano 2^>nul ^| findstr /R "[:.]8000 " ^| findstr "LISTENING"') do (
    echo   偵測到舊進程 (PID: %%a)，正在關閉...
    taskkill /F /PID %%a >nul 2>&1
    echo   舊進程已清除
    timeout /t 1 /nobreak >nul
)

echo 系統啟動中，瀏覽器將自動開啟...
echo 若瀏覽器未自動開啟，請手動前往：http://localhost:8000
echo.
echo 請勿關閉此視窗（關閉後系統停止運作）
echo ================================================
echo.

:: 啟動 Flask server
python server.py

pause
