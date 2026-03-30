#!/bin/bash
# 銷售案 Forecast 分析系統 — Mac 啟動腳本
# 雙擊此檔案即可啟動

# 切換到腳本所在目錄
cd "$(dirname "$0")"

echo "================================================"
echo "  銷售案 Forecast 分析系統  啟動中..."
echo "================================================"
echo ""

# 清除舊的 server.py 進程（避免 port 8000 被佔用）
if pgrep -f server.py > /dev/null 2>&1; then
    echo "⚠️  偵測到舊的 server.py 正在執行，正在關閉..."
    pkill -f server.py
    sleep 1
    echo "✅ 舊進程已清除"
    echo ""
fi

# 確認 Python3 是否安裝
if ! command -v python3 &>/dev/null; then
    echo "[錯誤] 找不到 Python3"
    echo "       請開啟 Terminal 執行：brew install python3"
    echo "       或至 https://www.python.org 下載安裝"
    read -p "按 Enter 關閉..."
    exit 1
fi

echo "Python 版本：$(python3 --version)"
echo ""

# 安裝必要套件（第一次需要網路，之後離線可用）
echo "檢查並安裝必要套件..."
python3 -m pip install flask --quiet --disable-pip-version-check 2>/dev/null || \
python3 -m pip install flask --quiet 2>/dev/null || \
echo "[警告] 套件安裝失敗，嘗試繼續..."

echo ""
echo "系統啟動中，瀏覽器將自動開啟..."
echo "若瀏覽器未自動開啟，請手動前往：http://localhost:8000"
echo ""
echo "請勿關閉此視窗（關閉後系統停止運作）"
echo "================================================"
echo ""

# 啟動 Flask server
python3 server.py
