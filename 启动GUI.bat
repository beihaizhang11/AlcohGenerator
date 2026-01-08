@echo off
chcp 65001 >nul
echo 🚀 正在启动Excel表格合并工具...
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 错误: 未找到Python，请先安装Python3
    pause
    exit /b 1
)

echo 📦 检查依赖...
python -c "import pandas, openpyxl, tkinter" >nul 2>&1
if errorlevel 1 (
    echo ⚠️  检测到缺少依赖，正在安装...
    pip install -q pandas openpyxl
)

python -c "import tkinterdnd2" >nul 2>&1
if errorlevel 1 (
    echo ⚠️  安装拖拽支持库...
    pip install -q tkinterdnd2
)

echo ✓ 依赖检查完成
echo.
echo 📊 启动应用程序...
python excel_merger_gui.py

pause
