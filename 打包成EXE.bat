@echo off
chcp 65001 >nul
title Excel表格合并工具 - 打包程序

echo ========================================
echo  📦 Excel表格合并工具 - EXE打包程序
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 错误: 未找到Python
    echo    请先安装Python 3.6或更高版本
    echo.
    pause
    exit /b 1
)

echo ✓ Python环境检测通过
echo.

REM 运行打包脚本
python build_exe.py

echo.
pause
