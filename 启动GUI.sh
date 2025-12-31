#!/bin/bash
# Excel表格合并工具 - GUI启动脚本

echo "🚀 正在启动Excel表格合并工具..."
echo ""

# 检查Python是否安装
if ! command -v python3 &> /dev/null; then
    echo "❌ 错误: 未找到Python3，请先安装Python3"
    exit 1
fi

# 检查依赖是否安装
echo "📦 检查依赖..."
python3 -c "import pandas, openpyxl, tkinter" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "⚠️  检测到缺少依赖，正在安装..."
    pip install -q pandas openpyxl 2>/dev/null
    
    # 检查tkinter
    python3 -c "import tkinter" 2>/dev/null
    if [ $? -ne 0 ]; then
        echo "ℹ️  需要安装图形界面支持(tkinter)..."
        echo "    请运行: sudo apt-get install python3-tk"
        exit 1
    fi
fi

echo "✓ 依赖检查完成"
echo ""
echo "📊 启动应用程序..."
python3 excel_merger_gui.py
