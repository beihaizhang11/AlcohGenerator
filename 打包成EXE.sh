#!/bin/bash
# Excel表格合并工具 - 打包程序（Linux/Mac版本）

echo "========================================"
echo " 📦 Excel表格合并工具 - EXE打包程序"
echo "========================================"
echo ""

# 检查Python是否安装
if ! command -v python3 &> /dev/null; then
    echo "❌ 错误: 未找到Python3"
    echo "   请先安装Python 3.6或更高版本"
    exit 1
fi

echo "✓ Python环境检测通过"
echo ""

# 运行打包脚本
python3 build_exe.py

echo ""
echo "按任意键退出..."
read -n 1
