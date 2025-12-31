#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel表格合并工具 - EXE打包脚本
使用PyInstaller将GUI程序打包成Windows可执行文件
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path


def check_pyinstaller():
    """检查PyInstaller是否已安装"""
    try:
        import PyInstaller
        print("✓ PyInstaller 已安装")
        return True
    except ImportError:
        print("⚠️  PyInstaller 未安装")
        return False


def install_pyinstaller():
    """安装PyInstaller"""
    print("正在安装 PyInstaller...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✓ PyInstaller 安装成功")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ 安装失败: {e}")
        return False


def create_spec_file():
    """创建PyInstaller配置文件"""
    spec_content = """# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['excel_merger_gui.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['openpyxl', 'pandas', 'tkinter'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Excel表格合并工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 不显示控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 可以添加图标文件路径
)
"""
    
    with open('excel_merger_gui.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    print("✓ 已创建配置文件: excel_merger_gui.spec")


def build_exe():
    """构建EXE文件"""
    print("\n" + "=" * 60)
    print("开始打包...")
    print("=" * 60 + "\n")
    
    try:
        # 使用spec文件打包
        subprocess.check_call([
            sys.executable,
            "-m",
            "PyInstaller",
            "excel_merger_gui.spec",
            "--clean"
        ])
        print("\n✓ 打包成功！")
        return True
    except subprocess.CalledProcessError as e:
        print(f"\n❌ 打包失败: {e}")
        return False


def cleanup():
    """清理临时文件"""
    print("\n清理临时文件...")
    
    # 删除build目录
    if os.path.exists('build'):
        try:
            shutil.rmtree('build')
            print("✓ 已删除 build 目录")
        except Exception as e:
            print(f"⚠️  无法删除 build 目录: {e}")
    
    # 删除__pycache__
    if os.path.exists('__pycache__'):
        try:
            shutil.rmtree('__pycache__')
            print("✓ 已删除 __pycache__ 目录")
        except Exception as e:
            print(f"⚠️  无法删除 __pycache__ 目录: {e}")


def create_readme_for_dist():
    """为dist目录创建说明文件"""
    readme_content = """# Excel表格合并工具 - 可执行文件版本

## 使用方法

1. 双击运行 "Excel表格合并工具.exe"
2. 点击"➕ 选择Excel文件"按钮，选择要合并的多个Excel文件
3. 输入输出文件名（或点击"📂 浏览"选择位置）
4. 点击"✨ 开始合并"按钮
5. 查看操作日志和结果

## 注意事项

- 所有输入的Excel文件必须具有相同的表头
- 第一列会被自动重新编号（从1开始）
- 支持 .xlsx 和 .xls 格式的Excel文件
- 如果遇到问题，请查看操作日志中的错误信息

## 系统要求

- Windows 7/8/10/11
- 无需安装Python环境
- 约100MB磁盘空间

## 技术支持

如有问题，请检查：
1. 文件是否被其他程序占用
2. 输出目录是否有写入权限
3. Excel文件格式是否正确
"""
    
    dist_path = Path('dist')
    if dist_path.exists():
        readme_path = dist_path / '使用说明.txt'
        with open(readme_path, 'w', encoding='utf-8') as f:
            f.write(readme_content)
        print(f"✓ 已创建使用说明: {readme_path}")


def main():
    """主函数"""
    print("=" * 60)
    print("📦 Excel表格合并工具 - EXE打包程序")
    print("=" * 60)
    print()
    
    # 检查当前目录
    if not os.path.exists('excel_merger_gui.py'):
        print("❌ 错误: 找不到 excel_merger_gui.py 文件")
        print("   请在项目根目录下运行此脚本")
        sys.exit(1)
    
    # 检查并安装PyInstaller
    if not check_pyinstaller():
        print("\n需要安装 PyInstaller 才能打包程序")
        response = input("是否现在安装? (y/n): ")
        if response.lower() == 'y':
            if not install_pyinstaller():
                sys.exit(1)
        else:
            print("取消打包")
            sys.exit(0)
    
    print()
    
    # 创建spec文件
    create_spec_file()
    
    print()
    print("配置说明:")
    print("  - 程序名称: Excel表格合并工具.exe")
    print("  - 打包模式: 单文件模式（所有依赖打包到一个exe中）")
    print("  - 控制台窗口: 隐藏")
    print("  - UPX压缩: 启用（减小文件体积）")
    print()
    
    # 询问是否继续
    response = input("是否开始打包? (y/n): ")
    if response.lower() != 'y':
        print("取消打包")
        sys.exit(0)
    
    # 开始打包
    if build_exe():
        # 创建说明文件
        create_readme_for_dist()
        
        # 清理临时文件
        cleanup()
        
        # 显示结果
        print("\n" + "=" * 60)
        print("✨ 打包完成！")
        print("=" * 60)
        print()
        print("可执行文件位置:")
        exe_path = Path('dist') / 'Excel表格合并工具.exe'
        if exe_path.exists():
            print(f"  📁 {exe_path.absolute()}")
            file_size = exe_path.stat().st_size / (1024 * 1024)
            print(f"  📊 文件大小: {file_size:.2f} MB")
        else:
            print("  ⚠️  未找到生成的exe文件，请检查dist目录")
        
        print()
        print("下一步:")
        print("  1. 进入 dist 目录")
        print("  2. 双击运行 'Excel表格合并工具.exe'")
        print("  3. 可以将整个 dist 目录分发给其他用户")
        print()
    else:
        print("\n打包失败，请检查错误信息")
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n用户取消操作")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ 发生错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
