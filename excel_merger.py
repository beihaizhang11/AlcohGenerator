#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel表格合并工具
功能：
1. 合并多个具有相同表头的Excel文件
2. 自动重新生成第一列的递增序号（从1开始）
"""

import os
import sys
import pandas as pd
from pathlib import Path


class ExcelMerger:
    """Excel文件合并器"""
    
    def __init__(self):
        self.dataframes = []
        self.header = None
    
    def load_excel_files(self, file_paths):
        """
        加载多个Excel文件
        
        Args:
            file_paths: Excel文件路径列表
        
        Returns:
            bool: 是否成功加载
        """
        if not file_paths:
            print("❌ 错误：没有提供Excel文件")
            return False
        
        print(f"\n📂 开始加载 {len(file_paths)} 个Excel文件...")
        
        for i, file_path in enumerate(file_paths, 1):
            try:
                # 检查文件是否存在
                if not os.path.exists(file_path):
                    print(f"❌ 文件不存在: {file_path}")
                    return False
                
                # 读取Excel文件
                df = pd.read_excel(file_path)
                
                # 检查是否为空
                if df.empty:
                    print(f"⚠️  警告：文件 {file_path} 是空的，跳过")
                    continue
                
                # 检查表头是否一致
                if self.header is None:
                    self.header = list(df.columns)
                    print(f"✓ 表头：{self.header}")
                else:
                    if list(df.columns) != self.header:
                        print(f"❌ 错误：文件 {file_path} 的表头与第一个文件不一致")
                        print(f"   预期：{self.header}")
                        print(f"   实际：{list(df.columns)}")
                        return False
                
                self.dataframes.append(df)
                print(f"✓ 已加载文件 {i}/{len(file_paths)}: {os.path.basename(file_path)} ({len(df)} 行)")
                
            except Exception as e:
                print(f"❌ 读取文件 {file_path} 时出错: {str(e)}")
                return False
        
        if not self.dataframes:
            print("❌ 错误：没有成功加载任何文件")
            return False
        
        return True
    
    def merge_and_reindex(self):
        """
        合并所有数据并重新生成序号
        
        Returns:
            pd.DataFrame: 合并后的数据框
        """
        if not self.dataframes:
            print("❌ 错误：没有数据可以合并")
            return None
        
        print("\n🔄 开始合并数据...")
        
        # 合并所有数据框
        merged_df = pd.concat(self.dataframes, ignore_index=True)
        total_rows = len(merged_df)
        print(f"✓ 已合并 {len(self.dataframes)} 个文件，共 {total_rows} 行数据")
        
        # 重新生成第一列的序号（从1开始）
        if len(merged_df.columns) > 0:
            first_column = merged_df.columns[0]
            merged_df[first_column] = range(1, total_rows + 1)
            print(f"✓ 已重新生成第一列序号：从 1 到 {total_rows}")
        
        return merged_df
    
    def save_to_excel(self, output_path, merged_df):
        """
        保存合并后的数据到Excel文件
        
        Args:
            output_path: 输出文件路径
            merged_df: 合并后的数据框
        
        Returns:
            bool: 是否保存成功
        """
        try:
            print(f"\n💾 正在保存到文件: {output_path}")
            merged_df.to_excel(output_path, index=False, engine='openpyxl')
            print(f"✅ 成功！文件已保存")
            return True
        except Exception as e:
            print(f"❌ 保存文件时出错: {str(e)}")
            return False


def main():
    """主函数"""
    print("=" * 60)
    print("📊 Excel表格合并工具")
    print("=" * 60)
    
    # 检查命令行参数
    if len(sys.argv) < 3:
        print("\n使用方法:")
        print("  python excel_merger.py <输出文件名> <输入文件1> <输入文件2> [输入文件3] ...")
        print("\n示例:")
        print("  python excel_merger.py merged.xlsx file1.xlsx file2.xlsx file3.xlsx")
        print("\n说明:")
        print("  - 所有输入文件必须有相同的表头")
        print("  - 第一列将自动重新编号（从1开始递增）")
        sys.exit(1)
    
    output_file = sys.argv[1]
    input_files = sys.argv[2:]
    
    print(f"\n输出文件: {output_file}")
    print(f"输入文件: {len(input_files)} 个")
    
    # 创建合并器实例
    merger = ExcelMerger()
    
    # 加载文件
    if not merger.load_excel_files(input_files):
        sys.exit(1)
    
    # 合并并重新索引
    merged_df = merger.merge_and_reindex()
    if merged_df is None:
        sys.exit(1)
    
    # 保存结果
    if not merger.save_to_excel(output_file, merged_df):
        sys.exit(1)
    
    print("\n" + "=" * 60)
    print("✨ 全部完成！")
    print("=" * 60)


if __name__ == "__main__":
    main()
