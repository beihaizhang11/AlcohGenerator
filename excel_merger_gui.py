#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel表格合并工具 - 图形界面版本
功能：
1. 合并多个具有相同表头的Excel文件
2. 自动重新生成第一列的递增序号（从1开始）
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk
import pandas as pd
from pathlib import Path
import threading


class ExcelMergerGUI:
    """Excel合并工具图形界面"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("📊 Excel表格合并工具")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # 设置样式
        style = ttk.Style()
        style.theme_use('clam')
        
        # 存储选择的文件
        self.selected_files = []
        
        # 创建界面
        self.create_widgets()
        
    def create_widgets(self):
        """创建GUI组件"""
        
        # 标题框架
        title_frame = tk.Frame(self.root, bg="#2c3e50", height=60)
        title_frame.pack(fill=tk.X, padx=0, pady=0)
        title_frame.pack_propagate(False)
        
        title_label = tk.Label(
            title_frame,
            text="📊 Excel表格合并工具",
            font=("Arial", 20, "bold"),
            fg="white",
            bg="#2c3e50"
        )
        title_label.pack(pady=15)
        
        # 主容器
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 文件选择区域
        file_frame = tk.LabelFrame(
            main_frame,
            text="📁 选择要合并的Excel文件",
            font=("Arial", 11, "bold"),
            padx=10,
            pady=10
        )
        file_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 按钮框架
        button_frame = tk.Frame(file_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 选择文件按钮
        select_btn = tk.Button(
            button_frame,
            text="➕ 选择Excel文件",
            command=self.select_files,
            bg="#3498db",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=20,
            pady=8,
            cursor="hand2",
            relief=tk.FLAT
        )
        select_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 清空列表按钮
        clear_btn = tk.Button(
            button_frame,
            text="🗑️ 清空列表",
            command=self.clear_files,
            bg="#e74c3c",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=20,
            pady=8,
            cursor="hand2",
            relief=tk.FLAT
        )
        clear_btn.pack(side=tk.LEFT)
        
        # 文件数量标签
        self.file_count_label = tk.Label(
            button_frame,
            text="已选择: 0 个文件",
            font=("Arial", 10),
            fg="#7f8c8d"
        )
        self.file_count_label.pack(side=tk.RIGHT)
        
        # 文件列表框
        list_frame = tk.Frame(file_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # 滚动条
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 列表框
        self.file_listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=("Consolas", 9),
            selectmode=tk.EXTENDED,
            bg="#ecf0f1",
            relief=tk.FLAT
        )
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)
        
        # 输出设置区域
        output_frame = tk.LabelFrame(
            main_frame,
            text="💾 输出设置",
            font=("Arial", 11, "bold"),
            padx=10,
            pady=10
        )
        output_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 输出文件名
        output_label = tk.Label(
            output_frame,
            text="输出文件名:",
            font=("Arial", 10)
        )
        output_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.output_entry = tk.Entry(
            output_frame,
            font=("Arial", 10),
            width=40
        )
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        self.output_entry.insert(0, "merged_result.xlsx")
        
        # 浏览按钮
        browse_btn = tk.Button(
            output_frame,
            text="📂 浏览",
            command=self.browse_output,
            bg="#95a5a6",
            fg="white",
            font=("Arial", 9, "bold"),
            padx=15,
            pady=5,
            cursor="hand2",
            relief=tk.FLAT
        )
        browse_btn.pack(side=tk.LEFT)
        
        # 日志区域
        log_frame = tk.LabelFrame(
            main_frame,
            text="📋 操作日志",
            font=("Arial", 11, "bold"),
            padx=10,
            pady=10
        )
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=8,
            font=("Consolas", 9),
            bg="#2c3e50",
            fg="#ecf0f1",
            relief=tk.FLAT,
            state=tk.DISABLED
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 合并按钮
        merge_btn = tk.Button(
            main_frame,
            text="✨ 开始合并",
            command=self.start_merge,
            bg="#27ae60",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=30,
            pady=12,
            cursor="hand2",
            relief=tk.FLAT
        )
        merge_btn.pack(fill=tk.X)
        
        # 初始日志
        self.log("欢迎使用Excel表格合并工具！")
        self.log("请选择要合并的Excel文件...")
        
    def log(self, message):
        """在日志区域添加消息"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update()
        
    def select_files(self):
        """选择Excel文件"""
        files = filedialog.askopenfilenames(
            title="选择Excel文件",
            filetypes=[
                ("Excel文件", "*.xlsx *.xls"),
                ("所有文件", "*.*")
            ]
        )
        
        if files:
            # 添加新文件到列表（避免重复）
            for file in files:
                if file not in self.selected_files:
                    self.selected_files.append(file)
                    self.file_listbox.insert(tk.END, os.path.basename(file))
            
            # 更新文件数量
            self.update_file_count()
            self.log(f"已添加 {len(files)} 个文件")
            
    def clear_files(self):
        """清空文件列表"""
        self.selected_files.clear()
        self.file_listbox.delete(0, tk.END)
        self.update_file_count()
        self.log("已清空文件列表")
        
    def update_file_count(self):
        """更新文件数量显示"""
        count = len(self.selected_files)
        self.file_count_label.config(text=f"已选择: {count} 个文件")
        
    def browse_output(self):
        """浏览输出位置"""
        file_path = filedialog.asksaveasfilename(
            title="选择输出位置",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        
        if file_path:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, file_path)
            
    def start_merge(self):
        """开始合并（在新线程中执行）"""
        # 验证输入
        if not self.selected_files:
            messagebox.showwarning("提示", "请先选择要合并的Excel文件！")
            return
            
        output_file = self.output_entry.get().strip()
        if not output_file:
            messagebox.showwarning("提示", "请输入输出文件名！")
            return
        
        # 在新线程中执行合并
        thread = threading.Thread(target=self.merge_files, args=(output_file,))
        thread.daemon = True
        thread.start()
        
    def merge_files(self, output_file):
        """执行文件合并"""
        try:
            self.log("\n" + "=" * 50)
            self.log("开始合并操作...")
            self.log("=" * 50)
            
            dataframes = []
            header = None
            
            # 加载所有文件
            self.log(f"\n📂 正在加载 {len(self.selected_files)} 个文件...")
            
            for i, file_path in enumerate(self.selected_files, 1):
                try:
                    # 检查文件是否存在
                    if not os.path.exists(file_path):
                        self.log(f"❌ 文件不存在: {os.path.basename(file_path)}")
                        messagebox.showerror("错误", f"文件不存在:\n{file_path}")
                        return
                    
                    # 读取Excel文件
                    df = pd.read_excel(file_path)
                    
                    # 检查是否为空
                    if df.empty:
                        self.log(f"⚠️  文件为空，跳过: {os.path.basename(file_path)}")
                        continue
                    
                    # 检查表头是否一致
                    if header is None:
                        header = list(df.columns)
                        self.log(f"✓ 表头: {header}")
                    else:
                        if list(df.columns) != header:
                            error_msg = f"文件表头不一致:\n{os.path.basename(file_path)}"
                            self.log(f"❌ {error_msg}")
                            messagebox.showerror("错误", error_msg)
                            return
                    
                    dataframes.append(df)
                    self.log(f"✓ [{i}/{len(self.selected_files)}] {os.path.basename(file_path)} ({len(df)} 行)")
                    
                except Exception as e:
                    self.log(f"❌ 读取文件出错: {os.path.basename(file_path)}")
                    self.log(f"   错误信息: {str(e)}")
                    messagebox.showerror("错误", f"读取文件出错:\n{file_path}\n\n{str(e)}")
                    return
            
            if not dataframes:
                self.log("❌ 没有可用的数据")
                messagebox.showerror("错误", "没有可用的数据可以合并！")
                return
            
            # 合并数据
            self.log("\n🔄 正在合并数据...")
            merged_df = pd.concat(dataframes, ignore_index=True)
            total_rows = len(merged_df)
            self.log(f"✓ 已合并 {len(dataframes)} 个文件，共 {total_rows} 行数据")
            
            # 重新生成序号
            if len(merged_df.columns) > 0:
                first_column = merged_df.columns[0]
                merged_df[first_column] = range(1, total_rows + 1)
                self.log(f"✓ 已重新生成第一列序号: 从 1 到 {total_rows}")
            
            # 保存文件
            self.log(f"\n💾 正在保存到: {output_file}")
            merged_df.to_excel(output_file, index=False, engine='openpyxl')
            self.log("✅ 保存成功！")
            
            self.log("\n" + "=" * 50)
            self.log("✨ 合并完成！")
            self.log("=" * 50 + "\n")
            
            # 显示成功消息
            messagebox.showinfo(
                "成功",
                f"Excel文件合并成功！\n\n"
                f"合并文件数: {len(dataframes)}\n"
                f"总行数: {total_rows}\n"
                f"输出文件: {output_file}"
            )
            
        except Exception as e:
            error_msg = f"合并过程中发生错误:\n{str(e)}"
            self.log(f"\n❌ {error_msg}")
            messagebox.showerror("错误", error_msg)


def main():
    """主函数"""
    root = tk.Tk()
    app = ExcelMergerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
