#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel表格合并工具 - 图形界面版本
功能：
1. 合并多个具有相同表头的Excel文件
2. 自动重新生成第一列的递增序号（从1开始）
3. 支持拖拽文件到窗口
4. 自动输出到桌面，文件名格式：账单汇总_YYYYMMDD HH:MM:SS.xlsx
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk
import pandas as pd
from pathlib import Path
import threading
from datetime import datetime

# 尝试导入拖拽支持库
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False


def get_desktop_path():
    """获取桌面路径"""
    # 尝试多种方式获取桌面路径
    # Windows
    if os.name == 'nt':
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        if not os.path.exists(desktop):
            # 尝试中文路径
            desktop = os.path.join(os.path.expanduser("~"), "桌面")
        if not os.path.exists(desktop):
            # 使用用户目录
            desktop = os.path.expanduser("~")
    else:
        # Linux/Mac
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        if not os.path.exists(desktop):
            desktop = os.path.join(os.path.expanduser("~"), "桌面")
        if not os.path.exists(desktop):
            desktop = os.path.expanduser("~")
    
    return desktop


def generate_output_filename():
    """生成输出文件名：账单汇总_YYYYMMDD HH-MM-SS.xlsx"""
    now = datetime.now()
    # Windows不允许文件名包含冒号，使用横杠代替
    filename = now.strftime("账单汇总_%Y%m%d %H-%M-%S.xlsx")
    return filename


class ExcelMergerGUI:
    """Excel合并工具图形界面"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("📊 Excel表格合并工具")
        self.root.geometry("1000x700")
        self.root.resizable(True, True)
        
        # 设置最小窗口大小
        self.root.minsize(800, 550)
        
        # 设置样式
        style = ttk.Style()
        style.theme_use('clam')
        
        # 存储选择的文件
        self.selected_files = []
        
        # 创建界面
        self.create_widgets()
        
        # 设置拖拽支持
        self.setup_drag_and_drop()
        
    def create_widgets(self):
        """创建GUI组件"""
        
        # 标题框架
        title_frame = tk.Frame(self.root, bg="#2c3e50", height=80)
        title_frame.pack(fill=tk.X, padx=0, pady=0)
        title_frame.pack_propagate(False)
        
        title_label = tk.Label(
            title_frame,
            text="📊 Excel表格合并工具",
            font=("Arial", 24, "bold"),
            fg="white",
            bg="#2c3e50"
        )
        title_label.pack(pady=20)
        
        subtitle_label = tk.Label(
            title_frame,
            text="合并相同表头的Excel文件，自动重新编号 | 输出到桌面",
            font=("Arial", 11),
            fg="#ecf0f1",
            bg="#2c3e50"
        )
        subtitle_label.pack(pady=(0, 10))
        
        # 主容器
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 文件选择区域
        file_frame = tk.LabelFrame(
            main_frame,
            text="📁 选择要合并的Excel文件（支持拖拽文件到此处）",
            font=("Arial", 12, "bold"),
            padx=15,
            pady=15
        )
        file_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
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
            font=("Arial", 11, "bold"),
            padx=25,
            pady=10,
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
            font=("Arial", 11, "bold"),
            padx=25,
            pady=10,
            cursor="hand2",
            relief=tk.FLAT
        )
        clear_btn.pack(side=tk.LEFT)
        
        # 文件数量标签
        self.file_count_label = tk.Label(
            button_frame,
            text="已选择: 0 个文件",
            font=("Arial", 11, "bold"),
            fg="#2c3e50"
        )
        self.file_count_label.pack(side=tk.RIGHT)
        
        # 拖拽提示区域
        self.drop_hint_frame = tk.Frame(file_frame, bg="#ecf0f1", height=60)
        self.drop_hint_frame.pack(fill=tk.X, pady=(0, 10))
        self.drop_hint_frame.pack_propagate(False)
        
        drop_hint_label = tk.Label(
            self.drop_hint_frame,
            text="🎯 拖拽Excel文件到此窗口即可添加",
            font=("Arial", 12),
            fg="#7f8c8d",
            bg="#ecf0f1"
        )
        drop_hint_label.pack(expand=True)
        
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
            font=("Consolas", 10),
            selectmode=tk.EXTENDED,
            bg="#ecf0f1",
            relief=tk.FLAT,
            height=10
        )
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)
        
        # 日志区域
        log_frame = tk.LabelFrame(
            main_frame,
            text="📋 操作日志",
            font=("Arial", 12, "bold"),
            padx=15,
            pady=15
        )
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=8,
            font=("Consolas", 10),
            bg="#2c3e50",
            fg="#ecf0f1",
            relief=tk.FLAT,
            state=tk.DISABLED
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 合并按钮
        merge_btn = tk.Button(
            main_frame,
            text="✨ 开始合并（自动保存到桌面）",
            command=self.start_merge,
            bg="#27ae60",
            fg="white",
            font=("Arial", 14, "bold"),
            padx=40,
            pady=15,
            cursor="hand2",
            relief=tk.FLAT
        )
        merge_btn.pack(fill=tk.X)
        
        # 初始日志
        self.log("欢迎使用Excel表格合并工具！")
        self.log("请选择或拖拽Excel文件到窗口...")
        self.log(f"合并后的文件将自动保存到桌面")
        
    def setup_drag_and_drop(self):
        """设置拖拽支持"""
        if HAS_DND:
            try:
                # 为整个窗口注册拖拽
                self.root.drop_target_register(DND_FILES)
                self.root.dnd_bind('<<Drop>>', self.handle_drop)
                self.log("✓ 拖拽功能已启用")
            except Exception as e:
                self.log(f"⚠️ 拖拽功能初始化失败: {str(e)}")
        else:
            self.log("⚠️ 拖拽功能不可用（需要安装 tkinterdnd2）")
            
    def handle_drop(self, event):
        """处理拖拽放置事件"""
        # 解析拖拽的文件路径
        files = self.parse_drop_data(event.data)
        
        added_count = 0
        for file_path in files:
            # 只接受Excel文件
            if file_path.lower().endswith(('.xlsx', '.xls')):
                if file_path not in self.selected_files:
                    self.selected_files.append(file_path)
                    self.file_listbox.insert(tk.END, os.path.basename(file_path))
                    added_count += 1
        
        if added_count > 0:
            self.update_file_count()
            self.log(f"✓ 通过拖拽添加了 {added_count} 个文件")
        else:
            self.log("⚠️ 没有有效的Excel文件被添加")
    
    def parse_drop_data(self, data):
        """解析拖拽数据，提取文件路径"""
        files = []
        # 处理不同操作系统的路径格式
        # Windows: {path1} {path2} 或 path1\npath2
        # Linux: file://path1\nfile://path2
        
        if '{' in data:
            # Windows格式，花括号包围的路径
            import re
            matches = re.findall(r'\{([^}]+)\}', data)
            if matches:
                files.extend(matches)
            else:
                # 没有花括号，按空格分割
                files.extend(data.split())
        else:
            # 按换行或空格分割
            items = data.replace('\r', '').split('\n')
            for item in items:
                item = item.strip()
                if item:
                    # 移除 file:// 前缀
                    if item.startswith('file://'):
                        item = item[7:]
                    files.append(item)
        
        # 清理路径
        cleaned_files = []
        for f in files:
            f = f.strip()
            if f and os.path.isfile(f):
                cleaned_files.append(f)
        
        return cleaned_files
        
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
        
    def start_merge(self):
        """开始合并（在新线程中执行）"""
        # 验证输入
        if not self.selected_files:
            messagebox.showwarning("提示", "请先选择要合并的Excel文件！")
            return
        
        # 自动生成输出文件路径（桌面 + 时间戳文件名）
        desktop_path = get_desktop_path()
        output_filename = generate_output_filename()
        output_file = os.path.join(desktop_path, output_filename)
        
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
    if HAS_DND:
        # 使用支持拖拽的TkinterDnD
        root = TkinterDnD.Tk()
    else:
        # 使用普通的Tk
        root = tk.Tk()
    
    app = ExcelMergerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
