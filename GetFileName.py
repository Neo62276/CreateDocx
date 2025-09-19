#!/usr/bin/env python3
import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime

class FileScannerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("文件扫描器")
        self.root.geometry("800x600")  # 增大默认窗口尺寸
        self.root.minsize(700, 550)  # 设置更大的最小窗口大小
        
        # 设置中文字体支持
        self.style = ttk.Style()
        
        # 创建顶部框架 - 放置在主窗口顶部
        self.top_frame = ttk.Frame(root, padding="10")
        self.top_frame.pack(fill=tk.X, side=tk.TOP)
        
        # 选择文件夹部分
        self.folder_frame = ttk.Frame(self.top_frame)
        self.folder_frame.pack(fill=tk.X, pady=5)
        
        self.folder_label = ttk.Label(self.folder_frame, text="选择文件夹:")
        self.folder_label.pack(side=tk.LEFT, padx=5)
        
        self.folder_path_var = tk.StringVar()
        self.folder_entry = ttk.Entry(self.folder_frame, textvariable=self.folder_path_var, width=60)
        self.folder_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        self.browse_folder_btn = ttk.Button(self.folder_frame, text="浏览...", command=self.browse_folder)
        self.browse_folder_btn.pack(side=tk.LEFT, padx=5)
        
        # 显示选项部分
        self.options_frame = ttk.Frame(self.top_frame)
        self.options_frame.pack(fill=tk.X, pady=5)
        
        self.show_path_var = tk.BooleanVar(value=True)
        self.show_path_checkbox = ttk.Checkbutton(
            self.options_frame, 
            text="显示完整路径", 
            variable=self.show_path_var, 
            command=self.update_display
        )
        self.show_path_checkbox.pack(side=tk.LEFT, padx=10)
        
        self.show_extension_var = tk.BooleanVar(value=True)
        self.show_extension_checkbox = ttk.Checkbutton(
            self.options_frame, 
            text="显示文件扩展名", 
            variable=self.show_extension_var, 
            command=self.update_display
        )
        self.show_extension_checkbox.pack(side=tk.LEFT, padx=10)
        
        # 扫描按钮
        self.scan_btn = ttk.Button(self.top_frame, text="扫描文件", command=self.scan_files)
        self.scan_btn.pack(fill=tk.X, pady=10)
        
        # 创建主框架 - 放置在顶部框架和底部框架之间
        self.main_frame = ttk.Frame(root, padding=(10, 0, 10, 10))
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 文件列表部分
        self.list_frame = ttk.LabelFrame(self.main_frame, text="扫描结果")
        self.list_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建滚动条
        self.scrollbar = ttk.Scrollbar(self.list_frame)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建文本框显示文件列表
        self.file_list_text = tk.Text(self.list_frame, wrap=tk.NONE, yscrollcommand=self.scrollbar.set)
        self.file_list_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.scrollbar.config(command=self.file_list_text.yview)
        
        # 创建底部框架 - 放置在主窗口底部
        self.bottom_frame = ttk.Frame(root)
        self.bottom_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        # 保存部分
        self.save_frame = ttk.Frame(self.bottom_frame, padding="10")
        self.save_frame.pack(fill=tk.X)
        
        self.format_label = ttk.Label(self.save_frame, text="保存格式:")
        self.format_label.pack(side=tk.LEFT, padx=5)
        
        self.format_var = tk.StringVar(value="txt")
        self.format_combo = ttk.Combobox(self.save_frame, textvariable=self.format_var, values=["txt", "csv", "json"], state="readonly", width=5)
        self.format_combo.pack(side=tk.LEFT, padx=5)
        
        # 添加空白空间，让保存按钮靠右显示
        self.spacer = ttk.Frame(self.save_frame)
        self.spacer.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.save_btn = ttk.Button(self.save_frame, text="保存结果", command=self.save_results)
        self.save_btn.pack(side=tk.RIGHT, padx=5)
        
        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        self.status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 存储扫描的文件路径
        self.scanned_files = []
    
    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path_var.set(folder_selected)
            
    def update_display(self):
        """根据复选框状态更新显示内容"""
        if not self.scanned_files:
            return
        
        self.file_list_text.delete(1.0, tk.END)
        
        for file_path in self.scanned_files:
            display_text = self.format_file_path(file_path)
            self.file_list_text.insert(tk.END, f"{display_text}\n")
            self.root.update_idletasks()
    
    def format_file_path(self, file_path):
        """根据复选框状态格式化文件路径"""
        show_path = self.show_path_var.get()
        show_extension = self.show_extension_var.get()
        
        if show_path:
            if show_extension:
                # 显示完整路径和扩展名
                return file_path
            else:
                # 显示完整路径但不显示扩展名
                base_name = os.path.basename(file_path)
                name_without_ext = os.path.splitext(base_name)[0]
                dir_name = os.path.dirname(file_path)
                return os.path.join(dir_name, name_without_ext)
        else:
            base_name = os.path.basename(file_path)
            if show_extension:
                # 只显示文件名和扩展名
                return base_name
            else:
                # 只显示文件名不显示扩展名
                return os.path.splitext(base_name)[0]
    
    def natural_sort_key(self, s):
        """用于自然数排序的键生成函数"""
        # 提取文件名（不包括路径）
        filename = os.path.basename(s)
        # 将字符串分割成数字和非数字部分
        return [int(text) if text.isdigit() else text.lower() for text in re.split('(\\d+)', filename)]
    
    def scan_files(self):
        folder_path = self.folder_path_var.get()
        
        # 检查文件夹是否存在
        if not os.path.isdir(folder_path):
            messagebox.showerror("错误", f"'{folder_path}' 不是一个有效的目录")
            return
        
        self.status_var.set(f"正在扫描 '{folder_path}'...")
        self.root.update_idletasks()
        
        # 清空之前的结果
        self.file_list_text.delete(1.0, tk.END)
        self.scanned_files = []
        
        try:
            # 扫描文件夹中的所有文件
            for root_dir, dirs, files in os.walk(folder_path):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    self.scanned_files.append(file_path)
            
            # 按文件名进行自然数序排序
            self.scanned_files.sort(key=self.natural_sort_key)
            
            # 显示排序后的结果
            for file_path in self.scanned_files:
                display_text = self.format_file_path(file_path)
                self.file_list_text.insert(tk.END, f"{display_text}\n")
                self.root.update_idletasks()  # 实时更新UI
            
            self.status_var.set(f"扫描完成，共找到 {len(self.scanned_files)} 个文件（已按文件名自然数序排序）")
            messagebox.showinfo("完成", f"扫描完成，共找到 {len(self.scanned_files)} 个文件\n已按文件名自然数序排序")
        except Exception as e:
            self.status_var.set("扫描过程中发生错误")
            messagebox.showerror("错误", f"扫描过程中发生错误: {str(e)}")
    
    def save_results(self):
        if not self.scanned_files:
            messagebox.showwarning("警告", "没有可保存的扫描结果，请先扫描文件")
            return
        
        # 获取保存格式
        file_format = self.format_var.get()
        
        # 获取默认文件名（使用当前时间）
        default_filename = f"文件列表_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{file_format}"
        
        # 打开文件对话框让用户选择保存路径
        file_path = filedialog.asksaveasfilename(
            defaultextension=f".{file_format}",
            filetypes=[
                (f"{file_format.upper()} 文件", f"*.{file_format}"),
                ("所有文件", "*.*")
            ],
            initialfile=default_filename
        )
        
        if not file_path:
            return  # 用户取消保存
        
        try:
            show_path = self.show_path_var.get()
            show_extension = self.show_extension_var.get()
            
            if file_format == "txt":
                # 保存为文本文件
                with open(file_path, 'w', encoding='utf-8') as f:
                    for file in self.scanned_files:
                        if show_path:
                            if show_extension:
                                f.write(f"{file}\n")
                            else:
                                base_name = os.path.basename(file)
                                name_without_ext = os.path.splitext(base_name)[0]
                                dir_name = os.path.dirname(file)
                                f.write(f"{os.path.join(dir_name, name_without_ext)}\n")
                        else:
                            base_name = os.path.basename(file)
                            if show_extension:
                                f.write(f"{base_name}\n")
                            else:
                                f.write(f"{os.path.splitext(base_name)[0]}\n")
            elif file_format == "csv":
                # 保存为CSV文件
                with open(file_path, 'w', encoding='utf-8') as f:
                    # 写入CSV标题
                    f.write("文件名,完整路径\n")
                    for file in self.scanned_files:
                        file_name = os.path.basename(file)
                        if not show_extension:
                            file_name = os.path.splitext(file_name)[0]
                        # 确保CSV格式正确（处理包含逗号的文件名）
                        f.write(f'"{file_name}","{file}"\n')
            elif file_format == "json":
                # 保存为JSON文件
                import json
                # 准备JSON数据
                # 根据显示选项准备显示文件列表
                display_files = []
                for file in self.scanned_files:
                    display_files.append(self.format_file_path(file))
                
                json_data = {
                    "scan_time": datetime.now().isoformat(),
                    "scan_path": self.folder_path_var.get(),
                    "file_count": len(self.scanned_files),
                    "display_options": {
                        "show_full_path": show_path,
                        "show_extension": show_extension
                    },
                    "original_files": self.scanned_files,
                    "display_files": display_files
                }
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(json_data, f, ensure_ascii=False, indent=2)
            
            self.status_var.set(f"结果已保存到 '{file_path}'")
            messagebox.showinfo("成功", f"结果已成功保存到\n{file_path}")
        except Exception as e:
            self.status_var.set("保存过程中发生错误")
            messagebox.showerror("错误", f"保存过程中发生错误: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    # 设置中文字体
    root.option_add("*Font", "SimHei 10")
    app = FileScannerApp(root)
    root.mainloop()