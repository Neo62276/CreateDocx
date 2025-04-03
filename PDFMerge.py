import os
import re  # 新增
import tkinter as tk
from tkinter import filedialog
from PyPDF2 import PdfMerger


# 自然排序函数
def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]


def merge_pdf_files(input_path, output_file):
    pdf_files = []
    # 获取指定路径下的所有 .pdf 文件
    for f in os.listdir(input_path):
        if f.endswith('.pdf'):
            file_path = os.path.join(input_path, f)
            pdf_files.append((f, file_path))  # 直接以文件名作为标题

    # 使用自然排序
    pdf_files.sort(key=lambda x: natural_sort_key(x[0]))

    if not pdf_files:
        print("指定路径下没有找到 .pdf 文件。")
        return

    try:
        with PdfMerger() as merger:
            for title, file_path in pdf_files:
                # 修改此处，使用 outline_item 替代 bookmark
                merger.append(file_path, outline_item=title)

            # 保存合并后的 PDF 文件
            merger.write(output_file)
        print(f"成功合并所有 PDF 文件到 {output_file}。")
    except Exception as e:
        print(f"合并 PDF 文件时出现错误: {e}")


def select_input_path():
    path = filedialog.askdirectory()
    input_path_entry.delete(0, tk.END)
    input_path_entry.insert(0, path)


def select_output_path():
    file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF 文件", "*.pdf")])
    output_path_entry.delete(0, tk.END)
    output_path_entry.insert(0, file_path)


def start_merge():
    input_path = input_path_entry.get()
    output_path = output_path_entry.get()
    if not input_path:
        print("请选择要合并的 PDF 文件所在路径。")
        return
    if not output_path:
        print("请选择合并后 PDF 文件的保存路径。")
        return
    merge_pdf_files(input_path, output_path)


# 创建主窗口
root = tk.Tk()
root.title("合并 PDF 文件")

# 输入路径选择
input_path_label = tk.Label(root, text="选择要合并的 PDF 文件所在路径:")
input_path_label.pack()
input_path_entry = tk.Entry(root, width=50)
input_path_entry.pack()
input_path_button = tk.Button(root, text="选择路径", command=select_input_path)
input_path_button.pack()

# 输出路径选择
output_path_label = tk.Label(root, text="选择合并后 PDF 文件的保存路径:")
output_path_label.pack()
output_path_entry = tk.Entry(root, width=50)
output_path_entry.pack()
output_path_button = tk.Button(root, text="选择路径", command=select_output_path)
output_path_button.pack()

# 开始合并按钮
merge_button = tk.Button(root, text="开始合并", command=start_merge)
merge_button.pack()

# 运行主循环
root.mainloop()
