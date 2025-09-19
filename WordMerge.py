import os
import re  # 新增
from docx import Document
from docxcompose.composer import Composer
import tkinter as tk
from tkinter import filedialog

# 自然排序函数，将字符串按数字和字母分割，数字部分按数值大小排序，字母部分按字母顺序排序
def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]

def merge_word_documents(selected_path, output_file):
    try:
        # 获取指定路径下的所有 .docx 文件
        docx_files = [os.path.join(selected_path, f) for f in os.listdir(selected_path) if f.endswith('.docx')]
        # 使用自然排序
        docx_files.sort(key=natural_sort_key)

        if not docx_files:
            print("指定路径下没有找到 .docx 文件。")
            return

        # 打开第一个文档作为基础文档
        first_doc = Document(docx_files[0])
        composer = Composer(first_doc)

        # 依次合并其他文档
        for doc_file in docx_files[1:]:
            doc = Document(doc_file)
#           composer.doc.add_section()  # 添加分节符
            composer.doc.add_page_break()  # 添加分页符
            composer.append(doc)

        # 保存合并后的文档
        composer.save(output_file)
        print(f"成功合并所有 Word 文档到 {output_file}。")
    except Exception as e:
        print(f"合并 Word 文档时出现错误: {e}")

def select_input_path():
    path = filedialog.askdirectory()
    input_path_entry.delete(0, tk.END)
    input_path_entry.insert(0, path)

def select_output_path():
    file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
    output_path_entry.delete(0, tk.END)
    output_path_entry.insert(0, file_path)

def start_merge():
    input_path = input_path_entry.get()
    output_path = output_path_entry.get()
    if not input_path:
        print("请选择要合并的 Word 文档所在路径。")
        return
    if not output_path:
        print("请选择合并后文档的保存路径。")
        return
    merge_word_documents(input_path, output_path)

# 创建主窗口
root = tk.Tk()
root.title("合并 Word 文档")

# 输入路径选择
input_path_label = tk.Label(root, text="选择要合并的 Word 文档所在路径:")
input_path_label.pack()
input_path_entry = tk.Entry(root, width=50)
input_path_entry.pack()
input_path_button = tk.Button(root, text="选择路径", command=select_input_path)
input_path_button.pack()

# 输出路径选择
output_path_label = tk.Label(root, text="选择合并后文档的保存路径:")
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