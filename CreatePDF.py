import os
import re
import pandas as pd
import logging
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
# 新增导入 tkinter 库
import tkinter as tk
from tkinter import filedialog, messagebox

# 定义一个常量，代表每个数字所占的宽度
DIGIT_WIDTH = 4
# 定义序号和文本之间的固定间隔
SPACE_WIDTH = 6

def create_title_table(doc, title):
    """
    创建包含标题的表格
    :param doc: 文档对象
    :param title: 标题文本
    :return: 标题表格对象
    """
    styles = getSampleStyleSheet()
    # 调整小标题字体大小
    title_style = styles['Heading1']
    title_style.fontSize = 12
    title = Paragraph(title, title_style)
    # 创建一个包含标题的表格，用于添加边框
    title_table = Table([[title]], colWidths=[doc.width])
    title_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),  # 背景颜色设为白色
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 12),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LINEABOVE', (0, 0), (-1, 0), 1, colors.lightblue),  # 上侧淡蓝色线条
        ('LINELEFT', (0, 0), (0, -1), 1, colors.lightblue),  # 左侧淡蓝色线条，覆盖整个左侧
    ]))
    return title_table

def create_requirement_table(doc, requirement_purpose, background):
    """
    创建需求、背景部分的表格
    :param doc: 文档对象
    :param requirement_purpose: 需求目的
    :param background: 背景信息
    :return: 需求表格对象
    """
    styles = getSampleStyleSheet()
    # 自定义段落样式，设置字体大小
    custom_style = styles['Normal']
    custom_style.fontSize = 8
    # 需求、背景部分生成表格
    requirement_data = [
        ["Requirement/Purpose", requirement_purpose],
        ["Background", Paragraph(background, custom_style)]
    ]
    # 进一步调整列宽比例
    requirement_col_widths = [
        doc.width * 0.2,  # 第一列宽度
        doc.width * 0.8   # 第二列宽度
    ]
    requirement_table = Table(requirement_data, colWidths=requirement_col_widths)
    requirement_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), colors.lightblue),  # 设置 Requirement/Purpose 背景颜色为浅蓝色
        ('BACKGROUND', (0, 1), (0, 1), colors.lightblue),  # 设置 Background 背景颜色为浅蓝色
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),  # 适当减小字体大小
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('WORDWRAP', (0, 0), (-1, -1), 'CJK'),  # 启用自动换行
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # 垂直顶部对齐，确保内容完整显示
        ('LEFTPADDING', (0, 0), (-1, -1), 6),  # 添加左内边距
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),  # 添加右内边距
    ]))
    return requirement_table

def create_test_table(doc, test_area, mode, node_number, fmea):
    """
    创建测试区域、模式和节点编号部分的表格
    :param doc: 文档对象
    :param test_area: 测试区域
    :param mode: 模式
    :param node_number: 节点编号
    :param fmea: FMEA编号
    :return: 测试表格对象
    """
    styles = getSampleStyleSheet()
    # 自定义段落样式，设置字体大小
    custom_style = styles['Normal']
    custom_style.fontSize = 8
    # 测试区域、模式和节点编号部分生成表格
    test_data = [
        ["Test Area", test_area, "Mode", mode],
        ["Node Number", node_number, "FMEA#", fmea]
    ]
    # 调整列宽比例
    test_col_widths = [
        doc.width * 0.2,  # 第一列宽度
        doc.width * 0.3,  # 第二列宽度
        doc.width * 0.2,  # 第三列宽度
        doc.width * 0.3   # 第四列宽度
    ]
    test_table = Table(test_data, colWidths=test_col_widths)
    test_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), colors.lightblue),  # 设置 Test Area 背景颜色为浅蓝色
        ('BACKGROUND', (2, 0), (2, 0), colors.lightblue),  # 设置 Mode 背景颜色为浅蓝色
        ('BACKGROUND', (0, 1), (0, 1), colors.lightblue),  # 设置 Node Number 背景颜色为浅蓝色
        ('BACKGROUND', (2, 1), (2, 1), colors.lightblue),  # 设置 FMEA# 背景颜色为浅蓝色
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),  # 适当减小字体大小
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('WORDWRAP', (0, 0), (-1, -1), 'CJK'),  # 启用自动换行
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # 垂直顶部对齐，确保内容完整显示
        ('LEFTPADDING', (0, 0), (-1, -1), 6),  # 添加左内边距
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),  # 添加右内边距
    ]))
    return test_table

def create_procedure_table(doc, steps):
    """
    创建过程表格
    :param doc: 文档对象
    :param steps: 步骤列表
    :return: 过程表格对象
    """
    styles = getSampleStyleSheet()
    # 自定义段落样式，设置字体大小
    custom_style = styles['Normal']
    custom_style.fontSize = 8
    data = [["Procedure", "Pass"]]
    for i, step in enumerate(steps, start=1):
        # 把步骤内容中的换行符替换为 <br/>
        step = step.replace('\n', '<br/>')
        # 创建一个带缩进的段落
        bullet_style = custom_style.clone('BulletStyle')
        # 计算序号的位数
        digit_count = len(str(i))
        # 动态调整 firstLineIndent 的值
        bullet_style.firstLineIndent = - (SPACE_WIDTH + digit_count * DIGIT_WIDTH)
        bullet_style.leftIndent = 15  # 正值用于缩进内容，保证换行后文本首字对齐
        bullet_style.spaceBefore = 6  # 可以根据需要调整段落前间距
        bullet_style.spaceAfter = 6  # 可以根据需要调整段落后续间距
        step_text = f"{i}. {step}"
        step_paragraph = Paragraph(step_text, bullet_style)
        data.append([step_paragraph, ""])
    # 调整列宽比例为 9:1
    col_widths = [doc.width * 0.9, doc.width * 0.1]
    table = Table(data, colWidths=col_widths, repeatRows=1)  # 设置 repeatRows=1 使表头跨页显示
    # 设置表格样式，允许内容自动换行
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 1), (-1, -1), 'LEFT'),  # 修改表格内容为左对齐
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),  # 适当减小字体大小
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),  # 除第一行外，其余行背景设为白色
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('WORDWRAP', (0, 1), (-1, -1), 'CJK'),  # 启用自动换行
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # 垂直顶部对齐，确保内容完整显示
        ('LEFTPADDING', (0, 0), (-1, -1), 6),  # 添加左内边距
        ('RIGHTPADDING', (0, 0), (-1, -1), 6)  # 添加右内边距
    ]))
    return table

def create_comments_table(doc):
    """
    创建注释部分的表格
    :param doc: 文档对象
    :return: 注释表格对象
    """
    styles = getSampleStyleSheet()
    # 自定义段落样式，设置字体大小
    custom_style = styles['Normal']
    custom_style.fontSize = 8
    # 注释部分生成表格
    comments_data = [
        ["Comments/Notes:"],
        [" "],
        [" "],
        [" "]
    ]
    comments_col_widths = [doc.width]
    comments_table = Table(comments_data, colWidths=comments_col_widths)
    comments_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('BOX', (0, 0), (-1, -1), 1, colors.black),  # 只显示外边框
        ('LEFTPADDING', (0, 0), (-1, -1), 6),  # 添加左内边距
        ('RIGHTPADDING', (0, 0), (-1, -1), 6)  # 添加右内边距
    ]))
    return comments_table

def create_ewa_table(doc):
    """
    创建执行、见证、批准部分的表格
    :param doc: 文档对象
    :return: 执行、见证、批准表格对象
    """
    styles = getSampleStyleSheet()
    # 自定义段落样式，设置字体大小
    custom_style = styles['Normal']
    custom_style.fontSize = 8
    # 执行、见证、批准部分生成表格
    executed_witnessed_approved_data = [
        ["Executed by :", "", "Date :", ""],
        ["Witnessed by :", "", "Date :", ""],
        ["Approved by:", "", "Date :", ""]
    ]
    total_width = doc.width
    ewa_col_widths = [
        total_width * 2 / 10,  # 第一列宽度
        total_width * 4.5 / 10,  # 第二列宽度
        total_width * 1 / 10,  # 第三列宽度
        total_width * 2.5 / 10   # 第四列宽度
    ]
    ewa_table = Table(executed_witnessed_approved_data, colWidths=ewa_col_widths)
    ewa_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.burlywood),  # 第一列背景设为棕黄色
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 15),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),  # 显示表格边框
        ('LEFTPADDING', (0, 0), (-1, -1), 6),  # 添加左内边距
        ('RIGHTPADDING', (0, 0), (-1, -1), 6)  # 添加右内边距
    ]))
    return ewa_table

# 修改函数，添加 steps 和 doc_name 参数
def create_pdf(requirement_purpose, background, test_area, mode, node_number, fmea, title, steps, doc_name):
    """
    创建 PDF 文档
    :param requirement_purpose: 需求目的
    :param background: 背景信息
    :param test_area: 测试区域
    :param mode: 模式
    :param node_number: 节点编号
    :param fmea: FMEA编号
    :param title: 标题
    :param steps: 步骤列表
    :param doc_name: 文档名称
    """
    # 设置页面边距，使用传入的文件名
    doc = SimpleDocTemplate(doc_name, pagesize=A4, leftMargin=0.5*inch, rightMargin=0.5*inch, topMargin=0.5*inch, bottomMargin=0.5*inch)
    elements = []
    styles = getSampleStyleSheet()

    # 创建标题表格
    title_table = create_title_table(doc, title)
    elements.append(title_table)

    # 添加有明确高度的空白段落用于分隔表格
    spacer = Paragraph("<br/>" * 1, styles['Heading1'])  # 可调整 <br/> 的数量改变间距
    elements.append(spacer)

    # 创建需求表格
    requirement_table = create_requirement_table(doc, requirement_purpose, background)
    elements.append(requirement_table)

    # 创建测试表格
    test_table = create_test_table(doc, test_area, mode, node_number, fmea)
    elements.append(test_table)

    # 添加有明确高度的空白段落用于分隔表格
    spacer = Paragraph("<br/>" * 2, styles['Normal'])  # 可调整 <br/> 的数量改变间距
    elements.append(spacer)

    # 创建过程表格
    procedure_table = create_procedure_table(doc, steps)
    elements.append(procedure_table)

    # 添加有明确高度的空白段落用于分隔表格
    spacer = Paragraph("<br/>" * 2, styles['Normal'])  # 可调整 <br/> 的数量改变间距
    elements.append(spacer)

    # 创建注释表格
    comments_table = create_comments_table(doc)

    # 创建执行、见证、批准表格
    ewa_table = create_ewa_table(doc)

    # 添加有明确高度的空白段落用于分隔表格
    spacer = Paragraph("<br/>" * 2, styles['Normal'])  # 可调整 <br/> 的数量改变间距

    # 使用 KeepTogether 确保注释表格和 EWA 表格不跨页
    from reportlab.platypus import KeepTogether
    combined_elements = KeepTogether([comments_table, spacer, ewa_table])
    elements.append(combined_elements)

    # 添加分页符（如果需要）
    elements.append(PageBreak())

    doc.build(elements)

def sanitize_filename(filename):
    """
    移除或替换文件名中不允许的字符
    :param filename: 原始文件名
    :return: 处理后的文件名
    """
    return re.sub(r'[\\/*?:"<>|]', '_', filename)

def select_excel_file():
    file_path = filedialog.askopenfilename(title="选择 Excel 文件", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        excel_path_entry.delete(0, tk.END)
        excel_path_entry.insert(0, file_path)

def select_output_directory():
    dir_path = filedialog.askdirectory(title="选择保存 PDF 文件的目录")
    if dir_path:
        output_dir_entry.delete(0, tk.END)
        output_dir_entry.insert(0, dir_path)

def convert_to_string(value):
    """
    将值转换为字符串，如果值为 NaN 则转换为空字符串
    :param value: 输入值
    :return: 字符串类型的值
    """
    return str(value) if pd.notna(value) else ""

def generate_pdfs():
    excel_file_path = excel_path_entry.get()
    save_directory = output_dir_entry.get()

    if not excel_file_path or not save_directory:
        messagebox.showerror("错误", "请选择 Excel 文件和保存目录")
        return

    try:
        df = pd.read_excel(excel_file_path)
        for index, row in df.iterrows():
            # 确保所有需要的参数都是字符串类型
            requirement_purpose = convert_to_string(row['requirement_purpose'])
            background = convert_to_string(row['background'])
            test_area = convert_to_string(row['test_area'])
            mode = convert_to_string(row['mode'])
            node_number = convert_to_string(row['node_number'])
            fmea = convert_to_string(row['fmea'])
            title = convert_to_string(row['title'])
            steps_str = convert_to_string(row['steps'])
            steps = steps_str.split(';')
            sanitized_title = sanitize_filename(title)
            doc_name = os.path.join(save_directory, f"{sanitized_title}.pdf")
            create_pdf(
                requirement_purpose,
                background,
                test_area,
                mode,
                node_number,
                fmea,
                title,
                steps,
                doc_name
            )
            logging.info(f"Generated PDF: {doc_name}")
        messagebox.showinfo("完成", "PDF 文件生成完成")
    except Exception as e:
        messagebox.showerror("错误", f"生成 PDF 时出现错误: {str(e)}")

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 创建主窗口
root = tk.Tk()
root.title("PDF 生成器")

# 创建选择 Excel 文件的组件
excel_path_label = tk.Label(root, text="选择 Excel 文件:")
excel_path_label.pack(pady=10)

excel_path_entry = tk.Entry(root, width=50)
excel_path_entry.pack(pady=5)

excel_select_button = tk.Button(root, text="选择文件", command=select_excel_file)
excel_select_button.pack(pady=5)

# 创建选择输出目录的组件
output_dir_label = tk.Label(root, text="选择保存目录:")
output_dir_label.pack(pady=10)

output_dir_entry = tk.Entry(root, width=50)
output_dir_entry.pack(pady=5)

output_dir_select_button = tk.Button(root, text="选择目录", command=select_output_directory)
output_dir_select_button.pack(pady=5)

# 创建生成 PDF 的按钮
generate_button = tk.Button(root, text="生成 PDF", command=generate_pdfs)
generate_button.pack(pady=20)

# 运行主循环
root.mainloop()