CreateDocx.py 
允许用户选择 Excel 文件、Word 模板文件和输出文件夹，然后根据 Excel 文件中的数据替换 Word 模板中的占位符，生成多个 Word 文档。
具体功能包括：

输入：
Excel 文件（.xlsx）：包含用于替换占位符的数据。
Word 模板文件（.docx）：包含占位符的 Word 文档。
输出文件夹：用于保存生成的 Word 文档。

输出：
多个 Word 文档（.docx）：根据 Excel 文件中的数据替换占位符后生成的文档, 以内容标题命名。

导入模块：
导入所需的 Python 模块，包括 Pandas、python-docx、os、logging、re 和 tkinter。



***
CreatePDF.py
允许用户能够选择 Excel 文件和保存 PDF 文件的目录。

数据读取：运用 pandas 库读取 Excel 文件中的数据。
PDF 生成：利用 reportlab 库依据 Excel 数据生成 PDF 文档。

输入：
一个 Excel 文件（.xlsx），包含 requirement_purpose、background、test_area、mode、node_number、fmea、title 和 steps 等列。
一个保存 PDF 文件的目录。

输出：
多个 PDF 文件，文件名基于 Excel 中的 title 列生成，保存于用户指定的目录。

导入模块：导入所需的 Python 库，如 os、re、pandas、logging、reportlab 和 tkinter。
