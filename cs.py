import re
import fitz  # PyMuPDF的导入名称是fitz
import os
from openpyxl import Workbook
from datetime import datetime
from openpyxl.styles import Border, Side, Font, Alignment, numbers
from openpyxl.utils import get_column_letter
def calculate_width(line):
    """
    计算文本行的宽度

    :param line: 文本行
    :return: 文本行的宽度
    """
    width = 0
    for char in line:
        if ord(char) > 127:  # 中文字符
            width += 2
        else:  # 英文字符
            width += 1
    return width

def extract_invoice_data(pdf_path):
    """
    从 PDF 文件中提取发票信息

    :param pdf_path: PDF 文件路径
    :return: 包含发票信息的字典
    """
    doc = fitz.open(pdf_path)
    # 获取第一页内容（通常发票都是一页）
    page = doc.load_page(0)
    text = page.get_text()
    print(text)  # 打印提取的文本，方便检查问题

    # 按行分割文本
    lines = text.split("\n")
    invoice_number = None

    # 提取开票日期
    issue_date = (
        re.search(r"(\d{4}年\d{1,2}月\d{1,2}日)", text).group(1)
        if re.search(r"(\d{4}年\d{1,2}月\d{1,2}日)", text)
        else None
    )
    # 提取发票号码
    # 从“开票人”关键字开始向下查找，或者从开票日期向上查找
    for i, line in enumerate(lines):
        if "开票人：" in line:
            if i + 1 < len(lines):
                invoice_number = lines[i + 1].strip()
        elif issue_date in line:
            if i > 0:
                invoice_number = lines[i - 1].strip()

    # 提取项目名称
    #  初始化一个空列表，用于存储项目名称的行
    project_name = ""
    #  初始化一个布尔变量，用于标记是否开始提取项目名称
    start_extracting = False 
    #  初始化一个布尔变量，用于标记上一行是否超过22个字符
    previous_line_over_22 = False 
    for line in lines:
        line_width = calculate_width(line)
        if line.startswith("*"):
            start_extracting = True
            print(line_width)
            project_name += line.strip()
            if line_width >= 22:
                previous_line_over_22 = True
            else:
                previous_line_over_22 = False
        # 是否继续提取
        elif start_extracting:
            if previous_line_over_22 and line_width <= 22:
                project_name += line.strip()
                break
            elif line_width > 22:
                project_name += line.strip()
                previous_line_over_22 = True
            elif not line.startswith("*"):
                break
    # project_name = " ".join(project_name_lines).strip() if project_name_lines else None
    print(project_name)

    # 提取价税合计（小写），根据文本结构调整正则表达式
    total_amount_matches = re.findall(r"¥\s*(\d+\.\d{2})", text)
    total_amount = total_amount_matches[2] if len(total_amount_matches) > 2 else None

    data = {
        "发票号码": invoice_number,
        "开票日期": issue_date,
        "报销项目": project_name,
        "价税合计": total_amount,
    }
    doc.close()
    return data
def main():
    pdf_path = "6xM5x40轴肩螺钉25.59.pdf"  # 替换为实际文件路径
    data = extract_invoice_data(pdf_path)
    print(data)

if __name__ == "__main__":
    main()
    input("按回车键退出...")
