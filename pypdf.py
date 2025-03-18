import re
import fitz  # PyMuPDF的导入名称是fitz
import os
import sys
from openpyxl import Workbook
from datetime import datetime
from openpyxl.styles import Border, Side, Font, Alignment, numbers
from openpyxl.utils import get_column_letter


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
    # print(text)  # 打印提取的文本，方便检查问题

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

    # 提取项目名称
    #  初始化一个空列表，用于存储项目名称的行
    project_name = ""
    #  初始化一个布尔变量，用于标记是否开始提取项目名称
    start_extracting = False
    #  初始化一个布尔变量，用于标记上一行是否超过22个字符 也可以考虑21个字符
    previous_line_over_22 = False
    for line in lines:
        line_width = calculate_width(line)
        if line.startswith("*"):
            start_extracting = True
            print(line_width,line)
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

    # 提取价税合计（小写），根据文本结构调整正则表达式
    total_amount_matches = re.findall(r"¥\s*(\d+\.\d{2})", text)
    total_amount = (
        float(total_amount_matches[2]) if len(total_amount_matches) > 2 else None
    )

    data = {
        "发票号码": invoice_number,
        "开票日期": issue_date,
        "报销项目": project_name,
        "价税合计": total_amount,
    }
    doc.close()
    return data


def traverse_pdf_files(document_dir):
    """
    遍历指定目录下的所有 PDF 文件

    :param document_dir: 文件目录
    :return: 包含所有 PDF 文件名列表
    """
    pdf_data = []
    for filename in os.listdir(document_dir):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(document_dir, filename)
            invoice_data = extract_invoice_data(pdf_path)
            pdf_data.append((filename, invoice_data))
    return pdf_data


def write_to_excel(pdf_data, excel_path):
    """
    将提取的发票信息写入 Excel 文件

    :param pdf_data: 包含发票信息的列表
    :param excel_path: Excel 文件路径
    """
    # openpyxl 使用的是“写入时构建”的方法，
    # 即在你调用 ws.append() 方法时，数据会被直接写入到内存中的一个数据结构中，而不是整个Excel文件。
    # 如果后期数据量过大，可以考虑使用 openpyxl 的优化方法。如 write_only 模式。
    wb = Workbook()
    ws = wb.active
    header = ["文件名", "发票号码", "开票日期", "项目名称", "价税合计（小写）"]
    ws.append(header)

    # 设置表头样式：加粗且居中
    header_font = Font(bold=True)
    header_alignment = Alignment(horizontal="center")
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = header_alignment

    for filename, invoice_data in pdf_data:
        row = [filename] + list(invoice_data.values())
        ws.append(row)

    # 适应列宽
    for column in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        for cell in column:
            try:
                cell_value = str(cell.value)
                # 计算字符宽度，中文字符宽度乘以 2
                length = sum(2 if ord(c) > 127 else 1 for c in cell_value)
                if length > max_length:
                    max_length = length
            except:
                pass
        adjusted_width = max_length + 2
        ws.column_dimensions[column_letter].width = adjusted_width

    # 添加框线
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    for row in ws.iter_rows():
        for cell in row:
            cell.border = thin_border

    wb.save(excel_path)


def resource_path(relative_path):
    """
    获取资源文件的绝对路径..

    :param relative_path: 相对路径
    :return: 绝对路径
    """
    if getattr(sys, "frozen", False):
        # 运行在打包后的exe中
        base_path = os.path.dirname(sys.executable)
        print(f"打包后 base_path: {base_path}")
    else:
        # 运行在普通Python环境中
        base_path = os.path.dirname(os.path.abspath(__file__))
        print(f"普通环境 base_path: {base_path}")
    return os.path.join(base_path, relative_path)


def main():
    try:
        # 获取当前时间
        now = datetime.now()
        # 格式化时间字符串
        formatted_time = now.strftime("%Y年%m月%d日%H点%M分导出")

        # 构建 document 文件夹的路径 源数据文件夹
        document_dir = resource_path("document")
        if not os.path.exists(document_dir):
            raise FileNotFoundError("document 文件夹不存在，请检查路径是否正确")

        # 构建 data 文件夹的路径 结果数据文件夹
        data_dir = resource_path("data")
        # 如果 data 文件夹不存在，则创建
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)
        # 构建 Excel 文件路径 实时时间中不包含秒，短时间内重复运行可能会导致文件名重复
        excel_path = os.path.join(data_dir, f"{formatted_time}导出发票信息.xlsx")

        pdf_data = traverse_pdf_files(document_dir)
        write_to_excel(pdf_data, excel_path)
        print(f"发票信息已保存到 {excel_path}")
    except Exception as e:
        print(f"发生错误: {e}")
    input("按任意键退出...")


if __name__ == "__main__":
    main()
    # pdf_path = "6x8x4衬套66.6.pdf"  # 替换为实际文件路径
    # data = extract_invoice_data(pdf_path)
    # print(data)
