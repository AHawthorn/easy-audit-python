import openpyxl
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from config import Config
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def read_excel(file_path):
    """读取Excel文件并返回工作簿对象"""
    workbook = openpyxl.load_workbook(file_path, data_only=True)
    return workbook


def extract_subjects(workbook):
    """从目录sheet页提取科目名称"""
    sheet = workbook['目录']
    subjects = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[3] == '是':  # 假设"是否做附注表格"在第2列
            subjects.append(row[1])  # 假设"科目名称"在第1列
    return subjects


def format_number(value):
    """格式化数字，添加千位分隔符"""
    if isinstance(value, (int, float)):
        return '{:,.2f}'.format(value)
    return value


def extract_table_data(workbook, sheet_name):
    """提取财务表格数据并进行格式化"""
    sheet = workbook[sheet_name]
    data = []
    for row in sheet.iter_rows(min_row=3, min_col=3, values_only=True):  # 从第3行第3列开始
        formatted_row = [format_number(cell) if cell is not None else "" for cell in row]
        data.append(formatted_row)
    return data


def set_cell_font(cell, font_name, font_size, bold=False):
    """设置单元格字体"""
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = font_name
            run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
            run.font.size = Pt(font_size)
            run.bold = bold


def set_table_borders(table):
    """设置表格的上下两条外边框加粗，左右两条外边框不显示"""
    tbl = table._element
    tblBorders = OxmlElement('w:tblBorders')

    # 设置上下边框加粗
    top = OxmlElement('w:top')
    top.set(qn('w:val'), 'single')
    top.set(qn('w:sz'), '12')
    tblBorders.append(top)

    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '12')
    tblBorders.append(bottom)

    # 设置左右边框不显示
    left = OxmlElement('w:left')
    left.set(qn('w:val'), 'none')
    tblBorders.append(left)

    right = OxmlElement('w:right')
    right.set(qn('w:val'), 'none')
    tblBorders.append(right)

    # 其他边框设置为默认
    insideH = OxmlElement('w:insideH')
    insideH.set(qn('w:val'), 'single')
    tblBorders.append(insideH)

    insideV = OxmlElement('w:insideV')
    insideV.set(qn('w:val'), 'single')
    tblBorders.append(insideV)

    tbl.tblPr.append(tblBorders)

def insert_table_into_doc(doc, table_data, placeholder):
    """将表格插入到Word文档的指定位置"""
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            # 替换占位符段落
            p = paragraph._element

            # 在匹配内容后添加一个换行符
            run = paragraph.add_run()
            run.add_break()

            # 创建表格并设置样式
            tbl = doc.add_table(rows=0, cols=len(table_data[0]))
            tbl.style = 'Table Grid'
            tbl.autofit = False

            # 插入表格数据
            for row_idx, row_data in enumerate(table_data):
                row = tbl.add_row().cells
                for i, cell_data in enumerate(row_data):
                    row[i].text = str(cell_data)

                    # 表头样式设置
                    if row_idx == 0:
                        set_cell_font(row[i], '宋体', 10, bold=True)
                        row[i].paragraphs[0].alignment = 1  # 居中
                    # 表格最后一行如果包含“合计”，也要加粗且居中显示
                    elif cell_data and "合计" in str(cell_data):
                        set_cell_font(row[i], '宋体', 10, bold=True)
                        row[i].paragraphs[0].alignment = 1  # 居中
                    else:
                        if isinstance(cell_data, (int, float)):
                            set_cell_font(row[i], 'Arial Narrow', 10)
                        else:
                            set_cell_font(row[i], '宋体', 10)

            # 设置表格的上下两条外边框加粗，左右两条外边框不显示
            set_table_borders(tbl)

            # 在表格后插入两个换行符以避免粘连
            new_paragraph = doc.add_paragraph()
            new_paragraph.add_run('\n')

            break

def export_report(template, report_type, export_format):
    """导出报告"""
    excel_path = Config.EXCEL_FILE_PATH
    word_path = Config.TEMPLATE_PATHS.get((template, report_type))
    if not word_path:
        raise ValueError("未找到对应的报告模板")

    # 读取Excel文件
    workbook = read_excel(excel_path)

    # 提取科目名称
    subjects = extract_subjects(workbook)

    # 读取Word文档
    doc = Document(word_path)

    # 提取并插入表格数据
    for subject in subjects:
        table_data = extract_table_data(workbook, subject)
        insert_table_into_doc(doc, table_data, f'«{subject}»')

    # 保存修改后的Word文档
    output_path = Config.OUT_WORD_FILE_PATH
    doc.save(output_path)

    if export_format == "PDF":
        # 如果需要将Word转换为PDF，可以使用第三方库如`python-docx2pdf`
        pass

    return output_path
