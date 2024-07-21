import os

class Config:
    BASE_DIR = r"C:\\audit"  # 固定的文件夹位置

    TEMPLATE_PATHS = {
        ("高新", "年报"): os.path.join(BASE_DIR, "模板", "企业会计准则年审报告模板高新.docx"),
        ("普通", "年报"): os.path.join(BASE_DIR, "模板", "企业会计准则年审报告模板普通.docx"),
        # 其他模板路径可以在这里添加
    }
    EXPORT_FORMATS = ["PDF", "Word"]
    EXCEL_FILE_PATH = os.path.join(BASE_DIR, '数据', '附注.xlsm')  # Excel 数据文件路径
    OUT_WORD_FILE_PATH = os.path.join(BASE_DIR, '更新后的目标文档.docx')  # 生成的Word文件路径
    BASIC_EXCEL_FILE_PATH = os.path.join(BASE_DIR, '数据', '3.报表.xlsx')  # 基本信息数据文件路径
