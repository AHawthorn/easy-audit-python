class Config:
    TEMPLATE_PATHS = {
        ("高新", "年报"): "C:\\Users\\lenovo\\Desktop\\audit\\企业会计准则年审报告模板高新.docx",
        ("普通", "年报"): "C:\\Users\\lenovo\\Desktop\\audit\\企业会计准则年审报告模板普通.docx",
        # 其他模板路径可以在这里添加
    }
    EXPORT_FORMATS = ["PDF", "Word"]
    EXCEL_FILE_PATH = 'C:\\Users\\lenovo\\Desktop\\audit\\附注.xlsm'  # Excel 数据文件路径
    OUT_WORD_FILE_PATH = 'C:\\Users\\lenovo\\Desktop\\audit\\更新后的目标文档.docx'  # Word 数据文件路径
