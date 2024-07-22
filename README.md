
审计项目
# 功能模块
## 年报
已完成
## 费用专项
## 收入专项

# 部署

1. 生成easy-audit.spec文件，执行命令： pyinstaller --onefile --windowed --name easy-audit app/main.py
2. 修改easy-audit.spec文件：  

```
import os

block_cipher = None

a = Analysis(
    ['app/main.py'],
    pathex=['.'],
    binaries=[('C:\\Users\\lenovo\\AppData\\Local\\Programs\\Python\\Python312\\python312.dll', '.')],
    datas=[
        ('app/config.py', 'app'),
        ('app/logic.py', 'app'),
        ('app/ui.py', 'app'),
        ('app/__init__.py', 'app'),
        ('模板/企业会计准则年审报告模板高新.docx', '模板'),
        ('模板/企业会计准则年审报告模板普通.docx', '模板'),
        ('数据/附注.xlsm', '数据'),
        ('数据/3.报表.xlsx', '数据')
    ],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='easy-audit',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 设置为 False 以去除命令行窗口
    icon='favicon.ico'  # 指定图标文件路径
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='easy-audit'
)
```

3. 执行命令：
pyinstaller easy-audit.spec
