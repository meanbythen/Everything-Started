import sys
import os
from PyInstaller.building.build_main import Analysis, PYZ, EXE, COLLECT

# 增加递归限制
sys.setrecursionlimit(5000)

# 创建logs目录
if not os.path.exists('logs'):
    os.makedirs('logs')

# 获取openpyxl路径
import openpyxl
import et_xmlfile
openpyxl_path = os.path.dirname(openpyxl.__file__)
et_xmlfile_path = os.path.dirname(et_xmlfile.__file__)

a = Analysis(
    ['run.py'],  # 使用 run.py 作为入口程序
    pathex=[],
    binaries=[],
    datas=[
        ('icon.ico', '.'),
        ('logs', 'logs'),  # 包含logs目录
        (openpyxl_path, 'openpyxl'),  # 添加整个openpyxl包
        (et_xmlfile_path, 'et_xmlfile'),  # 添加openpyxl的依赖包
    ],
    hiddenimports=[
        'fitz',
        'tkinter',
        'json',
        'queue',
        'threading',
        'traceback',
        'logging',
        'invoice_gui',  # 添加主程序模块
        'get_coordinates',  # 添加发票处理模块
        'openpyxl',
        'openpyxl.cell',
        'openpyxl.cell.cell',
        'openpyxl.cell.read_only',
        'openpyxl.cell.text',
        'openpyxl.workbook',
        'openpyxl.workbook.workbook',
        'openpyxl.worksheet',
        'openpyxl.worksheet.worksheet',
        'openpyxl.worksheet.dimensions',
        'openpyxl.writer',
        'openpyxl.writer.excel',
        'openpyxl.styles',
        'openpyxl.styles.alignment',
        'openpyxl.styles.borders',
        'openpyxl.styles.fills',
        'openpyxl.styles.fonts',
        'openpyxl.styles.numbers',
        'openpyxl.styles.protection',
        'openpyxl.utils',
        'openpyxl.utils.cell',
        'openpyxl.utils.datetime',
        'openpyxl.utils.exceptions',
        'openpyxl.utils.units',
        'et_xmlfile',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pandas', 'PIL', 'pdf2image', 'pytesseract'],
    noarchive=False,
)

# 添加运行时钩子
rt_hooks = []
hook_contents = '''
import os
import sys

def _append_openpyxl_path():
    openpyxl_path = os.path.join(sys._MEIPASS, "openpyxl")
    if openpyxl_path not in sys.path:
        sys.path.append(openpyxl_path)
    et_xmlfile_path = os.path.join(sys._MEIPASS, "et_xmlfile")
    if et_xmlfile_path not in sys.path:
        sys.path.append(et_xmlfile_path)

_append_openpyxl_path()
'''

hook_file = 'openpyxl_hook.py'
with open(hook_file, 'w') as f:
    f.write(hook_contents)
rt_hooks.append(hook_file)

a.runtime_hooks = rt_hooks

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='发票处理程序',
    debug=False,  # 禁用调试模式
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # 禁用控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='发票处理程序',
) 