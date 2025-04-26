# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

# 获取当前目录的绝对路径
current_dir = os.path.abspath(os.path.dirname('__file__'))

a = Analysis(
    ['src/gui/simple_main.py'],  # 主程序入口
    pathex=[
        current_dir,
        os.path.join(current_dir, 'src'),
        os.path.join(current_dir, 'src/gui'),
    ],
    binaries=[],
    datas=[
        ('resources/logo.png', 'resources'),
        ('resources/logo.ico', 'resources')
    ],
    hiddenimports=[
        'PIL._tkinter_finder',
        'PIL._imaging',
        'PIL._imagingft',
        'PIL._imagingmath',
        'PIL._imagingtk',
        'PIL._webp',
        'PIL.Image',
        'PIL.ImageTk',
        'openpyxl',
        'openpyxl.cell',
        'openpyxl.workbook',
        'openpyxl.writer',
        'openpyxl.styles',
        'xlrd',
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# 添加所有数据文件
a.datas += collect_data_files('PIL')

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='专业录取系统',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # 临时设置为True以查看错误信息
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join('resources', 'logo.ico'),  # 使用相对路径
    version='file_version_info.txt',
) 