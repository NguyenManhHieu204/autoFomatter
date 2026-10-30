# -*- mode: python ; coding: utf-8 -*-
"""
build_exe.spec: Cấu hình đóng gói PyInstaller cho ứng dụng Auto-Formatter
Tối ưu hóa khởi động tức thì (< 1s) với cấu trúc Onedir & Không dùng UPX
"""

import sys
import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# Thư mục gốc dự án
project_dir = os.path.abspath(SPECPATH)

# Dữ liệu đính kèm (Configs, Templates, Assets Icon, Từ điển)
datas = [
    (os.path.join(project_dir, 'config'), 'config'),
    (os.path.join(project_dir, 'templates'), 'templates'),
    (os.path.join(project_dir, 'assets'), 'assets'),
]

try:
    datas += collect_data_files('rapidocr_onnxruntime')
except Exception:
    pass

# Hidden imports cần thiết cho docx, pywin32, PyQt6 và OCR
hiddenimports = [
    'win32com',
    'win32com.client',
    'pythoncom',
    'win32timezone',
    'docx',
    'docx.oxml',
    'docx.oxml.ns',
    'docx.enum.text',
    'docx.enum.table',
    'PyQt6',
    'PyQt6.QtCore',
    'PyQt6.QtGui',
    'PyQt6.QtWidgets',
    'PyQt6.QtSvg',
    'PyQt6.QtSvgWidgets',
    'PyQt6.QtPdf',
    'PyQt6.QtPdfWidgets',
    'pypdf',
    'rapidocr_onnxruntime',
    'onnxruntime',
    'PIL',
    'cv2',
    'yaml',
    'shapely',
    'pyclipper',
]

a = Analysis(
    ['main.py'],
    pathex=[project_dir],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'scipy', 'pandas', 'IPython', 'notebook'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='AutoFormatter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join(project_dir, 'assets', 'app_icon.ico')
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='AutoFormatter'
)
