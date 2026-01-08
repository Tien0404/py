# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for NRL Lookup Tool
"""

import os

block_cipher = None

# Đường dẫn hiện tại
CURRENT_DIR = os.path.dirname(os.path.abspath(SPEC))

a = Analysis(
    ['launcher.py'],
    pathex=[CURRENT_DIR],
    binaries=[],
    datas=[
        ('templates', 'templates'),  # Copy thư mục templates
    ],
    hiddenimports=[
        'flask',
        'werkzeug',
        'werkzeug.serving',
        'werkzeug.debug',
        'jinja2',
        'markupsafe',
        'itsdangerous',
        'click',
        'blinker',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.cell',
        'openpyxl.workbook',
        'openpyxl.worksheet',
        'et_xmlfile',
        'requests',
        'requests.adapters',
        'urllib3',
        'charset_normalizer',
        'certifi',
        'idna',
        'unidecode',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'PIL',
        'tkinter',
        'PyQt5',
        'PyQt6',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='NRL_Lookup',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # Hiển thị console để user thấy trạng thái
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Có thể thêm icon: icon='icon.ico'
)
