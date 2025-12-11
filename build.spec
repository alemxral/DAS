# -*- mode: python ; coding: utf-8 -*-
import os
from pathlib import Path

block_cipher = None

# Get the project directory
project_dir = Path('.').absolute()

a = Analysis(
    ['main.py'],
    pathex=[str(project_dir)],
    binaries=[],
    datas=[
        # Include templates and static folders
        ('templates', 'templates'),
        ('static', 'static'),
        ('config', 'config'),
        # Include portable LibreOffice
        ('portable/libreoffice', 'portable/libreoffice'),
        # Include .env if exists
        ('.env', '.') if Path('.env').exists() else None,
    ],
    hiddenimports=[
        # Flask and extensions
        'flask',
        'flask_cors',
        'jinja2',
        'werkzeug',
        'click',
        'itsdangerous',
        'markupsafe',
        
        # PyWebView
        'webview',
        'webview.platforms',
        'webview.platforms.winforms',
        'clr',
        'System',
        'System.Windows.Forms',
        'System.Threading',
        
        # Office automation
        'win32com',
        'win32com.client',
        'pythoncom',
        'pywintypes',
        
        # Data processing
        'openpyxl',
        'openpyxl.cell',
        'openpyxl.cell.cell',
        'pandas',
        'numpy',
        
        # Document processing
        'PyPDF2',
        'docx',
        'docx.document',
        'docx.oxml',
        'docx.oxml.ns',
        'docx.shared',
        'reportlab',
        'reportlab.pdfgen',
        'reportlab.lib',
        
        # Other dependencies
        'dotenv',
        'pathlib',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Remove None entries from datas
a.datas = [item for item in a.datas if item is not None]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='DocumentAutomation',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # No console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='static/icon.png' if os.path.exists('static/icon.png') else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='DocumentAutomation',
)
