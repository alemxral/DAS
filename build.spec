# -*- mode: python ; coding: utf-8 -*-
import os
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

# Get the project directory
project_dir = Path('.').absolute()

# Use ultra-minimal LibreOffice (114MB instead of 1.4GB)
def collect_libreoffice_files():
    """Collect ultra-minimal LibreOffice for headless PDF conversion."""
    lo_ultra = project_dir / 'portable' / 'libreoffice-ultra'
    if not lo_ultra.exists():
        print("WARNING: Ultra-minimal LibreOffice not found! Run create-ultra-minimal-libreoffice.ps1")
        return []
    
    files = []
    for item in lo_ultra.rglob('*'):
        if item.is_file():
            # Map libreoffice-ultra -> libreoffice in final build
            rel_path = item.relative_to(lo_ultra)
            dest_dir = str(Path('portable/libreoffice') / rel_path.parent).replace('\\', '/')
            files.append((str(item), dest_dir))
    
    print(f"Collected {len(files)} ultra-minimal LibreOffice files")
    return files

a = Analysis(
    ['main.py'],
    pathex=[str(project_dir)],
    binaries=[],
    datas=[
        # Include templates and static folders
        ('templates', 'templates'),
        ('static', 'static'),
        ('config', 'config'),
        # Include .env if exists
        ('.env', '.') if Path('.env').exists() else None,
    ] + collect_libreoffice_files(),
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
        'openpyxl.worksheet',
        'openpyxl.worksheet.dimensions',
        'openpyxl.utils',
        'openpyxl.utils.cell',
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
        'requests',
        'requests.adapters',
        'requests.auth',
        'requests.cookies',
        'requests.models',
        'requests.sessions',
        'urllib3',
        'PIL',
        'PIL.Image',
        'PIL.ImageDraw',
        'PIL.ImageFont',
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
