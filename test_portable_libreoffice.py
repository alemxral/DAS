"""Test portable LibreOffice integration."""
import sys
import os

# Add project to path
sys.path.insert(0, r'c:\Users\pc\autoarendt')

from services.format_converter import _get_portable_soffice_path, LIBREOFFICE_AVAILABLE

path = _get_portable_soffice_path()
print(f'Portable LibreOffice path: {path}')
print(f'Path exists: {os.path.exists(path)}')
print(f'LibreOffice available: {LIBREOFFICE_AVAILABLE}')

if LIBREOFFICE_AVAILABLE:
    print('\n[OK] LibreOffice is ready to use!')
else:
    print('\n[ERROR] LibreOffice not detected')
