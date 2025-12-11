"""Debug LibreOffice detection."""
import subprocess
import os

portable_path = r'c:\Users\pc\autoarendt\portable\libreoffice\program\soffice.exe'

print(f'Testing: {portable_path}')
print(f'Exists: {os.path.exists(portable_path)}')

try:
    result = subprocess.run(
        [portable_path, '--version'],
        capture_output=True,
        timeout=5,
        text=True
    )
    print(f'\nReturn code: {result.returncode}')
    print(f'stdout: {repr(result.stdout)}')
    print(f'stderr: {repr(result.stderr)}')
    print(f'\nSuccess: {result.returncode == 0}')
except Exception as e:
    print(f'\nError: {e}')
