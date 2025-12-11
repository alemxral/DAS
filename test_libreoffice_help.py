"""Test LibreOffice with detailed output."""
import subprocess
import time

portable_path = r'c:\Users\pc\autoarendt\portable\libreoffice\program\soffice.exe'

print('Testing LibreOffice with help command (should be faster)...')
try:
    result = subprocess.run(
        [portable_path, '--help'],
        capture_output=True,
        timeout=10,
        text=True
    )
    print(f'Return code: {result.returncode}')
    print(f'stdout length: {len(result.stdout)} chars')
    print(f'stderr length: {len(result.stderr)} chars')
    
    if result.stdout:
        print('\nFirst 500 chars of stdout:')
        print(result.stdout[:500])
        
except subprocess.TimeoutExpired:
    print('[ERROR] Timed out')
except Exception as e:
    print(f'[ERROR] Error: {e}')
