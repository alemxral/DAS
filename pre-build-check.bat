@echo off
REM Pre-build validation script - checks main.py safety before building exe
echo.
echo ======================================================================
echo PRE-BUILD VALIDATION
echo ======================================================================
echo.

REM Activate virtual environment
call venv\Scripts\activate.bat

echo [1/3] Running frozen environment tests...
python test_main_frozen.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [FAIL] Frozen environment tests failed!
    echo Please review test_main_frozen.py output above.
    pause
    exit /b 1
)

echo.
echo [2/3] Running license validation path tests...
python test_license_path.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [FAIL] License path tests failed!
    echo Please review test_license_path.py output above.
    pause
    exit /b 1
)

echo.
echo [3/3] Checking for Unicode characters in production files...
python -c "import sys; import re; files=['main.py','services/license_validator.py']; failed=[f for f in files if re.search(r'[^\x00-\x7F]',open(f,'r',encoding='utf-8').read())]; sys.exit(len(failed))"
if %ERRORLEVEL% NEQ 0 (
    echo [FAIL] Unicode characters found in production files!
    pause
    exit /b 1
) else (
    echo [OK] No problematic Unicode characters found
)

echo.
echo ======================================================================
echo [SUCCESS] All pre-build validation tests passed!
echo main.py is safe to build into executable.
echo ======================================================================
echo.
echo You can now run: build.bat
echo.
pause
