@echo off
REM Test Runner Batch Script for Windows
REM Runs the Document Automation System test suite

echo ========================================
echo Document Automation System - Test Suite
echo ========================================
echo.

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found in PATH
    echo Please install Python or add it to your PATH
    pause
    exit /b 1
)

REM Activate virtual environment if it exists
if exist venv\Scripts\activate.bat (
    echo Activating virtual environment...
    call venv\Scripts\activate.bat
)

REM Check if pytest is installed
python -c "import pytest" >nul 2>&1
if errorlevel 1 (
    echo ERROR: pytest not installed
    echo Installing test dependencies...
    pip install pytest pytest-html pytest-cov psutil
)

REM Parse command line argument
set MODE=%1
if "%MODE%"=="" set MODE=all

echo Running tests in '%MODE%' mode...
echo.

REM Run tests
python run_tests.py %MODE%

REM Capture exit code
set EXIT_CODE=%ERRORLEVEL%

echo.
if %EXIT_CODE%==0 (
    echo ========================================
    echo Tests completed successfully!
    echo ========================================
) else (
    echo ========================================
    echo Tests failed with exit code: %EXIT_CODE%
    echo ========================================
)

REM Keep window open if double-clicked
if "%2"=="" pause

exit /b %EXIT_CODE%
