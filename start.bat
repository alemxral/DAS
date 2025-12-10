@echo off
REM Quick Start Script for Document Automation System

echo ============================================================
echo Document Automation System - Quick Start
echo ============================================================
echo.

REM Check if virtual environment exists
if not exist "venv" (
    echo [1/4] Creating virtual environment...
    python -m venv venv
    if errorlevel 1 (
        echo Error creating virtual environment!
        pause
        exit /b 1
    )
    echo Virtual environment created successfully!
    echo.
) else (
    echo [1/4] Virtual environment already exists.
    echo.
)

REM Activate virtual environment
echo [2/4] Activating virtual environment...
call venv\Scripts\activate.bat
if errorlevel 1 (
    echo Error activating virtual environment!
    pause
    exit /b 1
)
echo.

REM Install dependencies
echo [3/4] Installing dependencies...
echo This may take a few minutes...
pip install -r requirements.txt
if errorlevel 1 (
    echo Error installing dependencies!
    pause
    exit /b 1
)
echo Dependencies installed successfully!
echo.

REM Create .env file if it doesn't exist
if not exist ".env" (
    echo Creating .env file from template...
    copy .env.example .env
    echo .env file created. Please review and update if needed.
    echo.
)

REM Start the application
echo [4/4] Starting the application...
echo.
echo ============================================================
echo Server will start at http://localhost:5000
echo Press Ctrl+C to stop the server
echo ============================================================
echo.
python run.py

pause
