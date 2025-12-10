@echo off
REM Build script for DAS (Document Automation System)
REM This script builds the standalone executable using PyInstaller

echo ====================================
echo Building DAS Executable
echo ====================================
echo.

REM Check if virtual environment exists
if not exist "venv\Scripts\activate.bat" (
    echo Error: Virtual environment not found!
    echo Please create a virtual environment first with: python -m venv venv
    pause
    exit /b 1
)

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat

REM Check if PyInstaller is installed
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo Error: PyInstaller not found!
    echo Installing PyInstaller...
    pip install pyinstaller
    if errorlevel 1 (
        echo Failed to install PyInstaller
        pause
        exit /b 1
    )
)

REM Clean previous builds
echo.
echo Cleaning previous builds...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist

REM Build the executable
echo.
echo Building executable with PyInstaller...
echo This may take several minutes...
echo.
pyinstaller --clean build.spec

REM Check if build was successful
if errorlevel 1 (
    echo.
    echo ====================================
    echo Build FAILED!
    echo ====================================
    pause
    exit /b 1
)

REM Check if executable exists
if exist "dist\DocumentAutomation\DocumentAutomation.exe" (
    echo.
    echo ====================================
    echo Build SUCCESSFUL!
    echo ====================================
    echo.
    echo Executable location: dist\DocumentAutomation\DocumentAutomation.exe
    echo.
    echo You can now run the application by executing:
    echo   dist\DocumentAutomation\DocumentAutomation.exe
    echo.
    echo Or copy the entire dist\DocumentAutomation folder to another machine.
    echo.
) else (
    echo.
    echo ====================================
    echo Build completed but executable not found!
    echo ====================================
)

pause
