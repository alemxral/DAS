#!/bin/bash
# Build script for DAS (Document Automation System)
# This script builds the standalone executable using PyInstaller

echo "===================================="
echo "Building DAS Executable"
echo "===================================="
echo ""

# Check if virtual environment exists
if [ ! -f "venv/bin/activate" ] && [ ! -f "venv/Scripts/activate" ]; then
    echo "Error: Virtual environment not found!"
    echo "Please create a virtual environment first with: python -m venv venv"
    exit 1
fi

# Activate virtual environment (Linux/Mac)
if [ -f "venv/bin/activate" ]; then
    echo "Activating virtual environment..."
    source venv/bin/activate
# Activate virtual environment (Git Bash on Windows)
elif [ -f "venv/Scripts/activate" ]; then
    echo "Activating virtual environment..."
    source venv/Scripts/activate
fi

# Check if PyInstaller is installed
if ! python -c "import PyInstaller" 2>/dev/null; then
    echo "Error: PyInstaller not found!"
    echo "Installing PyInstaller..."
    pip install pyinstaller
    if [ $? -ne 0 ]; then
        echo "Failed to install PyInstaller"
        exit 1
    fi
fi

# Clean previous builds
echo ""
echo "Cleaning previous builds..."
rm -rf build dist

# Build the executable
echo ""
echo "Building executable with PyInstaller..."
echo "This may take several minutes..."
echo ""
pyinstaller --clean build.spec

# Check if build was successful
if [ $? -ne 0 ]; then
    echo ""
    echo "===================================="
    echo "Build FAILED!"
    echo "===================================="
    exit 1
fi

# Check if executable exists (Windows)
if [ -f "dist/DocumentAutomation/DocumentAutomation.exe" ]; then
    echo ""
    echo "===================================="
    echo "Build SUCCESSFUL!"
    echo "===================================="
    echo ""
    echo "Executable location: dist/DocumentAutomation/DocumentAutomation.exe"
    echo ""
    echo "You can now run the application by executing:"
    echo "  dist/DocumentAutomation/DocumentAutomation.exe"
    echo ""
    echo "Or copy the entire dist/DocumentAutomation folder to another machine."
    echo ""
# Check if executable exists (Linux/Mac)
elif [ -f "dist/DocumentAutomation/DocumentAutomation" ]; then
    echo ""
    echo "===================================="
    echo "Build SUCCESSFUL!"
    echo "===================================="
    echo ""
    echo "Executable location: dist/DocumentAutomation/DocumentAutomation"
    echo ""
    echo "You can now run the application by executing:"
    echo "  ./dist/DocumentAutomation/DocumentAutomation"
    echo ""
    echo "Or copy the entire dist/DocumentAutomation folder to another machine."
    echo ""
else
    echo ""
    echo "===================================="
    echo "Build completed but executable not found!"
    echo "===================================="
fi
