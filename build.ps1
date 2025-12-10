# Build script for DAS (Document Automation System)
# This script builds the standalone executable using PyInstaller

Write-Host "====================================" -ForegroundColor Cyan
Write-Host "Building DAS Executable" -ForegroundColor Cyan
Write-Host "====================================" -ForegroundColor Cyan
Write-Host ""

# Check if virtual environment exists
if (-not (Test-Path "venv\Scripts\Activate.ps1")) {
    Write-Host "Error: Virtual environment not found!" -ForegroundColor Red
    Write-Host "Please create a virtual environment first with: python -m venv venv" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

# Activate virtual environment
Write-Host "Activating virtual environment..." -ForegroundColor Yellow
& "venv\Scripts\Activate.ps1"

# Check if PyInstaller is installed
try {
    python -c "import PyInstaller" 2>$null
    if ($LASTEXITCODE -ne 0) { throw }
} catch {
    Write-Host "Error: PyInstaller not found!" -ForegroundColor Red
    Write-Host "Installing PyInstaller..." -ForegroundColor Yellow
    pip install pyinstaller
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Failed to install PyInstaller" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# Clean previous builds
Write-Host ""
Write-Host "Cleaning previous builds..." -ForegroundColor Yellow
if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
if (Test-Path "dist") { Remove-Item -Recurse -Force "dist" }

# Build the executable
Write-Host ""
Write-Host "Building executable with PyInstaller..." -ForegroundColor Yellow
Write-Host "This may take several minutes..." -ForegroundColor Yellow
Write-Host ""
pyinstaller --clean build.spec

# Check if build was successful
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "====================================" -ForegroundColor Red
    Write-Host "Build FAILED!" -ForegroundColor Red
    Write-Host "====================================" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Check if executable exists
if (Test-Path "dist\DocumentAutomation\DocumentAutomation.exe") {
    Write-Host ""
    Write-Host "====================================" -ForegroundColor Green
    Write-Host "Build SUCCESSFUL!" -ForegroundColor Green
    Write-Host "====================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Executable location: dist\DocumentAutomation\DocumentAutomation.exe" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "You can now run the application by executing:" -ForegroundColor Yellow
    Write-Host "  dist\DocumentAutomation\DocumentAutomation.exe" -ForegroundColor White
    Write-Host ""
    Write-Host "Or copy the entire dist\DocumentAutomation folder to another machine." -ForegroundColor Yellow
    Write-Host ""
} else {
    Write-Host ""
    Write-Host "====================================" -ForegroundColor Red
    Write-Host "Build completed but executable not found!" -ForegroundColor Red
    Write-Host "====================================" -ForegroundColor Red
}

Read-Host "Press Enter to exit"
