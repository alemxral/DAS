# Setup script for portable LibreOffice integration
# This extracts LibreOffice from MSI and configures it for portable use

Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "Setting up Portable LibreOffice" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

$portableDir = "C:\Users\pc\autoarendt\portable"
$msiFile = "$portableDir\LibreOffice.msi"
$extractDir = "$portableDir\libreoffice"

# Check if MSI exists
if (-not (Test-Path $msiFile)) {
    Write-Host "Error: LibreOffice.msi not found!" -ForegroundColor Red
    Write-Host "Downloading LibreOffice..." -ForegroundColor Yellow
    
    $url = "https://downloadarchive.documentfoundation.org/libreoffice/old/7.6.4.1/win/x86_64/LibreOffice_7.6.4.1_Win_x86-64.msi"
    try {
        Invoke-WebRequest -Uri $url -OutFile $msiFile -UseBasicParsing
        Write-Host "Download complete!" -ForegroundColor Green
    } catch {
        Write-Host "Download failed: $_" -ForegroundColor Red
        exit 1
    }
}

Write-Host "Extracting LibreOffice from MSI..." -ForegroundColor Yellow

# Create extraction directory
if (Test-Path $extractDir) {
    Write-Host "Removing old extraction..." -ForegroundColor Yellow
    Remove-Item -Recurse -Force $extractDir
}
New-Item -ItemType Directory -Path $extractDir | Out-Null

# Extract MSI using msiexec
$extractCmd = "msiexec /a `"$msiFile`" /qn TARGETDIR=`"$extractDir`""
Write-Host "Running: $extractCmd" -ForegroundColor Gray

$process = Start-Process -FilePath "msiexec" -ArgumentList "/a", "`"$msiFile`"", "/qn", "TARGETDIR=`"$extractDir`"" -Wait -PassThru -NoNewWindow

if ($process.ExitCode -eq 0) {
    Write-Host "Extraction successful!" -ForegroundColor Green
} else {
    Write-Host "Extraction failed with exit code: $($process.ExitCode)" -ForegroundColor Red
    exit 1
}

# Find soffice.exe
$sofficeExe = Get-ChildItem -Path $extractDir -Recurse -Filter "soffice.exe" | Select-Object -First 1

if ($sofficeExe) {
    Write-Host "" -ForegroundColor Green
    Write-Host "=====================================" -ForegroundColor Green
    Write-Host "Setup Complete!" -ForegroundColor Green
    Write-Host "=====================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "LibreOffice portable installed at:" -ForegroundColor Cyan
    Write-Host "  $($sofficeExe.FullName)" -ForegroundColor White
    Write-Host ""
    Write-Host "Testing LibreOffice..." -ForegroundColor Yellow
    
    $version = & $sofficeExe.FullName --version 2>&1
    Write-Host "Version: $version" -ForegroundColor Green
    
    # Clean up MSI
    Write-Host ""
    Write-Host "Cleaning up MSI file..." -ForegroundColor Yellow
    Remove-Item $msiFile -Force
    
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "1. Update format_converter.py to use portable path" -ForegroundColor White
    Write-Host "2. Test PDF conversion" -ForegroundColor White
    Write-Host "3. Rebuild with PyInstaller" -ForegroundColor White
    
} else {
    Write-Host "Error: soffice.exe not found after extraction" -ForegroundColor Red
    exit 1
}

Read-Host "Press Enter to exit"
