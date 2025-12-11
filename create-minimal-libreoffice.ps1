# Strip LibreOffice Portable to minimal size for DAS
# Keeps only essential files for headless PDF conversion (should be ~100-150MB)

$source = "portable\libreoffice"
$dest = "portable\libreoffice-minimal"

if (-not (Test-Path $source)) {
    Write-Host "Error: Source LibreOffice not found at $source" -ForegroundColor Red
    exit 1
}

Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "Creating Minimal LibreOffice Build" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

# Remove destination if exists
if (Test-Path $dest) {
    Write-Host "Removing old minimal build..." -ForegroundColor Yellow
    Remove-Item -Recurse -Force $dest
}

# Create directory structure
New-Item -ItemType Directory -Path $dest -Force | Out-Null

# CORE ESSENTIAL FILES for headless conversion
Write-Host "[1/4] Copying core program files..." -ForegroundColor Yellow

# Program directory essentials
$programFiles = @(
    "soffice.exe",
    "soffice.bin", 
    "soffice.com",
    "fundamental.ini",
    "uno.ini",
    "unorc",
    "bootstrap.ini",
    "versionrc",
    "*.dll"  # All DLLs needed
)

foreach ($file in $programFiles) {
    $items = Get-ChildItem -Path "$source\program\$file" -ErrorAction SilentlyContinue
    foreach ($item in $items) {
        $destPath = "$dest\program\$($item.Name)"
        $destDir = Split-Path -Parent $destPath
        New-Item -ItemType Directory -Path $destDir -Force -ErrorAction SilentlyContinue | Out-Null
        Copy-Item -Path $item.FullName -Destination $destPath -Force
    }
}

# CRITICAL: URE (UNO Runtime Environment) - Required for LibreOffice to run
Write-Host "[2/4] Copying URE (Runtime Environment)..." -ForegroundColor Yellow
if (Test-Path "$source\program\resource") {
    Copy-Item -Path "$source\program\resource" -Destination "$dest\program\resource" -Recurse -Force
}
if (Test-Path "$source\program\services") {
    Copy-Item -Path "$source\program\services" -Destination "$dest\program\services" -Recurse -Force  
}
if (Test-Path "$source\program\types") {
    Copy-Item -Path "$source\program\types" -Destination "$dest\program\types" -Recurse -Force
}

# FILTERS: Required for document format conversion
Write-Host "[3/4] Copying conversion filters..." -ForegroundColor Yellow
if (Test-Path "$source\program\filter") {
    Copy-Item -Path "$source\program\filter" -Destination "$dest\program\filter" -Recurse -Force
}

# SHARE: Configuration and registry (essential for conversion)
Write-Host "[4/4] Copying configuration..." -ForegroundColor Yellow

$shareDirs = @(
    "registry",
    "config", 
    "filter"
)

foreach ($dir in $shareDirs) {
    if (Test-Path "$source\share\$dir") {
        Copy-Item -Path "$source\share\$dir" -Destination "$dest\share\$dir" -Recurse -Force
    }
}

$originalSize = (Get-ChildItem -Path $source -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
$minimalSize = (Get-ChildItem -Path $dest -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum

$originalMB = [math]::Round($originalSize/1MB, 2)
$minimalMB = [math]::Round($minimalSize/1MB, 2)
$savings = [math]::Round(($originalSize - $minimalSize) / $originalSize * 100, 1)

Write-Host ""
Write-Host "=====================================" -ForegroundColor Green
Write-Host "Minimal LibreOffice Created!" -ForegroundColor Green
Write-Host "=====================================" -ForegroundColor Green
Write-Host ""
Write-Host "Original size: $originalMB MB" -ForegroundColor Yellow
Write-Host "Minimal size:  $minimalMB MB" -ForegroundColor Green
Write-Host "Savings:       $savings%" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next step:" -ForegroundColor Yellow
Write-Host "Update build.spec to use 'portable/libreoffice-minimal' instead" -ForegroundColor White
Write-Host ""
