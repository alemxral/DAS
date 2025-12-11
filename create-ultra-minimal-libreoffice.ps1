# Ultra-minimal LibreOffice for DAS
# ONLY what's needed for soffice --headless --convert-to PDF
# Target: ~100-200MB

$source = "portable\libreoffice"
$dest = "portable\libreoffice-ultra"

if (-not (Test-Path $source)) {
    Write-Host "Error: Source not found" -ForegroundColor Red
    exit 1
}

Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "Ultra-Minimal LibreOffice (Headless)" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan

# Clean slate
if (Test-Path $dest) {
    Remove-Item -Recurse -Force $dest
}
New-Item -ItemType Directory -Path $dest -Force | Out-Null

Write-Host "[1/7] Core executables..." -ForegroundColor Yellow
Copy-Item "$source\program\soffice.exe" "$dest\program\soffice.exe" -Force
Copy-Item "$source\program\soffice.bin" "$dest\program\soffice.bin" -Force

Write-Host "[2/7] Required DLLs..." -ForegroundColor Yellow
# Only core runtime DLLs
$requiredDLLs = @(
    "sal3.dll",
    "uno*.dll",
    "cppu3.dll",
    "cppuhelper3.dll",
    "msvcr*.dll",
    "msvcp*.dll",
    "vcruntime*.dll",
    "tl*.dll",
    "utl*.dll",
    "comphelper*.dll",
    "i18nlangtag*.dll",
    "xmlreader*.dll",
    "store3.dll",
    "reg3.dll",
    "bootstrap*.dll"
)

foreach ($dll in $requiredDLLs) {
    Get-ChildItem "$source\program\$dll" -ErrorAction SilentlyContinue | ForEach-Object {
        Copy-Item $_.FullName "$dest\program\$($_.Name)" -Force
    }
}

Write-Host "[3/7] Config files..." -ForegroundColor Yellow
$configs = @("fundamental.ini", "uno.ini", "bootstrap.ini", "versionrc")
foreach ($cfg in $configs) {
    if (Test-Path "$source\program\$cfg") {
        Copy-Item "$source\program\$cfg" "$dest\program\$cfg" -Force
    }
}

Write-Host "[4/7] URE essentials..." -ForegroundColor Yellow
Copy-Item "$source\program\types" "$dest\program\types" -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "[5/7] Core filters..." -ForegroundColor Yellow
# Only PDF export filters
if (Test-Path "$source\program\filter") {
    New-Item -ItemType Directory "$dest\program\filter" -Force | Out-Null
    Copy-Item "$source\program\filter\*pdf*.dll" "$dest\program\filter\" -Force -ErrorAction SilentlyContinue
}

Write-Host "[6/7] Registry..." -ForegroundColor Yellow
Copy-Item "$source\share\registry" "$dest\share\registry" -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "[7/7] Minimal config..." -ForegroundColor Yellow  
# Bare minimum config
if (Test-Path "$source\share\config\soffice.cfg\modules") {
    New-Item -ItemType Directory "$dest\share\config\soffice.cfg\modules" -Force | Out-Null
    Copy-Item "$source\share\config\soffice.cfg\modules\*" "$dest\share\config\soffice.cfg\modules\" -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "=====================================" -ForegroundColor Green

$origSize = (Get-ChildItem -Path $source -Recurse -File | Measure-Object -Property Length -Sum).Sum / 1MB
$newSize = (Get-ChildItem -Path $dest -Recurse -File | Measure-Object -Property Length -Sum).Sum / 1MB
$savings = (($origSize - $newSize) / $origSize) * 100

Write-Host "Original:  $([math]::Round($origSize, 2)) MB" -ForegroundColor White
Write-Host "Ultra-min: $([math]::Round($newSize, 2)) MB" -ForegroundColor Green  
Write-Host "Saved:     $([math]::Round($savings, 1))%" -ForegroundColor Cyan
Write-Host ""
Write-Host "Update build.spec:" -ForegroundColor Yellow
Write-Host "  ('portable/libreoffice-ultra', 'portable/libreoffice')" -ForegroundColor White
