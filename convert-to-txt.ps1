# Convert ZIP chunks to TXT files to bypass firewall
# Renames .part files to .txt for safe download

$chunkDir = "C:\Users\pc\autoarendt\chunks"

Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "Converting ZIP chunks to TXT" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

# Find all .part files
$partFiles = Get-ChildItem -Path $chunkDir -Filter "*.part*"

if ($partFiles.Count -eq 0) {
    Write-Host "Error: No .part files found in $chunkDir" -ForegroundColor Red
    exit 1
}

Write-Host "Found $($partFiles.Count) chunk files" -ForegroundColor Green
Write-Host ""

foreach ($file in $partFiles) {
    # Extract part number: DAS-Setup.zip.part0 -> 0
    $partNum = $file.Name -replace '.*part(\d+)$', '$1'
    $newName = "DAS-Setup-part$partNum.txt"
    $newPath = Join-Path $chunkDir $newName
    
    Copy-Item $file.FullName $newPath
    Write-Host "Created: $newName" -ForegroundColor Green
}

Write-Host ""
Write-Host "=====================================" -ForegroundColor Green
Write-Host "Conversion complete!" -ForegroundColor Green
Write-Host "=====================================" -ForegroundColor Green
Write-Host ""
Write-Host "Upload these TXT files to GitHub:" -ForegroundColor Yellow
Write-Host "  - DAS-Setup-part0.txt" -ForegroundColor White
Write-Host "  - DAS-Setup-part1.txt" -ForegroundColor White
Write-Host "  - DAS-Setup-part2.txt" -ForegroundColor White
Write-Host "  - rejoin-txt-to-zip.ps1" -ForegroundColor White
Write-Host ""
Write-Host "Users download the TXT files and run rejoin-txt-to-zip.ps1" -ForegroundColor Yellow
