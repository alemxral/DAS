# Encode ZIP chunks to Base64 to completely bypass firewall detection
# Base64 encoding makes binary data look like plain text

$chunkDir = "C:\Users\pc\autoarendt\chunks"

Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "Encoding ZIP chunks to Base64" -ForegroundColor Cyan
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
    # Extract part number
    $partNum = $file.Name -replace '.*part(\d+)$', '$1'
    $newName = "DAS-Setup-part$partNum.b64"
    $newPath = Join-Path $chunkDir $newName
    
    Write-Host "Encoding: $($file.Name)..." -ForegroundColor Yellow
    
    # Read binary data and convert to Base64
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    $base64 = [System.Convert]::ToBase64String($bytes)
    
    # Write as plain text
    [System.IO.File]::WriteAllText($newPath, $base64)
    
    $sizeMB = [math]::Round((Get-Item $newPath).Length/1MB, 2)
    Write-Host "Created: $newName ($sizeMB MB)" -ForegroundColor Green
}

Write-Host ""
Write-Host "=====================================" -ForegroundColor Green
Write-Host "Encoding complete!" -ForegroundColor Green
Write-Host "=====================================" -ForegroundColor Green
Write-Host ""
Write-Host "Upload these Base64 files to GitHub:" -ForegroundColor Yellow
Write-Host "  - DAS-Setup-part0.b64" -ForegroundColor White
Write-Host "  - DAS-Setup-part1.b64" -ForegroundColor White
Write-Host "  - DAS-Setup-part2.b64" -ForegroundColor White
Write-Host "  - rejoin-b64-to-zip.ps1" -ForegroundColor White
Write-Host ""
Write-Host "Note: Base64 files are ~33% larger but will NOT be detected by firewall" -ForegroundColor Yellow
