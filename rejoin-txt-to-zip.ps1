# Rejoin TXT files back to ZIP
# Downloads TXT files, converts back to binary chunks, and rejoins to ZIP

$outputFile = "DAS-Setup.zip"
$txtPattern = "DAS-Setup-part*.txt"

Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "Rejoining DAS Setup from TXT files" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

# Find all TXT chunk files
$txtFiles = Get-ChildItem -Filter $txtPattern | Sort-Object {
    [int]($_.Name -replace '.*part(\d+)\.txt$', '$1')
}

if ($txtFiles.Count -eq 0) {
    Write-Host "Error: No TXT chunk files found (DAS-Setup-part*.txt)" -ForegroundColor Red
    Write-Host "Make sure you downloaded all these files:" -ForegroundColor Yellow
    Write-Host "  - DAS-Setup-part0.txt" -ForegroundColor White
    Write-Host "  - DAS-Setup-part1.txt" -ForegroundColor White
    Write-Host "  - DAS-Setup-part2.txt" -ForegroundColor White
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "Found $($txtFiles.Count) TXT chunk files" -ForegroundColor Green
Write-Host ""

# Create output stream
$outputStream = [System.IO.File]::Create($outputFile)

try {
    foreach ($txtFile in $txtFiles) {
        Write-Host "Processing: $($txtFile.Name)..." -ForegroundColor Yellow
        
        # Read the TXT file as raw bytes (they're actually binary data)
        $bytes = [System.IO.File]::ReadAllBytes($txtFile.FullName)
        $outputStream.Write($bytes, 0, $bytes.Length)
        
        $sizeMB = [math]::Round($bytes.Length/1MB, 2)
        Write-Host "  Added $sizeMB MB" -ForegroundColor Gray
    }
    
    Write-Host ""
    Write-Host "=====================================" -ForegroundColor Green
    Write-Host "Success! Created: $outputFile" -ForegroundColor Green
    Write-Host "=====================================" -ForegroundColor Green
    Write-Host ""
    
    $finalSize = [math]::Round((Get-Item $outputFile).Length/1MB, 2)
    Write-Host "Final ZIP size: $finalSize MB" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "1. Extract DAS-Setup.zip" -ForegroundColor White
    Write-Host "2. Open the DocumentAutomation folder" -ForegroundColor White
    Write-Host "3. Run DocumentAutomation.exe" -ForegroundColor White
    Write-Host ""
    
} catch {
    Write-Host ""
    Write-Host "Error: $_" -ForegroundColor Red
    Write-Host ""
} finally {
    $outputStream.Close()
}

Read-Host "Press Enter to exit"
