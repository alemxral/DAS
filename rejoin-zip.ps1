# Rejoin split ZIP file chunks
# Run this after downloading all DAS-Setup.zip.part* files

$outputFile = "DAS-Setup.zip"
$chunkPattern = "DAS-Setup.zip.part*"

Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "Rejoining DAS Setup ZIP" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

# Find all chunk files
$chunks = Get-ChildItem -Filter $chunkPattern | Sort-Object {
    [int]($_.Name -replace '.*part(\d+)$', '$1')
}

if ($chunks.Count -eq 0) {
    Write-Host "Error: No chunk files found (DAS-Setup.zip.part*)" -ForegroundColor Red
    Write-Host "Make sure all .part files are in the current directory" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "Found $($chunks.Count) chunks" -ForegroundColor Green
Write-Host ""

# Create output stream
$outputStream = [System.IO.File]::Create($outputFile)

try {
    foreach ($chunk in $chunks) {
        Write-Host "Processing: $($chunk.Name)..." -ForegroundColor Yellow
        $bytes = [System.IO.File]::ReadAllBytes($chunk.FullName)
        $outputStream.Write($bytes, 0, $bytes.Length)
    }
    
    Write-Host ""
    Write-Host "=====================================" -ForegroundColor Green
    Write-Host "Success! Created: $outputFile" -ForegroundColor Green
    Write-Host "=====================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Size: $([math]::Round((Get-Item $outputFile).Length/1MB, 2)) MB" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Extract the ZIP and run DocumentAutomation.exe" -ForegroundColor Yellow
    
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
} finally {
    $outputStream.Close()
}

Write-Host ""
Read-Host "Press Enter to exit"
