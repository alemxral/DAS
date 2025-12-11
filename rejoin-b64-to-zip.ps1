# Rejoin Base64 encoded files back to ZIP
# Decodes Base64 chunks and reassembles the ZIP file

$outputFile = "DAS-Setup.zip"
$b64Pattern = "DAS-Setup-part*.b64"

Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "Decoding and Rejoining DAS Setup" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

# Find all Base64 chunk files
$b64Files = Get-ChildItem -Filter $b64Pattern | Sort-Object {
    [int]($_.Name -replace '.*part(\d+)\.b64$', '$1')
}

if ($b64Files.Count -eq 0) {
    Write-Host "Error: No Base64 chunk files found (DAS-Setup-part*.b64)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Make sure you downloaded all these files:" -ForegroundColor Yellow
    Write-Host "  - DAS-Setup-part0.b64" -ForegroundColor White
    Write-Host "  - DAS-Setup-part1.b64" -ForegroundColor White
    Write-Host "  - DAS-Setup-part2.b64" -ForegroundColor White
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "Found $($b64Files.Count) Base64 chunk files" -ForegroundColor Green
Write-Host ""

# Create output stream
$outputStream = [System.IO.File]::Create($outputFile)

try {
    foreach ($b64File in $b64Files) {
        Write-Host "Decoding: $($b64File.Name)..." -ForegroundColor Yellow
        
        # Read Base64 text and decode to binary
        $base64Text = [System.IO.File]::ReadAllText($b64File.FullName)
        $bytes = [System.Convert]::FromBase64String($base64Text)
        $outputStream.Write($bytes, 0, $bytes.Length)
        
        $sizeMB = [math]::Round($bytes.Length/1MB, 2)
        Write-Host "  Decoded $sizeMB MB" -ForegroundColor Gray
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
    Write-Host "1. Right-click DAS-Setup.zip and select 'Extract All...'" -ForegroundColor White
    Write-Host "2. Open the DocumentAutomation folder" -ForegroundColor White
    Write-Host "3. Double-click DocumentAutomation.exe to run" -ForegroundColor White
    Write-Host ""
    Write-Host "DAS will open in a window ready to use!" -ForegroundColor Green
    Write-Host ""
    
} catch {
    Write-Host ""
    Write-Host "Error: $_" -ForegroundColor Red
    Write-Host ""
} finally {
    $outputStream.Close()
}

Read-Host "Press Enter to exit"
