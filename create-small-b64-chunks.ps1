# Split ZIP into smaller chunks (15MB) then Base64 encode
# This ensures Base64 files stay under 25MB (15MB * 1.33 = 20MB)

$source = "C:\Users\pc\autoarendt\DAS-Setup.zip"
$chunkSize = 15MB  # 15MB chunks -> ~20MB after Base64
$outputDir = "C:\Users\pc\autoarendt\chunks-b64"

# Check if source file exists
if (-not (Test-Path $source)) {
    Write-Host "Error: Source file not found: $source" -ForegroundColor Red
    exit 1
}

# Create output directory
if (Test-Path $outputDir) {
    Remove-Item -Recurse -Force $outputDir
}
New-Item -ItemType Directory -Path $outputDir | Out-Null

Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "Creating Base64-encoded chunks" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

$bytes = [System.IO.File]::ReadAllBytes($source)
$totalSize = $bytes.Length
$chunks = [math]::Ceiling($totalSize / $chunkSize)

Write-Host "Source: $([math]::Round($totalSize/1MB, 2)) MB" -ForegroundColor Cyan
Write-Host "Creating $chunks chunks..." -ForegroundColor Yellow
Write-Host ""

for($i=0; $i -lt $chunks; $i++) {
    $start = $i * $chunkSize
    $length = [math]::Min($chunkSize, $totalSize - $start)
    $chunk = New-Object byte[] $length
    [Array]::Copy($bytes, $start, $chunk, 0, $length)
    
    # Convert to Base64
    $base64 = [System.Convert]::ToBase64String($chunk)
    
    # Save as .b64 file
    $outputFile = Join-Path $outputDir "DAS-Setup-part$i.b64"
    [System.IO.File]::WriteAllText($outputFile, $base64)
    
    $b64Size = [math]::Round((Get-Item $outputFile).Length/1MB, 2)
    Write-Host "Created: DAS-Setup-part$i.b64 ($b64Size MB)" -ForegroundColor Green
}

Write-Host ""
Write-Host "=====================================" -ForegroundColor Green
Write-Host "Complete! All chunks under 25MB" -ForegroundColor Green
Write-Host "=====================================" -ForegroundColor Green
Write-Host ""
Write-Host "Upload these files to GitHub Pages:" -ForegroundColor Yellow
Get-ChildItem $outputDir -Filter "*.b64" | ForEach-Object {
    Write-Host "  - $($_.Name)" -ForegroundColor White
}
Write-Host "  - rejoin-b64-to-zip.ps1" -ForegroundColor White
