# Split ZIP file into chunks for GitHub
# Each chunk will be under 25MB to comply with GitHub file size limits

$source = "C:\Users\pc\autoarendt\DAS-Setup.zip"
$chunkSize = 20MB  # 20MB chunks to stay safely under 25MB limit
$outputDir = "C:\Users\pc\autoarendt\chunks"

# Check if source file exists
if (-not (Test-Path $source)) {
    Write-Host "Error: Source file not found: $source" -ForegroundColor Red
    exit 1
}

# Create output directory
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
    Write-Host "Created directory: $outputDir" -ForegroundColor Green
}

Write-Host "Reading source file..." -ForegroundColor Yellow
$bytes = [System.IO.File]::ReadAllBytes($source)
$totalSize = $bytes.Length
$chunks = [math]::Ceiling($totalSize / $chunkSize)

Write-Host "File size: $([math]::Round($totalSize/1MB, 2)) MB" -ForegroundColor Cyan
Write-Host "Splitting into $chunks chunks..." -ForegroundColor Yellow
Write-Host ""

for($i=0; $i -lt $chunks; $i++) {
    $start = $i * $chunkSize
    $length = [math]::Min($chunkSize, $totalSize - $start)
    $chunk = New-Object byte[] $length
    [Array]::Copy($bytes, $start, $chunk, 0, $length)
    
    $outputFile = Join-Path $outputDir "DAS-Setup.zip.part$i"
    [System.IO.File]::WriteAllBytes($outputFile, $chunk)
    
    $chunkSizeMB = [math]::Round($length/1MB, 2)
    Write-Host "Created: DAS-Setup.zip.part$i ($chunkSizeMB MB)" -ForegroundColor Green
}

Write-Host ""
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "Split complete! Files in: $outputDir" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "To upload to GitHub:" -ForegroundColor Yellow
Write-Host "1. Commit all .part files from the chunks folder" -ForegroundColor White
Write-Host "2. Users download all parts and run rejoin-zip.ps1" -ForegroundColor White
