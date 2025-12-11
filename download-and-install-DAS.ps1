# Download and Install Document Automation System (DAS)
# Automatically downloads all Base64 chunks from GitHub Pages and reassembles the ZIP

$baseUrl = "https://alemxral.github.io/cv/"
$outputFile = "DAS-Setup.zip"
$tempDir = "DAS-Download-Temp"

Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "    Document Automation System (DAS) Installer" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host ""

# Create temporary directory for downloads
if (Test-Path $tempDir) {
    Write-Host "Cleaning up previous downloads..." -ForegroundColor Yellow
    Remove-Item -Recurse -Force $tempDir
}
New-Item -ItemType Directory -Path $tempDir | Out-Null

# Function to download a file with progress
function Download-File {
    param($url, $outputPath)
    try {
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($url, $outputPath)
        return $true
    } catch {
        return $false
    }
}

# First, try to detect how many chunks exist by trying to download them
Write-Host "Detecting available chunks..." -ForegroundColor Yellow
$chunkNumber = 0
$downloadedChunks = @()

while ($true) {
    $chunkUrl = "${baseUrl}DAS-Setup-part${chunkNumber}.b64"
    $chunkPath = Join-Path $tempDir "DAS-Setup-part${chunkNumber}.b64"
    
    Write-Host "Checking for chunk $chunkNumber..." -ForegroundColor Gray -NoNewline
    
    $success = Download-File -url $chunkUrl -outputPath $chunkPath
    
    if ($success) {
        $fileSize = [math]::Round((Get-Item $chunkPath).Length/1MB, 2)
        Write-Host " Found! ($fileSize MB)" -ForegroundColor Green
        $downloadedChunks += $chunkPath
        $chunkNumber++
    } else {
        Write-Host " Not found" -ForegroundColor Gray
        break
    }
}

Write-Host ""

if ($downloadedChunks.Count -eq 0) {
    Write-Host "======================================================" -ForegroundColor Red
    Write-Host "ERROR: No chunks found at $baseUrl" -ForegroundColor Red
    Write-Host "======================================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please verify:" -ForegroundColor Yellow
    Write-Host "1. The URL is correct" -ForegroundColor White
    Write-Host "2. The files are uploaded to GitHub Pages" -ForegroundColor White
    Write-Host "3. You have internet connection" -ForegroundColor White
    Write-Host ""
    
    # Cleanup
    Remove-Item -Recurse -Force $tempDir -ErrorAction SilentlyContinue
    
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "Successfully downloaded $($downloadedChunks.Count) chunks!" -ForegroundColor Green
Write-Host ""

# Decode and reassemble ZIP file
Write-Host "Decoding and reassembling DAS Setup..." -ForegroundColor Yellow
Write-Host ""

$outputStream = [System.IO.File]::Create($outputFile)

try {
    foreach ($chunkPath in $downloadedChunks) {
        $chunkName = [System.IO.Path]::GetFileName($chunkPath)
        Write-Host "  Processing: $chunkName..." -ForegroundColor Gray
        
        # Read Base64 text and decode to binary
        $base64Text = [System.IO.File]::ReadAllText($chunkPath)
        $bytes = [System.Convert]::FromBase64String($base64Text)
        $outputStream.Write($bytes, 0, $bytes.Length)
    }
    
    $outputStream.Close()
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Green
    Write-Host "    SUCCESS! DAS Setup Ready to Install" -ForegroundColor Green
    Write-Host "======================================================" -ForegroundColor Green
    Write-Host ""
    
    $finalSize = [math]::Round((Get-Item $outputFile).Length/1MB, 2)
    Write-Host "Created: $outputFile ($finalSize MB)" -ForegroundColor Cyan
    Write-Host ""
    
    # Cleanup temporary directory
    Write-Host "Cleaning up temporary files..." -ForegroundColor Yellow
    Remove-Item -Recurse -Force $tempDir -ErrorAction SilentlyContinue
    
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Yellow
    Write-Host "    Next Steps - Installation Instructions" -ForegroundColor Yellow
    Write-Host "======================================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "1. Extract the ZIP file:" -ForegroundColor White
    Write-Host "   - Right-click 'DAS-Setup.zip'" -ForegroundColor Gray
    Write-Host "   - Select 'Extract All...'" -ForegroundColor Gray
    Write-Host "   - Choose a location and click 'Extract'" -ForegroundColor Gray
    Write-Host ""
    Write-Host "2. Run the application:" -ForegroundColor White
    Write-Host "   - Open the 'DocumentAutomation' folder" -ForegroundColor Gray
    Write-Host "   - Double-click 'DocumentAutomation.exe'" -ForegroundColor Gray
    Write-Host ""
    Write-Host "3. Start using DAS!" -ForegroundColor White
    Write-Host "   - The application window will open" -ForegroundColor Gray
    Write-Host "   - Upload your templates and data files" -ForegroundColor Gray
    Write-Host "   - Generate your documents automatically" -ForegroundColor Gray
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Green
    Write-Host ""
    
    # Ask if user wants to extract now
    $extract = Read-Host "Would you like to extract the ZIP now? (Y/N)"
    
    if ($extract -eq "Y" -or $extract -eq "y") {
        Write-Host ""
        Write-Host "Extracting ZIP file..." -ForegroundColor Yellow
        
        try {
            # Use Windows built-in extraction
            $extractPath = Join-Path $PSScriptRoot "DocumentAutomation"
            
            if (Test-Path $extractPath) {
                Write-Host "Removing existing installation..." -ForegroundColor Yellow
                Remove-Item -Recurse -Force $extractPath
            }
            
            # Extract using .NET
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            [System.IO.Compression.ZipFile]::ExtractToDirectory($outputFile, $PSScriptRoot)
            
            Write-Host "Extraction complete!" -ForegroundColor Green
            Write-Host ""
            Write-Host "Application installed to: $extractPath" -ForegroundColor Cyan
            Write-Host ""
            
            # Ask if user wants to run now
            $run = Read-Host "Would you like to run DAS now? (Y/N)"
            
            if ($run -eq "Y" -or $run -eq "y") {
                $exePath = Join-Path $extractPath "DocumentAutomation.exe"
                if (Test-Path $exePath) {
                    Write-Host ""
                    Write-Host "Starting Document Automation System..." -ForegroundColor Green
                    Start-Process $exePath
                } else {
                    Write-Host "Error: Could not find DocumentAutomation.exe" -ForegroundColor Red
                }
            }
        } catch {
            Write-Host "Error extracting ZIP: $_" -ForegroundColor Red
            Write-Host "Please extract DAS-Setup.zip manually" -ForegroundColor Yellow
        }
    }
    
} catch {
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Red
    Write-Host "ERROR during reassembly" -ForegroundColor Red
    Write-Host "======================================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Error details: $_" -ForegroundColor Red
    Write-Host ""
    
    $outputStream.Close()
    
    # Cleanup
    Remove-Item -Force $outputFile -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force $tempDir -ErrorAction SilentlyContinue
    
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Read-Host "Press Enter to exit"
