$printerName = "Microsoft Print to PDF"
$zipUrl = "Needed URL for printtopdf.zip"
$installDir = "C:\Install\PDFtoPrint"
$zipFile = "C:\Install\printtopdf.zip"

try {
    $existingPrinter = Get-Printer -Name $printerName -ErrorAction SilentlyContinue
    if ($existingPrinter) {
        Write-Output "Printer '$printerName' already installed. Aborting installation."
        exit 0
    }
} catch {
    Write-Error "Failed to check existing printers: $_"
    exit 1
}

try {
    if (-not (Test-Path -Path "C:\Install")) {
        New-Item -ItemType Directory -Path "C:\Install" -Force | Out-Null
    }
} catch {
    Write-Error "Failed to create base install directory: $_"
    exit 1
}

try {
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipFile -ErrorAction Stop
} catch {
    Write-Error "Download failed: $_"
    exit 1
}

try {
    Expand-Archive -Path $zipFile -DestinationPath $installDir -Force -ErrorAction Stop
} catch {
    Write-Error "Extraction failed: $_"
    exit 1
}

try {
    Start-Process -FilePath "rundll32.exe" -ArgumentList 'printui.dll,PrintUIEntry /if /f "C:\Install\PDFtoPrint\prnms009.inf" /r "PORTPROMPT:" /m "Microsoft Print To PDF" /b "Microsoft Print to PDF" /u /Y' -Wait -NoNewWindow
} catch {
    Write-Error "Installation command failed: $_"
    exit 1
}
