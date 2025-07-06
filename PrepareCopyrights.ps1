Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Select-FolderDialog {
    param([string]$description = "Select a folder")

    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
    $dialog.Title = $description
    $dialog.Filter = "Folder|."
    $dialog.ValidateNames = $false
    $dialog.CheckFileExists = $false
    $dialog.CheckPathExists = $true
    $dialog.FileName = "Select Folder"

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return [System.IO.Path]::GetDirectoryName($dialog.FileName)
    } else {
        exit
    }
}

# Load or create environment config
$scriptPath = $MyInvocation.MyCommand.Path
$scriptDir = Split-Path $scriptPath
$envFile = Join-Path $scriptDir "CopyrightEnv.txt"

if (!(Test-Path $envFile)) {
    [System.Windows.Forms.MessageBox]::Show("Environment config file not found. You'll be prompted to define it now.", "Setup", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    $excelPath = Join-Path (Select-FolderDialog "Select directory to save Excel Log (Copyrighted.xlsx will be created)") "Copyrighted.xlsx"
    $zipPath   = Select-FolderDialog "Select directory to save ZIP files"

    Add-Type -AssemblyName Microsoft.VisualBasic
    $txtLog = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name for the text log file (e.g. Copyright-Data.txt)", "Text Log Filename", "Copyright-Data.txt")

    # Set temp path automatically
    $tempPath = Join-Path (Split-Path $excelPath) "TempDir-AutoDelete"

    @"
ExcelLogPath=$excelPath
ZipSavePath=$zipPath
TextLogFileName=$txtLog
TempJPEGPath=$tempPath
"@ | Set-Content -Path $envFile -Encoding UTF8
}

# Load config
$config = @{}
Get-Content $envFile | ForEach-Object {
    if ($_ -match "^\s*([^#].*?)\s*=\s*(.+?)\s*$") {
        $config[$matches[1]] = $matches[2] -replace '%TEMP%', $env:TEMP
    }
}

$logPath = $config["ExcelLogPath"]
$zipBasePath = $config["ZipSavePath"]
$textLogFileName = $config["TextLogFileName"]
$tempJpegFolder = $config["TempJPEGPath"]

function Get-ExifData {
    param($filePath)
    $shell = New-Object -ComObject Shell.Application
    $folder = $shell.Namespace((Split-Path $filePath))
    $item = $folder.ParseName((Split-Path $filePath -Leaf))

    $tags = @{
        "Date Taken" = 12
        "Camera Make" = 271
        "Camera Model" = 272
        "Focal Length" = 37386
        "Exposure Time" = 33434
        "ISO" = 34855
    }

    $results = @{}
    foreach ($key in $tags.Keys) {
        $results[$key] = $folder.GetDetailsOf($item, $tags[$key])
    }

    try {
        $img = [System.Drawing.Image]::FromFile($filePath)
        $results["Width"] = $img.Width
        $results["Height"] = $img.Height
        $img.Dispose()
    } catch {
        $results["Width"] = ""
        $results["Height"] = ""
    }

    return $results
}

# Constants
$maxPerZip = 750
$today = Get-Date -Format "MM-dd-yy"

if (!(Test-Path $zipBasePath)) {
    New-Item -ItemType Directory -Path $zipBasePath | Out-Null
}

# Select folder to process
$sourceFolder = Select-FolderDialog "Select folder containing images to process"

$imageExtensions = @(".dng", ".jpg", ".jpeg")
$allFiles = Get-ChildItem -Path $sourceFolder -Recurse -File
$imageFiles = $allFiles | Where-Object { $imageExtensions -contains $_.Extension.ToLower() }

if ($imageFiles.Count -eq 0) {
    Write-Host "No image files found." -ForegroundColor Yellow
    exit
}

# Excel setup
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

while ($true) {
    try {
        if (Test-Path $logPath) {
            $workbook = $excel.Workbooks.Open($logPath)
        } else {
            $workbook = $excel.Workbooks.Add()
            $ws = $workbook.Sheets.Item(1)
            $ws.Name = "Copyright Log"
            $headers = @("DNG Filename", "JPEG Filename", "Original Path", "Date Added", "ZIP Archive", "Date Taken", "Camera Make", "Camera Model", "Focal Length", "Exposure Time", "ISO", "Width", "Height")
            for ($i = 0; $i -lt $headers.Count; $i++) {
                $ws.Cells.Item(1, $i + 1) = $headers[$i]
            }
            $workbook.SaveAs($logPath)
        }
        break
    } catch {
        $choice = [System.Windows.Forms.MessageBox]::Show("Excel file is locked for editing.`nPlease close:`n$logPath", "Excel Locked", [System.Windows.Forms.MessageBoxButtons]::RetryCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($choice -ne "Retry") {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            exit
        }
    }
}

$ws = $workbook.Sheets.Item(1)
$row = $ws.UsedRange.Rows.Count + 1

# Prepare temp folder
if (Test-Path $tempJpegFolder) { Remove-Item $tempJpegFolder -Recurse -Force }
New-Item $tempJpegFolder -ItemType Directory | Out-Null

$chunks = [math]::Ceiling($imageFiles.Count / $maxPerZip)
$zipLogMap = @{}

for ($i = 0; $i -lt $chunks; $i++) {
    $startIndex = $i * $maxPerZip
    $chunk = $imageFiles | Select-Object -Skip $startIndex -First $maxPerZip

    $seq = 1
    do {
        $zipName = "$today---$seq.zip"
        $zipPath = Join-Path $zipBasePath $zipName
        $seq++
    } while (Test-Path $zipPath)

    $zipArchive = [System.IO.Compression.ZipFile]::Open($zipPath, "Update")

    foreach ($file in $chunk) {
        $ext = $file.Extension.ToLower()
        if ($imageExtensions -contains $ext) {
            $jpegName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name) + ".jpg"
            $jpegTempPath = Join-Path $tempJpegFolder $jpegName
            $dngName = if ($ext -eq ".dng") { $file.Name } else { "" }

            & magick "$($file.FullName)" -resize 2000x -density 200 -quality 92 "$jpegTempPath"

            if (!(Test-Path $jpegTempPath)) {
                Write-Warning "Conversion failed: $($file.Name)"
                continue
            }

            $entry = $zipArchive.CreateEntry($jpegName)
            $entryStream = $entry.Open()
            $fileStream = [System.IO.File]::OpenRead($jpegTempPath)
            $fileStream.CopyTo($entryStream)
            $fileStream.Close()
            $entryStream.Close()

            $exif = Get-ExifData $jpegTempPath

            $ws.Cells.Item($row, 1) = $dngName
            $ws.Cells.Item($row, 2) = $jpegName
            $ws.Cells.Item($row, 3) = $file.FullName
            $ws.Cells.Item($row, 4) = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            $ws.Cells.Item($row, 5) = $zipName
            $ws.Cells.Item($row, 6) = $exif["Date Taken"]
            $ws.Cells.Item($row, 7) = $exif["Camera Make"]
            $ws.Cells.Item($row, 8) = $exif["Camera Model"]
            $ws.Cells.Item($row, 9) = $exif["Focal Length"]
            $ws.Cells.Item($row, 10) = $exif["Exposure Time"]
            $ws.Cells.Item($row, 11) = $exif["ISO"]
            $ws.Cells.Item($row, 12) = $exif["Width"]
            $ws.Cells.Item($row, 13) = $exif["Height"]
            $row++

            if (-not $zipLogMap.ContainsKey($zipName)) {
                $zipLogMap[$zipName] = @()
            }
            $zipLogMap[$zipName] += $file.Name

            Remove-Item $jpegTempPath -Force
        }
    }

    $zipArchive.Dispose()
}

Remove-Item $tempJpegFolder -Recurse -Force

# Save and close Excel
$workbook.Save()
$workbook.Close($true)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# Write TXT log
$txtLogPath = Join-Path $sourceFolder $textLogFileName
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

$logContent = @()
$logContent += "==============================="
$logContent += "Processing Date: $timestamp"
$logContent += "Folder: $sourceFolder"
$logContent += ""
$logContent += "ZIP Files Created:"
$logContent += ""
foreach ($zip in $zipLogMap.Keys | Sort-Object) {
    $logContent += "  - $zip"
}
$logContent += ""
$logContent += "Total Images Processed: $($imageFiles.Count)"
$logContent += ""
$logContent += "Image ZIP Mapping:"
foreach ($zip in $zipLogMap.Keys | Sort-Object) {
    $logContent += "  [$zip]"
    foreach ($img in $zipLogMap[$zip]) {
        $logContent += "    - $img"
    }
}
$logContent += ""

Add-Content -Path $txtLogPath -Value $logContent

Write-Host "`n✅ All images processed, zipped, logged to Excel, and logged to TXT." -ForegroundColor Green
