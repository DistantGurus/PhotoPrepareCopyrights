# PhotoPrepareCopyrights
Automates image conversion, metadata extraction, ZIP archiving, and Excel logging for photo copyright prep. Converts DNG/JPEG files, logs EXIF data, creates ZIPs with 750-image limits, and writes summary logs to text files for easy tracking.


# 📸 Copyright Image Archiver

A PowerShell utility to automate the organization, conversion, archiving, and logging of large batches of DNG/JPEG images into ZIP files — perfect for photo workflows that require copyright compliance and metadata logging.

---

## 🔧 Features

- Converts `.DNG`, `.JPG`, and `.JPEG` to standardized JPEGs (2000px width, 200 DPI, 92% quality).
- Archives processed images into ZIP files (max 750 images per archive).
- Auto-names ZIPs using the format `MM-DD-YY---#.zip`.
- Logs image metadata (EXIF and file paths) to an Excel `.xlsx` file.
- Creates a `Copyright-Data.txt` in the source folder summarizing:
  - All ZIPs created
  - Total images processed
  - A breakdown of which images were placed in which ZIP
- Automatically ignores video and unsupported file types.
- Automatically handles locked Excel files and retries until they’re closed.
- Uses a `TempDir-AutoDelete` subfolder (created and removed each run) for intermediate JPEGs.
- Configuration is stored in a simple environment file: `CopyrightEnv.txt`.

---

## 📁 Folder Setup Example
C:
└── Users
└── YourName
└── Documents
├── ImageLogs
│ ├── Copyrighted.xlsx
│ └── TempDir-AutoDelete
├── ImageZips
└── PhotosToProcess
├── IMG001.DNG
├── IMG002.JPG
└── Copyright-Data.txt


---

## 📄 `CopyrightEnv.txt`

This file is created automatically the first time you run the script, and stores all key paths for reuse:

```ini
ExcelLogPath=C:\Users\YourName\Documents\ImageLogs\Copyrighted.xlsx
ZipSavePath=C:\Users\YourName\Documents\ImageZips
TextLogFileName=Copyright-Data.txt
TempJPEGPath=C:\Users\YourName\Documents\ImageLogs\TempDir-AutoDelete




---

## 📄 `CopyrightEnv.txt`

This file is created automatically the first time you run the script, and stores all key paths for reuse:

```ini
ExcelLogPath=C:\Users\YourName\Documents\ImageLogs\Copyrighted.xlsx
ZipSavePath=C:\Users\YourName\Documents\ImageZips
TextLogFileName=Copyright-Data.txt
TempJPEGPath=C:\Users\YourName\Documents\ImageLogs\TempDir-AutoDelete
