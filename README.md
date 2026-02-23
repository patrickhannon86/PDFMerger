# PDF Merger

A lightweight Windows tool for merging PDF files. Built with PowerShell + WPF — no dependencies beyond Ghostscript (auto-downloaded if missing).

## Features

- Drag-and-drop or browse to add PDFs
- Reorder files with Up/Down buttons
- First-page thumbnail previews with list/grid toggle
- Merges via Ghostscript (`pdfwrite` device)
- Auto-downloads Ghostscript if not installed

## Install

Open PowerShell and run:

```powershell
iwr https://raw.githubusercontent.com/patrickhannon86/PDFMerger/main/install.ps1 -OutFile "$env:TEMP\install-pdfmerger.ps1"
powershell -ExecutionPolicy Bypass -File "$env:TEMP\install-pdfmerger.ps1"
```

The installer will:

1. Copy `PDFMerger.ps1` to `C:\Program Files\PDFMerger\`
2. Download [Ghostscript](https://ghostscript.com/) if not already installed
3. Generate an app icon
4. Create a **PDF Merger** shortcut on your desktop

## Run without installing

```powershell
powershell -ExecutionPolicy Bypass -File PDFMerger.ps1
```

## Requirements

- Windows 10/11
- PowerShell 5.1+
- [Ghostscript](https://ghostscript.com/) — downloaded automatically on first merge if not found
