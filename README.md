# PDF Merger

A lightweight Windows tool for merging PDF files. Built with PowerShell + WPF — no dependencies beyond Ghostscript (auto-downloaded if missing).

## Features

- Drag-and-drop or browse to add PDFs
- Reorder files with Up/Down buttons
- First-page thumbnail previews with list/grid toggle
- Merges via Ghostscript (`pdfwrite` device)
- Auto-downloads Ghostscript if not installed

## Install

Paste this into PowerShell (UAC will prompt for admin):

```powershell
$d="$env:TEMP\PDFMerger_setup";md $d -Force >$null;'install.ps1','PDFMerger.ps1'|%{iwr "https://raw.githubusercontent.com/patrickhannon86/PDFMerger/main/$_" -OutFile "$d\$_"};powershell -ExecutionPolicy Bypass -File "$d\install.ps1"
```

This will:
1. Copy `PDFMerger.ps1` to `C:\Program Files\PDFMerger\`
2. Download Ghostscript if not already installed
3. Generate an app icon
4. Create a **PDF Merger** shortcut on your desktop

### From a cloned repo

```powershell
powershell -ExecutionPolicy Bypass -File install.ps1
```

## Run without installing

```powershell
powershell -ExecutionPolicy Bypass -File PDFMerger.ps1
```

## Requirements

- Windows 10/11
- PowerShell 5.1+
- [Ghostscript](https://ghostscript.com/) — downloaded automatically on first merge if not found
