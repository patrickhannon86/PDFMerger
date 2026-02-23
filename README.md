# PDF Merger

A lightweight Windows tool for merging PDF files. Built with PowerShell + WPF — no dependencies beyond Ghostscript (auto-downloaded if missing).

## Features

- Drag-and-drop or browse to add PDFs
- Reorder files with Up/Down buttons
- First-page thumbnail previews with list/grid toggle
- Merges via Ghostscript (`pdfwrite` device)
- Auto-downloads Ghostscript if not installed

## Install

Run the install script from an elevated (admin) PowerShell — it will prompt UAC automatically if needed:

```powershell
powershell -ExecutionPolicy Bypass -File install.ps1
```

This will:
1. Copy `PDFMerger.ps1` to `C:\Program Files\PDFMerger\`
2. Generate an app icon
3. Create a **PDF Merger** shortcut on your desktop

## Run without installing

```powershell
powershell -ExecutionPolicy Bypass -File PDFMerger.ps1
```

## Requirements

- Windows 10/11
- PowerShell 5.1+
- [Ghostscript](https://ghostscript.com/) — downloaded automatically on first merge if not found
