# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

PDFMerger — a tool for merging PDF files.

## Status

Implemented as a single-file PowerShell + WPF app (`PDFMerger.ps1`). Requires Ghostscript installed on the system.

## Running

```
powershell -ExecutionPolicy Bypass -File PDFMerger.ps1
```

## Architecture

- **Single file**: `PDFMerger.ps1` — WPF GUI defined in inline XAML, merge logic shells out to Ghostscript
- **Ghostscript detection**: Registry → Program Files scan → user browse dialog
- **Merge command**: `gswin64c.exe -dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite -sOutputFile="out.pdf" "in1.pdf" "in2.pdf" ...`
