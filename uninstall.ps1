# Uninstall PDFMerger -removes install dir, desktop shortcut, and context menu.

# --- Self-elevate if not admin ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe `
        "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

$installDir   = 'C:\Program Files\PDFMerger'
$desktop      = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop 'PDF Merger.lnk'
$regKey       = 'HKCU:\Software\Classes\SystemFileAssociations\.pdf\shell\PDFMerger'

# --- Remove install directory ---
if (Test-Path $installDir) {
    Remove-Item $installDir -Recurse -Force
    Write-Host "Removed $installDir" -ForegroundColor Green
} else {
    Write-Host "$installDir not found -skipping." -ForegroundColor Yellow
}

# --- Remove desktop shortcut ---
if (Test-Path $shortcutPath) {
    Remove-Item $shortcutPath -Force
    Write-Host "Removed desktop shortcut." -ForegroundColor Green
} else {
    Write-Host 'Desktop shortcut not found -skipping.' -ForegroundColor Yellow
}

# --- Remove context menu registry key ---
if (Test-Path $regKey) {
    Remove-Item $regKey -Recurse -Force
    Write-Host 'Removed Explorer context menu entry.' -ForegroundColor Green
} else {
    Write-Host 'Context menu entry not found -skipping.' -ForegroundColor Yellow
}

Write-Host ''
Write-Host 'PDF Merger has been uninstalled.' -ForegroundColor Green
Write-Host ''
