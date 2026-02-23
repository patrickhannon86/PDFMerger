#requires -Version 5.1
# Install PDFMerger to Program Files and create a desktop shortcut.
# Run: powershell -ExecutionPolicy Bypass -File install.ps1

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# --- Self-elevate if not admin ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe `
        "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

$installDir   = 'C:\Program Files\PDFMerger'
$scriptSrc    = Join-Path $PSScriptRoot 'PDFMerger.ps1'
$scriptDest   = Join-Path $installDir   'PDFMerger.ps1'
$iconDest     = Join-Path $installDir   'PDFMerger.ico'
$desktop      = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop 'PDF Merger.lnk'

# --- Copy script ---
Write-Host "Installing to $installDir ..."
New-Item -ItemType Directory -Path $installDir -Force | Out-Null
Copy-Item -Path $scriptSrc -Destination $scriptDest -Force

# --- Ensure Ghostscript is available ---
# Source the app script to get Find-Ghostscript / Install-Ghostscript
$binSrc = Join-Path $PSScriptRoot 'bin'
if (Test-Path $binSrc) {
    Write-Host 'Copying bundled Ghostscript ...'
    Copy-Item -Path $binSrc -Destination $installDir -Recurse -Force
}

# Point $binDir at the install location so Install-Ghostscript writes there
$binDir = Join-Path $installDir 'bin'
$scriptDir = $installDir

# Extract Find-Ghostscript & Install-Ghostscript from the app without running it
$ast = [System.Management.Automation.Language.Parser]::ParseFile(
    $scriptSrc, [ref]$null, [ref]$null)
$ast.FindAll({
    $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
    $args[0].Name -in @('Find-Ghostscript', 'Install-Ghostscript')
}, $false) | ForEach-Object { . ([scriptblock]::Create($_.Extent.Text)) }

$gs = Find-Ghostscript
if (-not $gs) {
    Write-Host 'Ghostscript not found — downloading ...'
    $gs = Install-Ghostscript
    if ($gs) {
        Write-Host "Ghostscript installed to $binDir" -ForegroundColor Green
    } else {
        Write-Host 'WARNING: Could not download Ghostscript. Thumbnails will not work until GS is installed.' -ForegroundColor Yellow
    }
} else {
    Write-Host "Ghostscript found: $gs"
}

# --- Generate .ico from the same WPF drawing the app uses ---
function New-PdfIcon {
    param([string]$outPath, [int]$size = 48)

    $scale = $size / 16.0
    $dv = New-Object System.Windows.Media.DrawingVisual
    $dc = $dv.RenderOpen()
    $dc.PushTransform([System.Windows.Media.ScaleTransform]::new($scale, $scale))

    # Page background
    $dc.DrawRoundedRectangle(
        [System.Windows.Media.Brushes]::White,
        (New-Object System.Windows.Media.Pen ([System.Windows.Media.Brushes]::Gray), 0.5),
        (New-Object System.Windows.Rect 2, 0, 12, 16), 1, 1)
    # Red banner
    $dc.DrawRectangle(
        (New-Object System.Windows.Media.SolidColorBrush (
            [System.Windows.Media.Color]::FromRgb(220, 50, 50))),
        $null,
        (New-Object System.Windows.Rect 2, 3, 12, 6))
    # "PDF" text
    $tf = New-Object System.Windows.Media.FormattedText(
        'PDF',
        [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Windows.FlowDirection]::LeftToRight,
        (New-Object System.Windows.Media.Typeface 'Segoe UI'),
        5,
        [System.Windows.Media.Brushes]::White,
        1.0)
    $dc.DrawText($tf, (New-Object System.Windows.Point 3.2, 3))

    $dc.Pop()
    $dc.Close()

    $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap `
        $size, $size, 96, 96, ([System.Windows.Media.PixelFormats]::Pbgra32)
    $rtb.Render($dv)

    # Encode bitmap as PNG bytes
    $encoder = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
    $encoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
    $pngStream = New-Object System.IO.MemoryStream
    $encoder.Save($pngStream)
    $pngBytes = $pngStream.ToArray()
    $pngStream.Dispose()

    # Build ICO: header (6 B) + one directory entry (16 B) + PNG payload
    $ico = New-Object System.IO.MemoryStream
    $w   = New-Object System.IO.BinaryWriter $ico
    $w.Write([uint16]0)                       # Reserved
    $w.Write([uint16]1)                       # Type = ICO
    $w.Write([uint16]1)                       # Image count
    $w.Write([byte]$size)                     # Width
    $w.Write([byte]$size)                     # Height
    $w.Write([byte]0)                         # Palette colours
    $w.Write([byte]0)                         # Reserved
    $w.Write([uint16]1)                       # Colour planes
    $w.Write([uint16]32)                      # Bits per pixel
    $w.Write([uint32]$pngBytes.Length)         # PNG data size
    $w.Write([uint32]22)                      # Offset to PNG data
    $w.Write($pngBytes)
    $w.Flush()
    [System.IO.File]::WriteAllBytes($outPath, $ico.ToArray())
    $w.Dispose(); $ico.Dispose()
}

Write-Host 'Generating icon ...'
New-PdfIcon -outPath $iconDest -size 48

# --- Generate silent launcher (no console flash) ---
$launchVbs = Join-Path $installDir 'launch.vbs'
$vbsContent = "CreateObject(""WScript.Shell"").Run ""powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """"$scriptDest"""""", 0, False"
[System.IO.File]::WriteAllText($launchVbs, $vbsContent)

# --- Create desktop shortcut ---
Write-Host 'Creating desktop shortcut ...'
$ws = New-Object -ComObject WScript.Shell
$sc = $ws.CreateShortcut($shortcutPath)
$sc.TargetPath       = 'wscript.exe'
$sc.Arguments        = "`"$launchVbs`""
$sc.WorkingDirectory = $installDir
$sc.IconLocation     = $iconDest
$sc.Save()

Write-Host ''
Write-Host 'Done! PDF Merger shortcut is on your desktop.' -ForegroundColor Green
