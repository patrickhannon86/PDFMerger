# Install PDFMerger to Program Files and create a desktop shortcut.
# Usage:  irm https://raw.githubusercontent.com/patrickhannon86/PDFMerger/main/install.ps1 | iex

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

$repo = 'https://raw.githubusercontent.com/patrickhannon86/PDFMerger/main'

# --- When piped via iex, save to disk so UAC elevation works ---
if (-not $PSCommandPath) {
    $tmp = Join-Path $env:TEMP 'install-pdfmerger.ps1'
    Invoke-WebRequest -Uri "$repo/install.ps1" -OutFile $tmp -UseBasicParsing
    powershell -ExecutionPolicy Bypass -File $tmp
    return
}

# --- Self-elevate if not admin ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe `
        "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}
$installDir   = 'C:\Program Files\PDFMerger'
$scriptDest   = Join-Path $installDir   'PDFMerger.ps1'
$iconDest     = Join-Path $installDir   'PDFMerger.ico'
$desktop      = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop 'PDF Merger.lnk'

# --- Get PDFMerger.ps1 (local repo copy or download latest) ---
New-Item -ItemType Directory -Path $installDir -Force | Out-Null

$localScript = Join-Path $PSScriptRoot 'PDFMerger.ps1'
$isRepoInstall = ($PSScriptRoot -ne $env:TEMP) -and (Test-Path $localScript)

if ($isRepoInstall) {
    Write-Host "Installing from local repo to $installDir ..."
    Copy-Item -Path $localScript -Destination $scriptDest -Force
    $scriptSrc = $localScript
} else {
    Write-Host "Downloading latest PDFMerger.ps1 to $installDir ..."
    Invoke-WebRequest -Uri "$repo/PDFMerger.ps1" -OutFile $scriptDest -UseBasicParsing
    $scriptSrc = $scriptDest
}

# --- Ensure Ghostscript is available ---
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
    # Red->Orange gradient banner
    $gradBrush = New-Object System.Windows.Media.LinearGradientBrush
    $gradBrush.StartPoint = [System.Windows.Point]::new(0, 0)
    $gradBrush.EndPoint   = [System.Windows.Point]::new(1, 0)
    $gradBrush.GradientStops.Add([System.Windows.Media.GradientStop]::new(
        [System.Windows.Media.Color]::FromRgb(220, 50, 50), 0))
    $gradBrush.GradientStops.Add([System.Windows.Media.GradientStop]::new(
        [System.Windows.Media.Color]::FromRgb(0, 120, 212), 1))
    $dc.DrawRectangle($gradBrush, $null,
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

# --- Generate context-merge.vbs (multi-select collector) ---
Write-Host 'Setting up Explorer context menu ...'
$contextVbs = Join-Path $installDir 'context-merge.vbs'
$contextVbsContent = @"
' context-merge.vbs — collects multi-select PDF paths and launches PDF Merger
Dim fso, importPath, lockPath, f
Set fso = CreateObject("Scripting.FileSystemObject")
importPath = fso.GetSpecialFolder(2) & "\pdfmerger_import.txt"
lockPath   = fso.GetSpecialFolder(2) & "\pdfmerger_import.lock"

' Append the file path passed via %1
Set f = fso.OpenTextFile(importPath, 8, True)
f.WriteLine WScript.Arguments(0)
f.Close

' Try to create lock file — first instance wins
On Error Resume Next
Set f = fso.CreateTextFile(lockPath, False)
If Err.Number <> 0 Then
    ' Another instance already has the lock — just exit
    WScript.Quit
End If
On Error GoTo 0
f.Close

' First instance: wait for other instances to finish appending
WScript.Sleep 500

' Launch PDF Merger
Dim launchVbs
launchVbs = fso.GetParentFolderName(WScript.ScriptFullName) & "\launch.vbs"
CreateObject("WScript.Shell").Run "wscript.exe """ & launchVbs & """", 0, False

' Clean up lock file
fso.DeleteFile lockPath, True
"@
[System.IO.File]::WriteAllText($contextVbs, $contextVbsContent)

# --- Register shell context menu for .pdf files ---
$regBase = 'HKCU:\Software\Classes\SystemFileAssociations\.pdf\shell\PDFMerger'
New-Item -Path "$regBase\command" -Force | Out-Null
Set-ItemProperty -Path $regBase -Name '(Default)' -Value 'Merge with PDF Merger'
Set-ItemProperty -Path $regBase -Name 'Icon' -Value $iconDest
Set-ItemProperty -Path $regBase -Name 'MultiSelectModel' -Value 'Player'
Set-ItemProperty -Path "$regBase\command" -Name '(Default)' `
    -Value "wscript.exe `"$contextVbs`" `"%1`""

# --- Copy uninstall script ---
$uninstallSrc  = Join-Path $PSScriptRoot 'uninstall.ps1'
$uninstallDest = Join-Path $installDir   'uninstall.ps1'
if (Test-Path $uninstallSrc) {
    Copy-Item -Path $uninstallSrc -Destination $uninstallDest -Force
    Write-Host 'Uninstall script copied.'
}

Write-Host ''
Write-Host 'Done! PDF Merger shortcut is on your desktop.' -ForegroundColor Green
Write-Host 'Right-click PDFs in Explorer to "Merge with PDF Merger".' -ForegroundColor Green
Write-Host ''
Read-Host 'Press Enter to close'
