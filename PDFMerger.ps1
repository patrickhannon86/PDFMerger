#requires -Version 5.1
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# --- Ghostscript Detection & Auto-Download ---
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$binDir    = Join-Path $scriptDir 'bin'

function Find-Ghostscript {
    # 1. Check local bin folder first
    $localGs = Get-Item "$binDir\gs*\bin\gswin64c.exe" -ErrorAction SilentlyContinue |
        Sort-Object FullName -Descending | Select-Object -First 1
    if ($localGs) { return $localGs.FullName }

    # 2. Check registry
    $regPath = 'HKLM:\SOFTWARE\GPL Ghostscript'
    if (Test-Path $regPath) {
        $versions = Get-ChildItem $regPath | Sort-Object Name -Descending
        foreach ($ver in $versions) {
            $installDir = (Get-ItemProperty $ver.PSPath -ErrorAction SilentlyContinue).GS_DLL
            if ($installDir) {
                $gsExe = Join-Path (Split-Path $installDir) 'gswin64c.exe'
                if (Test-Path $gsExe) { return $gsExe }
            }
        }
    }

    # 3. Scan Program Files
    $candidates = @(
        "$env:ProgramFiles\gs\*\bin\gswin64c.exe"
        "${env:ProgramFiles(x86)}\gs\*\bin\gswin64c.exe"
    )
    foreach ($pattern in $candidates) {
        $found = Get-Item $pattern -ErrorAction SilentlyContinue |
            Sort-Object FullName -Descending | Select-Object -First 1
        if ($found) { return $found.FullName }
    }

    return $null
}

function Install-Ghostscript {
    $tempDir = Join-Path $env:TEMP "pdfmerger_setup_$(Get-Random)"
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

    try {
        # --- Get 7-Zip so we can extract the NSIS installer without elevation ---
        $sevenZipExe = $null

        # Check if 7z is already installed
        foreach ($candidate in @(
            "$env:ProgramFiles\7-Zip\7z.exe"
            "${env:ProgramFiles(x86)}\7-Zip\7z.exe"
        )) {
            if (Test-Path $candidate) { $sevenZipExe = $candidate; break }
        }

        # If not installed, download the 7-Zip MSI and extract 7z.exe from it
        if (-not $sevenZipExe) {
            $msiUrl  = 'https://www.7-zip.org/a/7z2408-x64.msi'
            $msiPath = Join-Path $tempDir '7z.msi'
            Invoke-WebRequest -Uri $msiUrl -OutFile $msiPath -UseBasicParsing

            $sevenZipDir = Join-Path $tempDir '7zip'
            # msiexec /a extracts an MSI to a folder — no admin needed
            Start-Process -FilePath 'msiexec.exe' `
                -ArgumentList @('/a', $msiPath, '/qn', "TARGETDIR=$sevenZipDir") `
                -Wait -NoNewWindow
            $sevenZipExe = Join-Path $sevenZipDir 'Files\7-Zip\7z.exe'
            if (-not (Test-Path $sevenZipExe)) { return $null }
        }

        # --- Download the Ghostscript NSIS installer ---
        $apiUrl  = 'https://api.github.com/repos/ArtifexSoftware/ghostpdl-downloads/releases/latest'
        $release = Invoke-RestMethod -Uri $apiUrl -UseBasicParsing
        $asset   = $release.assets | Where-Object { $_.name -match 'w64\.exe$' } | Select-Object -First 1
        if (-not $asset) { return $null }

        $installerPath = Join-Path $tempDir $asset.name
        Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $installerPath -UseBasicParsing

        # --- Extract the NSIS installer with 7-Zip (no elevation required) ---
        if (-not (Test-Path $binDir)) { New-Item -ItemType Directory -Path $binDir -Force | Out-Null }

        $extractDir = Join-Path $tempDir 'gs_extracted'
        Start-Process -FilePath $sevenZipExe `
            -ArgumentList @('x', $installerPath, "-o$extractDir", '-y') `
            -Wait -WindowStyle Hidden

        # The NSIS extraction puts files at the root — move them into bin\gs
        $gsDestDir = Join-Path $binDir 'gs'
        if (Test-Path $gsDestDir) { Remove-Item $gsDestDir -Recurse -Force }
        Move-Item -Path $extractDir -Destination $gsDestDir -Force

        # Find the exe
        $gsExe = Get-Item "$gsDestDir\bin\gswin64c.exe" -ErrorAction SilentlyContinue
        if ($gsExe) { return $gsExe.FullName }
        return $null
    }
    catch {
        return $null
    }
    finally {
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# --- XAML UI ---
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="PDF Merger" Height="520" Width="640" MinHeight="400" MinWidth="520"
    WindowStartupLocation="CenterScreen" Background="#F0F2F5"
    FontFamily="Segoe UI" FontSize="13">
    <Window.Resources>
        <!-- Animated base button style: rounded, smooth hover/press transitions -->
        <Style x:Key="BtnBase" TargetType="Button">
            <Setter Property="Height" Value="32"/>
            <Setter Property="Padding" Value="14,0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Background" Value="#E4E6EB"/>
            <Setter Property="Foreground" Value="#1C1E21"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}"
                                BorderThickness="0" SnapsToDevicePixels="True">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <EventTrigger RoutedEvent="MouseEnter">
                                <BeginStoryboard>
                                    <Storyboard>
                                        <ColorAnimation Storyboard.TargetName="border"
                                            Storyboard.TargetProperty="(Border.Background).(SolidColorBrush.Color)"
                                            To="#D2D5DB" Duration="0:0:0.15"/>
                                    </Storyboard>
                                </BeginStoryboard>
                            </EventTrigger>
                            <EventTrigger RoutedEvent="MouseLeave">
                                <BeginStoryboard>
                                    <Storyboard>
                                        <ColorAnimation Storyboard.TargetName="border"
                                            Storyboard.TargetProperty="(Border.Background).(SolidColorBrush.Color)"
                                            To="#E4E6EB" Duration="0:0:0.2"/>
                                    </Storyboard>
                                </BeginStoryboard>
                            </EventTrigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#BEC2C9"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Accent (Merge) button style with animated hover -->
        <Style x:Key="BtnAccent" TargetType="Button">
            <Setter Property="Height" Value="32"/>
            <Setter Property="Padding" Value="14,0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}"
                                SnapsToDevicePixels="True">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <EventTrigger RoutedEvent="MouseEnter">
                                <BeginStoryboard>
                                    <Storyboard>
                                        <ColorAnimation Storyboard.TargetName="border"
                                            Storyboard.TargetProperty="(Border.Background).(SolidColorBrush.Color)"
                                            To="#106EBE" Duration="0:0:0.15"/>
                                    </Storyboard>
                                </BeginStoryboard>
                            </EventTrigger>
                            <EventTrigger RoutedEvent="MouseLeave">
                                <BeginStoryboard>
                                    <Storyboard>
                                        <ColorAnimation Storyboard.TargetName="border"
                                            Storyboard.TargetProperty="(Border.Background).(SolidColorBrush.Color)"
                                            To="#0078D4" Duration="0:0:0.2"/>
                                    </Storyboard>
                                </BeginStoryboard>
                            </EventTrigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#005A9E"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Alternating row colors + hover highlight for the ListBox -->
        <Style TargetType="ListBoxItem">
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Style.Triggers>
                <Trigger Property="ItemsControl.AlternationIndex" Value="1">
                    <Setter Property="Background" Value="#F7F8FA"/>
                </Trigger>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#E8F0FE"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>

    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header with file count badge -->
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
            <TextBlock FontSize="22" FontWeight="SemiBold" Foreground="#1a1a2e"
                       VerticalAlignment="Center">
                <Run Text="&#x1F4C4; "/>PDF Merger
            </TextBlock>
            <Border Name="FileBadge" Background="#0078D4" CornerRadius="10"
                    Padding="8,2" Margin="10,0,0,0" VerticalAlignment="Center"
                    Visibility="Collapsed">
                <TextBlock Name="FileBadgeText" Foreground="White" FontSize="11"
                           FontWeight="SemiBold"/>
            </Border>
        </StackPanel>

        <!-- File list with drop shadow -->
        <Border Grid.Row="1" CornerRadius="8" Background="White" Margin="0,0,0,10"
                ClipToBounds="False" BorderThickness="0">
            <Border.Effect>
                <DropShadowEffect Color="#000000" BlurRadius="12" ShadowDepth="2"
                                  Opacity="0.08" Direction="270"/>
            </Border.Effect>
            <Border CornerRadius="8" BorderBrush="#DDE0E4" BorderThickness="1"
                    ClipToBounds="True">
                <ListBox Name="FileList" SelectionMode="Extended" AllowDrop="True"
                         AlternationCount="2" BorderThickness="0" Background="Transparent"
                         ScrollViewer.HorizontalScrollBarVisibility="Auto">
                    <ListBox.ItemTemplate>
                        <DataTemplate>
                            <TextBlock ToolTip="{Binding}" Padding="2,0">
                                <Run Text="{Binding Mode=OneWay}" />
                            </TextBlock>
                        </DataTemplate>
                    </ListBox.ItemTemplate>
                </ListBox>
            </Border>
        </Border>

        <!-- Buttons -->
        <WrapPanel Grid.Row="2" Margin="0,0,0,10" HorizontalAlignment="Left">
            <Button Name="BtnAdd"    Style="{StaticResource BtnBase}"   Margin="0,0,6,0">&#x2795; Add Files</Button>
            <Button Name="BtnRemove" Style="{StaticResource BtnBase}"   Margin="0,0,6,0">&#x2716; Remove</Button>
            <Button Name="BtnUp"     Style="{StaticResource BtnBase}"   Margin="0,0,6,0">&#x25B2; Up</Button>
            <Button Name="BtnDown"   Style="{StaticResource BtnBase}"   Margin="0,0,6,0">&#x25BC; Down</Button>
            <Button Name="BtnMerge"  Style="{StaticResource BtnAccent}" Margin="0,0,0,0">&#x1F500; Merge</Button>
        </WrapPanel>

        <!-- Progress bar (hidden by default) -->
        <ProgressBar Grid.Row="3" Name="ProgressBar" Height="4" Margin="0,0,0,8"
                     IsIndeterminate="False" Visibility="Collapsed"
                     Foreground="#0078D4" Background="#E4E6EB" BorderThickness="0"/>

        <!-- Status bar -->
        <Border Grid.Row="4" Background="#E8EAF0" CornerRadius="6" Padding="10,6">
            <TextBlock Name="StatusText" Text="Ready. Add PDF files to begin."
                       Foreground="#555" FontSize="12"/>
        </Border>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Grab controls
$fileList    = $window.FindName('FileList')
$btnAdd      = $window.FindName('BtnAdd')
$btnRemove   = $window.FindName('BtnRemove')
$btnUp       = $window.FindName('BtnUp')
$btnDown     = $window.FindName('BtnDown')
$btnMerge    = $window.FindName('BtnMerge')
$statusText    = $window.FindName('StatusText')
$progressBar   = $window.FindName('ProgressBar')
$fileBadge     = $window.FindName('FileBadge')
$fileBadgeText = $window.FindName('FileBadgeText')

# Store full paths; display short names with full path as tooltip
$pdfFiles = [System.Collections.Generic.List[string]]::new()

function Refresh-FileList {
    $fileList.Items.Clear()
    foreach ($path in $pdfFiles) {
        $item = New-Object System.Windows.Controls.ListBoxItem
        $item.Content = [System.IO.Path]::GetFileName($path)
        $item.ToolTip = $path
        $item.Tag     = $path
        $fileList.Items.Add($item) | Out-Null
    }
    # Update file count badge
    if ($pdfFiles.Count -gt 0) {
        $fileBadgeText.Text = "$($pdfFiles.Count) file$(if ($pdfFiles.Count -ne 1) {'s'})"
        $fileBadge.Visibility = 'Visible'
    } else {
        $fileBadge.Visibility = 'Collapsed'
    }
}

# --- Drag-and-drop support ---
$fileList.Add_DragEnter({
    param($sender, $e)
    if ($e.Data.GetDataPresent([Windows.DataFormats]::FileDrop)) {
        $files = $e.Data.GetData([Windows.DataFormats]::FileDrop)
        $hasPdf = $files | Where-Object { $_ -match '\.pdf$' }
        if ($hasPdf) {
            $e.Effects = [Windows.DragDropEffects]::Copy
        } else {
            $e.Effects = [Windows.DragDropEffects]::None
        }
    } else {
        $e.Effects = [Windows.DragDropEffects]::None
    }
    $e.Handled = $true
})

$fileList.Add_Drop({
    param($sender, $e)
    if ($e.Data.GetDataPresent([Windows.DataFormats]::FileDrop)) {
        $files = $e.Data.GetData([Windows.DataFormats]::FileDrop)
        $added = 0
        foreach ($f in ($files | Sort-Object)) {
            if ($f -match '\.pdf$' -and (Test-Path $f)) {
                $pdfFiles.Add($f)
                $added++
            }
        }
        if ($added -gt 0) {
            Refresh-FileList
            $statusText.Text = "Added $added file(s). Total: $($pdfFiles.Count)"
        }
    }
})

# --- Button Handlers ---

# Add Files
$btnAdd.Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Title = 'Select PDF files'
    $dlg.Filter = 'PDF Files (*.pdf)|*.pdf'
    $dlg.Multiselect = $true
    if ($dlg.ShowDialog($window)) {
        foreach ($f in $dlg.FileNames) { $pdfFiles.Add($f) }
        Refresh-FileList
        $statusText.Text = "Added $($dlg.FileNames.Count) file(s). Total: $($pdfFiles.Count)"
    }
})

# Remove Selected
$btnRemove.Add_Click({
    $selected = @($fileList.SelectedItems)
    if ($selected.Count -eq 0) {
        $statusText.Text = 'Select files to remove.'
        return
    }
    foreach ($item in $selected) {
        $pdfFiles.Remove($item.Tag) | Out-Null
    }
    Refresh-FileList
    $statusText.Text = "Removed $($selected.Count) file(s). Total: $($pdfFiles.Count)"
})

# Move Up
$btnUp.Add_Click({
    $idx = $fileList.SelectedIndex
    if ($idx -lt 1) { return }
    $item = $pdfFiles[$idx]
    $pdfFiles.RemoveAt($idx)
    $pdfFiles.Insert($idx - 1, $item)
    Refresh-FileList
    $fileList.SelectedIndex = $idx - 1
})

# Move Down
$btnDown.Add_Click({
    $idx = $fileList.SelectedIndex
    if ($idx -lt 0 -or $idx -ge $pdfFiles.Count - 1) { return }
    $item = $pdfFiles[$idx]
    $pdfFiles.RemoveAt($idx)
    $pdfFiles.Insert($idx + 1, $item)
    Refresh-FileList
    $fileList.SelectedIndex = $idx + 1
})

# Merge
$btnMerge.Add_Click({
    if ($pdfFiles.Count -lt 2) {
        $statusText.Text = 'Add at least 2 PDF files to merge.'
        return
    }

    # Find Ghostscript — auto-download if missing
    $gs = Find-Ghostscript
    if (-not $gs) {
        $statusText.Text = 'Ghostscript not found. Downloading...'
        $progressBar.IsIndeterminate = $true
        $progressBar.Visibility = 'Visible'
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
            [System.Windows.Threading.DispatcherPriority]::Background, [Action]{ })

        $gs = Install-Ghostscript
        $window.Cursor = $null
        $progressBar.IsIndeterminate = $false
        $progressBar.Visibility = 'Collapsed'
        if (-not $gs) {
            $statusText.Text = 'Failed to download Ghostscript. Install it manually or place gswin64c.exe in bin\.'
            return
        }
        $statusText.Text = "Ghostscript installed to bin folder. Merging..."
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
            [System.Windows.Threading.DispatcherPriority]::Background, [Action]{ })
    }

    # Ask where to save
    $saveDlg = New-Object Microsoft.Win32.SaveFileDialog
    $saveDlg.Title = 'Save merged PDF'
    $saveDlg.Filter = 'PDF Files (*.pdf)|*.pdf'
    $saveDlg.DefaultExt = '.pdf'
    $saveDlg.FileName = 'merged.pdf'
    if (-not $saveDlg.ShowDialog($window)) { return }

    $outputFile = $saveDlg.FileName
    $statusText.Text = 'Merging...'
    $progressBar.IsIndeterminate = $true
    $progressBar.Visibility = 'Visible'
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background, [Action]{ })

    try {
        $inputFiles = $pdfFiles | ForEach-Object { "`"$_`"" }
        $args = @(
            '-dBATCH'
            '-dNOPAUSE'
            '-q'
            '-sDEVICE=pdfwrite'
            "-sOutputFile=`"$outputFile`""
        ) + $inputFiles

        $proc = Start-Process -FilePath $gs -ArgumentList $args `
            -NoNewWindow -Wait -PassThru -RedirectStandardError "$env:TEMP\gs_err.txt"

        if ($proc.ExitCode -eq 0) {
            $statusText.Text = "Merged $($pdfFiles.Count) files into: $outputFile"
        } else {
            $errMsg = ''
            if (Test-Path "$env:TEMP\gs_err.txt") {
                $errMsg = Get-Content "$env:TEMP\gs_err.txt" -Raw
            }
            $statusText.Text = "Ghostscript error (exit $($proc.ExitCode)). $errMsg".Trim()
        }
    }
    catch {
        $statusText.Text = "Error: $($_.Exception.Message)"
    }
    finally {
        $window.Cursor = $null
        $progressBar.IsIndeterminate = $false
        $progressBar.Visibility = 'Collapsed'
        if (Test-Path "$env:TEMP\gs_err.txt" -ErrorAction SilentlyContinue) {
            Remove-Item "$env:TEMP\gs_err.txt" -Force -ErrorAction SilentlyContinue
        }
    }
})

# --- Window Icon (generated at runtime) ---
# Draw a simple red PDF icon using WPF drawing
$iconVisual = New-Object System.Windows.Media.DrawingVisual
$dc = $iconVisual.RenderOpen()
# Page background
$dc.DrawRoundedRectangle(
    [System.Windows.Media.Brushes]::White,
    (New-Object System.Windows.Media.Pen ([System.Windows.Media.Brushes]::Gray), 0.5),
    (New-Object System.Windows.Rect 2, 0, 12, 16), 1, 1)
# Red banner
$dc.DrawRectangle(
    (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(220, 50, 50))),
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
$dc.Close()
$rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap 16, 16, 96, 96, ([System.Windows.Media.PixelFormats]::Pbgra32)
$rtb.Render($iconVisual)
$window.Icon = $rtb

# --- Show Window ---
$window.ShowDialog() | Out-Null
