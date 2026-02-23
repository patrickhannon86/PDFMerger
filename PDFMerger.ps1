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

        <!-- File list + sidebar buttons -->
        <Grid Grid.Row="1" Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>

            <!-- File list with drop shadow + empty state -->
            <Border Grid.Column="0" CornerRadius="8" Background="White"
                    ClipToBounds="False" BorderThickness="0">
                <Border.Effect>
                    <DropShadowEffect Color="#000000" BlurRadius="12" ShadowDepth="2"
                                      Opacity="0.08" Direction="270"/>
                </Border.Effect>
                <Grid>
                    <Border CornerRadius="8" BorderBrush="#DDE0E4" BorderThickness="1"
                            ClipToBounds="True">
                        <ListBox Name="FileList" SelectionMode="Extended" AllowDrop="True"
                                 AlternationCount="2" BorderThickness="0" Background="Transparent"
                                 ScrollViewer.HorizontalScrollBarVisibility="Auto"/>
                    </Border>
                    <!-- Empty state overlay (clickable) -->
                    <StackPanel Name="EmptyState" VerticalAlignment="Center"
                                HorizontalAlignment="Center" Cursor="Hand"
                                Background="Transparent" AllowDrop="True">
                        <TextBlock Text="&#x1F4C2;" FontSize="40" HorizontalAlignment="Center"
                                   Margin="0,0,0,8"/>
                        <TextBlock Text="Drag PDFs here" FontSize="15" FontWeight="SemiBold"
                                   Foreground="#888" HorizontalAlignment="Center"/>
                        <TextBlock Text="or click Add Files to get started"
                                   FontSize="12" Foreground="#AAA" HorizontalAlignment="Center"
                                   Margin="0,4,0,0"/>
                    </StackPanel>
                </Grid>
            </Border>

            <!-- Sidebar buttons -->
            <StackPanel Grid.Column="1" Margin="10,0,0,0" VerticalAlignment="Top">
                <!-- View mode toggle -->
                <StackPanel Orientation="Horizontal" Margin="0,0,0,8" HorizontalAlignment="Center">
                    <Border Name="BtnViewList" CornerRadius="4" Background="#E4E6EB"
                            Padding="8,4" Cursor="Hand">
                        <TextBlock Text="&#x2261;" FontSize="15" Foreground="#1C1E21"
                                   FontWeight="Bold"/>
                    </Border>
                    <Border Name="BtnViewPreview" CornerRadius="4" Background="#0078D4"
                            Padding="8,4" Margin="4,0,0,0" Cursor="Hand">
                        <TextBlock Text="&#x229E;" FontSize="15" Foreground="White"/>
                    </Border>
                </StackPanel>
                <Button Name="BtnAdd"    Style="{StaticResource BtnBase}"   Margin="0,0,0,6"
                        HorizontalContentAlignment="Left" Padding="10,0" Width="110">&#x2795; Add Files</Button>
                <Button Name="BtnRemove" Style="{StaticResource BtnBase}"   Margin="0,0,0,6"
                        HorizontalContentAlignment="Left" Padding="10,0" Width="110">&#x2716; Remove</Button>
                <Button Name="BtnUp"     Style="{StaticResource BtnBase}"   Margin="0,0,0,6"
                        HorizontalContentAlignment="Left" Padding="10,0" Width="110">&#x25B2; Up</Button>
                <Button Name="BtnDown"   Style="{StaticResource BtnBase}"   Margin="0,0,0,16"
                        HorizontalContentAlignment="Left" Padding="10,0" Width="110">&#x25BC; Down</Button>
                <Button Name="BtnMerge"  Style="{StaticResource BtnAccent}" Margin="0,0,0,0"
                        HorizontalContentAlignment="Center" Width="110">&#x1F500; Merge</Button>
            </StackPanel>
        </Grid>

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
$emptyState    = $window.FindName('EmptyState')
$btnViewList    = $window.FindName('BtnViewList')
$btnViewPreview = $window.FindName('BtnViewPreview')

# Store full paths; display short names with full path as tooltip
$pdfFiles = [System.Collections.Generic.List[string]]::new()

# Thumbnail state
$thumbDir = Join-Path $env:TEMP "pdfmerger_thumbs_$PID"
New-Item -ItemType Directory -Path $thumbDir -Force | Out-Null
$script:thumbCache = @{}
$script:viewMode = 'preview'
$script:gsPath = $null

function Generate-Thumbnail {
    param([string]$pdfPath)
    if ($script:thumbCache.ContainsKey($pdfPath)) { return $script:thumbCache[$pdfPath] }

    if (-not $script:gsPath) { $script:gsPath = Find-Ghostscript }
    if (-not $script:gsPath) { return $null }

    $hash = $pdfPath.GetHashCode().ToString('X8')
    $outPng = Join-Path $thumbDir "${hash}.png"

    if (Test-Path $outPng) {
        $script:thumbCache[$pdfPath] = $outPng
        return $outPng
    }

    try {
        $gsArgs = @(
            '-dNOPAUSE', '-dBATCH', '-dFirstPage=1', '-dLastPage=1',
            '-sDEVICE=png16m', '-r72',
            "-sOutputFile=`"$outPng`"",
            "`"$pdfPath`""
        )
        $proc = Start-Process -FilePath $script:gsPath -ArgumentList $gsArgs `
            -WindowStyle Hidden -Wait -PassThru
        if ($proc.ExitCode -eq 0 -and (Test-Path $outPng)) {
            $script:thumbCache[$pdfPath] = $outPng
            return $outPng
        }
    } catch {}
    return $null
}

function New-ThumbImage {
    param([string]$thumbPath)
    $img = New-Object System.Windows.Controls.Image
    $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
    $bmp.BeginInit()
    $bmp.UriSource = [Uri]::new($thumbPath)
    $bmp.DecodePixelWidth = 120
    $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    $bmp.EndInit()
    $img.Source = $bmp
    $img.Width  = 120
    $img.Margin = [System.Windows.Thickness]::new(0,0,8,0)
    return $img
}

function New-ThumbPlaceholder {
    $ph = New-Object System.Windows.Controls.Border
    $ph.Width  = 120
    $ph.Height = 155
    $ph.Background = [System.Windows.Media.SolidColorBrush]::new(
        [System.Windows.Media.Color]::FromRgb(0xF0, 0xF2, 0xF5))
    $ph.CornerRadius = [System.Windows.CornerRadius]::new(4)
    $ph.Margin = [System.Windows.Thickness]::new(0,0,8,0)

    $inner = New-Object System.Windows.Controls.StackPanel
    $inner.VerticalAlignment   = 'Center'
    $inner.HorizontalAlignment = 'Center'

    $icon = New-Object System.Windows.Controls.TextBlock
    $icon.Text = [char]::ConvertFromUtf32(0x1F4C4)
    $icon.FontSize = 24
    $icon.HorizontalAlignment = 'Center'
    $inner.Children.Add($icon) | Out-Null

    $bar = New-Object System.Windows.Controls.ProgressBar
    $bar.IsIndeterminate = $true
    $bar.Width  = 60
    $bar.Height = 4
    $bar.Foreground = [System.Windows.Media.SolidColorBrush]::new(
        [System.Windows.Media.Color]::FromRgb(0x00, 0x78, 0xD4))
    $bar.Background = [System.Windows.Media.SolidColorBrush]::new(
        [System.Windows.Media.Color]::FromRgb(0xE4, 0xE6, 0xEB))
    $bar.BorderThickness = [System.Windows.Thickness]::new(0)
    $bar.Margin = [System.Windows.Thickness]::new(0,4,0,0)
    $inner.Children.Add($bar) | Out-Null

    $ph.Child = $inner
    return $ph
}

function Refresh-FileList {
    $fileList.Items.Clear()
    $pendingThumbs = [System.Collections.Generic.List[int]]::new()

    foreach ($path in $pdfFiles) {
        $item = New-Object System.Windows.Controls.ListBoxItem
        $item.Tag     = $path
        $item.ToolTip = $path

        if ($script:viewMode -eq 'preview') {
            $sp = New-Object System.Windows.Controls.StackPanel
            $sp.Orientation = 'Horizontal'
            $sp.Margin = [System.Windows.Thickness]::new(2)

            $cached = $null
            if ($script:thumbCache.ContainsKey($path)) { $cached = $script:thumbCache[$path] }

            if ($cached -and (Test-Path $cached)) {
                $sp.Children.Add((New-ThumbImage $cached)) | Out-Null
            } else {
                $sp.Children.Add((New-ThumbPlaceholder)) | Out-Null
                $pendingThumbs.Add($fileList.Items.Count)
            }

            $textPanel = New-Object System.Windows.Controls.StackPanel
            $textPanel.VerticalAlignment = 'Center'
            $fnBlock = New-Object System.Windows.Controls.TextBlock
            $fnBlock.Text = [System.IO.Path]::GetFileName($path)
            $fnBlock.FontWeight = [System.Windows.FontWeights]::SemiBold
            $fnBlock.FontSize = 13
            $textPanel.Children.Add($fnBlock) | Out-Null
            $dirBlock = New-Object System.Windows.Controls.TextBlock
            $dirBlock.Text = [System.IO.Path]::GetDirectoryName($path)
            $dirBlock.FontSize = 11
            $dirBlock.Foreground = [System.Windows.Media.Brushes]::Gray
            $dirBlock.TextTrimming = 'CharacterEllipsis'
            $dirBlock.MaxWidth = 300
            $textPanel.Children.Add($dirBlock) | Out-Null
            $sp.Children.Add($textPanel) | Out-Null

            $item.Content = $sp
        } else {
            $item.Content = [System.IO.Path]::GetFileName($path)
        }

        $fileList.Items.Add($item) | Out-Null
    }

    # Update file count badge and empty state
    if ($pdfFiles.Count -gt 0) {
        $fileBadgeText.Text = "$($pdfFiles.Count) file$(if ($pdfFiles.Count -ne 1) {'s'})"
        $fileBadge.Visibility = 'Visible'
        $emptyState.Visibility = 'Collapsed'
    } else {
        $fileBadge.Visibility = 'Collapsed'
        $emptyState.Visibility = 'Visible'
    }

    # Generate pending thumbnails one-by-one, replacing placeholders as they finish
    if ($pendingThumbs.Count -gt 0) {
        # Flush so placeholders render before GS work starts
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
            [System.Windows.Threading.DispatcherPriority]::Background, [Action]{ })

        foreach ($idx in $pendingThumbs) {
            $listItem = $fileList.Items[$idx]
            $sp = $listItem.Content
            $thumbPath = Generate-Thumbnail $listItem.Tag

            if ($thumbPath -and (Test-Path $thumbPath)) {
                $sp.Children.RemoveAt(0)
                $sp.Children.Insert(0, (New-ThumbImage $thumbPath))
            } else {
                # Stop the spinner — show a static fallback
                $ph = $sp.Children[0]
                ($ph.Child.Children[1]).IsIndeterminate = $false
                ($ph.Child.Children[1]).Visibility = 'Collapsed'
            }

            # Let the UI repaint after each thumbnail
            [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
                [System.Windows.Threading.DispatcherPriority]::Background, [Action]{ })
        }
    }
}

# --- Empty state click to add files ---
$emptyState.Add_MouseLeftButtonDown({
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

# --- View toggle handlers ---
function Update-ViewToggle {
    if ($script:viewMode -eq 'list') {
        $btnViewList.Background = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(0x00, 0x78, 0xD4))
        ($btnViewList.Child).Foreground = [System.Windows.Media.Brushes]::White
        $btnViewPreview.Background = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(0xE4, 0xE6, 0xEB))
        ($btnViewPreview.Child).Foreground = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(0x1C, 0x1E, 0x21))
    } else {
        $btnViewPreview.Background = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(0x00, 0x78, 0xD4))
        ($btnViewPreview.Child).Foreground = [System.Windows.Media.Brushes]::White
        $btnViewList.Background = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(0xE4, 0xE6, 0xEB))
        ($btnViewList.Child).Foreground = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(0x1C, 0x1E, 0x21))
    }
}

$btnViewList.Add_MouseLeftButtonDown({
    if ($script:viewMode -ne 'list') {
        $script:viewMode = 'list'
        Update-ViewToggle
        Refresh-FileList
    }
})

$btnViewPreview.Add_MouseLeftButtonDown({
    if ($script:viewMode -ne 'preview') {
        $script:viewMode = 'preview'
        Update-ViewToggle
        Refresh-FileList
    }
})

# --- Drag-and-drop support (shared by FileList and EmptyState) ---
$onDragEnter = {
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
}

$onDrop = {
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
}

$fileList.Add_DragEnter($onDragEnter)
$fileList.Add_Drop($onDrop)
$emptyState.Add_DragEnter($onDragEnter)
$emptyState.Add_Drop($onDrop)

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

# --- Cleanup thumbnails on close ---
$window.Add_Closed({
    if (Test-Path $thumbDir) {
        Remove-Item $thumbDir -Recurse -Force -ErrorAction SilentlyContinue
    }
})

# --- Show Window ---
$window.ShowDialog() | Out-Null
