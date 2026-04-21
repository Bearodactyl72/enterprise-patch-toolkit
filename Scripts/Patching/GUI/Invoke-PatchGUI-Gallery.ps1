# DOTS formatting comment

<#
    .SYNOPSIS
        Theme picker for Invoke-PatchGUI.
    .DESCRIPTION
        Shows every theme defined in Themes.psd1 as a clickable card.
        Clicking a card writes the selected theme to preferences.json
        and launches Invoke-PatchGUI.ps1 in that theme.

        Heavily-customized themes (non-standard XAML) have matching
        files in ThemeOverrides\<ThemeKey>.ps1 that the production GUI
        delegates to; the Gallery preview is approximate for those.

        Written by Skyler Werner
        Date: 2026/04/16
        Version 2.0.0
    .EXAMPLE
        .\Invoke-PatchGUI-Gallery.ps1
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase


# ============================================================================
#  Load themes + preference helpers
# ============================================================================

$script:ThemesPath = Join-Path $PSScriptRoot 'Themes.psd1'
$script:PrefsDir   = Join-Path $env:APPDATA 'Patching'
$script:PrefsPath  = Join-Path $script:PrefsDir 'preferences.json'
$script:AllThemes  = Import-PowerShellDataFile -Path $script:ThemesPath
$script:PatchGUI   = Join-Path $PSScriptRoot 'Invoke-PatchGUI.ps1'

# Display order -- pairs Cobalt Slate dark with its Day sibling up top,
# then the rest roughly in sample-number order, jokes (98/99) at the end.
$script:DisplayOrder = @(
    'CobaltSlate', 'CobaltSlateDay',
    'TokyoNight', 'MonochromaticTeal', 'UltraDarkViolet',
    'IndustrialCharcoal', 'LightPurpleGray',
    'DarkForest', 'NavyAnalytics', 'PureMonochrome',
    'BlueCorporate', 'TaniumInspired',
    'LogitechTeal', '3DNetworkBn',
    'TacticalChrome', 'CyberCommand'
)

function Get-PatchPreferences {
    if (-not (Test-Path $script:PrefsPath)) { return @{} }
    try {
        $raw = Get-Content -Raw -Path $script:PrefsPath -Encoding UTF8
        $obj = $raw | ConvertFrom-Json
        $h = @{}
        foreach ($p in $obj.PSObject.Properties) { $h[$p.Name] = $p.Value }
        return $h
    } catch {
        return @{}
    }
}

function Set-PatchPreferences {
    param([hashtable]$Preferences)
    if (-not (Test-Path $script:PrefsDir)) {
        New-Item -ItemType Directory -Path $script:PrefsDir -Force | Out-Null
    }
    $Preferences | ConvertTo-Json | Set-Content -Path $script:PrefsPath -Encoding UTF8
}


# ============================================================================
#  Build XAML
# ============================================================================

[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Invoke-Patch Theme Gallery"
    Width="1280" Height="820"
    MinWidth="960" MinHeight="600"
    WindowStartupLocation="CenterScreen"
    Background="#0C0D14"
    FontFamily="Segoe UI">

    <Window.Resources>
        <SolidColorBrush x:Key="Bg"      Color="#0C0D14"/>
        <SolidColorBrush x:Key="Surface" Color="#181828"/>
        <SolidColorBrush x:Key="Border"  Color="#2A2D45"/>
        <SolidColorBrush x:Key="Text"    Color="#E8EAF0"/>
        <SolidColorBrush x:Key="SubText" Color="#7A7F9A"/>
        <SolidColorBrush x:Key="Accent"  Color="#8B70F0"/>

        <!-- ScrollBar -->
        <Style x:Key="ScrollThumb" TargetType="Thumb">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Thumb">
                        <Border CornerRadius="4" Background="#30334D" Margin="1"/>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="ScrollBar">
            <Setter Property="Width" Value="10"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollBar">
                        <Grid>
                            <Track x:Name="PART_Track" IsDirectionReversed="True">
                                <Track.Thumb>
                                    <Thumb Style="{StaticResource ScrollThumb}"/>
                                </Track.Thumb>
                            </Track>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="24">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <StackPanel Grid.Row="0" Margin="0,0,0,8" Orientation="Horizontal">
            <Border Width="4" Height="32" Background="#8B70F0"
                    CornerRadius="2" Margin="0,0,14,0" VerticalAlignment="Center"/>
            <StackPanel>
                <TextBlock Text="Invoke-Patch Theme Gallery" FontSize="22"
                           FontWeight="Bold" Foreground="#E8EAF0"/>
                <TextBlock Name="lblSubtitle" FontSize="12" Foreground="#7A7F9A"
                           Margin="0,2,0,0"
                           Text="Click a card to set it as your theme. Invoke-PatchGUI will relaunch."/>
            </StackPanel>
        </StackPanel>

        <Separator Grid.Row="1" Background="#2A2D45" Margin="0,8,0,16"/>

        <!-- Scrollable card grid -->
        <ScrollViewer Grid.Row="2" VerticalScrollBarVisibility="Auto"
                      HorizontalScrollBarVisibility="Disabled">
            <WrapPanel Name="pnlCards" Orientation="Horizontal"/>
        </ScrollViewer>
    </Grid>
</Window>
"@


$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)
$pnlCards = $window.FindName('pnlCards')


# ============================================================================
#  Build cards
# ============================================================================

function New-ThemeCard {
    param(
        [string]$Key,
        [hashtable]$Theme
    )

    # Primary accent color for swatches / hover / mini preview bar.
    # Most themes define Green as the "action" accent; cards surface it
    # so browsing is visually distinct across the grid.
    $accentHex = $Theme.Green

    $card = New-Object System.Windows.Controls.Border
    $card.Width         = 280
    $card.Height        = 180
    $card.Margin        = [System.Windows.Thickness]::new(8)
    $card.CornerRadius  = [System.Windows.CornerRadius]::new(10)
    $card.BorderThickness = [System.Windows.Thickness]::new(1)
    $card.Cursor        = [System.Windows.Input.Cursors]::Hand
    $card.Background    = [Windows.Media.BrushConverter]::new().ConvertFrom($Theme.Surface)
    $card.BorderBrush   = [Windows.Media.BrushConverter]::new().ConvertFrom($Theme.Border)
    $card.Padding       = [System.Windows.Thickness]::new(16)
    $card.Tag           = [pscustomobject]@{ Key = $Key; Theme = $Theme; AccentHex = $accentHex }

    # Outer grid
    $grid = New-Object System.Windows.Controls.Grid
    $row1 = New-Object System.Windows.Controls.RowDefinition
    $row1.Height = 'Auto'
    $row2 = New-Object System.Windows.Controls.RowDefinition
    $row2.Height = '*'
    $row3 = New-Object System.Windows.Controls.RowDefinition
    $row3.Height = 'Auto'
    $grid.RowDefinitions.Add($row1) > $null
    $grid.RowDefinitions.Add($row2) > $null
    $grid.RowDefinitions.Add($row3) > $null

    # Top row: accent bar + title/vibe
    $topPanel = New-Object System.Windows.Controls.StackPanel
    $topPanel.Orientation = 'Horizontal'
    $topPanel.VerticalAlignment = 'Top'
    [System.Windows.Controls.Grid]::SetRow($topPanel, 0)

    $bar = New-Object System.Windows.Controls.Border
    $bar.Width = 3
    $bar.Height = 32
    $bar.CornerRadius = [System.Windows.CornerRadius]::new(2)
    $bar.Background = [Windows.Media.BrushConverter]::new().ConvertFrom($accentHex)
    $bar.Margin = [System.Windows.Thickness]::new(0,2,10,0)
    $topPanel.Children.Add($bar) > $null

    $titlePanel = New-Object System.Windows.Controls.StackPanel

    $name = New-Object System.Windows.Controls.TextBlock
    $name.Text = $Theme.Name
    $name.FontSize = 15
    $name.FontWeight = 'SemiBold'
    $name.Foreground = [Windows.Media.BrushConverter]::new().ConvertFrom($Theme.Text)
    $titlePanel.Children.Add($name) > $null

    $vibe = New-Object System.Windows.Controls.TextBlock
    $vibe.Text = $Theme.Vibe
    $vibe.FontSize = 10
    $vibe.Foreground = [Windows.Media.BrushConverter]::new().ConvertFrom($Theme.SubText)
    $vibe.Margin = [System.Windows.Thickness]::new(0,2,0,0)
    $vibe.TextWrapping = 'Wrap'
    $titlePanel.Children.Add($vibe) > $null

    $topPanel.Children.Add($titlePanel) > $null
    $grid.Children.Add($topPanel) > $null

    # Middle: mini header preview (pipe + title), rendered on the theme's Bg
    $preview = New-Object System.Windows.Controls.Border
    $preview.Background = [Windows.Media.BrushConverter]::new().ConvertFrom($Theme.Bg)
    $preview.CornerRadius = [System.Windows.CornerRadius]::new(6)
    $preview.BorderBrush = [Windows.Media.BrushConverter]::new().ConvertFrom($Theme.Border)
    $preview.BorderThickness = [System.Windows.Thickness]::new(1)
    $preview.Padding = [System.Windows.Thickness]::new(10,8,10,8)
    $preview.VerticalAlignment = 'Center'
    $preview.Margin = [System.Windows.Thickness]::new(0,12,0,0)
    [System.Windows.Controls.Grid]::SetRow($preview, 1)

    $previewInner = New-Object System.Windows.Controls.StackPanel
    $previewInner.Orientation = 'Horizontal'

    $pbar = New-Object System.Windows.Controls.Border
    $pbar.Width = 2
    $pbar.Height = 18
    $pbar.CornerRadius = [System.Windows.CornerRadius]::new(1)
    $pbar.Background = [Windows.Media.BrushConverter]::new().ConvertFrom($Theme.Blue)
    $pbar.Margin = [System.Windows.Thickness]::new(0,0,8,0)
    $pbar.VerticalAlignment = 'Center'
    $previewInner.Children.Add($pbar) > $null

    $ptitle = New-Object System.Windows.Controls.TextBlock
    $ptitle.Text = 'Invoke-Patch'
    $ptitle.FontWeight = 'Bold'
    $ptitle.FontSize = 13
    $ptitle.Foreground = [Windows.Media.BrushConverter]::new().ConvertFrom($Theme.Blue)
    $ptitle.VerticalAlignment = 'Center'
    $previewInner.Children.Add($ptitle) > $null

    $preview.Child = $previewInner
    $grid.Children.Add($preview) > $null

    # Bottom: color swatches -- Bg / Surface / Border / Text / Blue / Green / Red
    $swatchPanel = New-Object System.Windows.Controls.StackPanel
    $swatchPanel.Orientation = 'Horizontal'
    $swatchPanel.Margin = [System.Windows.Thickness]::new(0,12,0,0)
    [System.Windows.Controls.Grid]::SetRow($swatchPanel, 2)

    $swatchKeys = @('Bg','Surface','Border','Text','Blue','Green','Red')
    foreach ($k in $swatchKeys) {
        $sw = New-Object System.Windows.Controls.Border
        $sw.Width = 18
        $sw.Height = 18
        $sw.CornerRadius = [System.Windows.CornerRadius]::new(3)
        $sw.Background = [Windows.Media.BrushConverter]::new().ConvertFrom($Theme[$k])
        $sw.BorderBrush = [Windows.Media.BrushConverter]::new().ConvertFrom('#3A3D55')
        $sw.BorderThickness = [System.Windows.Thickness]::new(1)
        $sw.Margin = [System.Windows.Thickness]::new(0,0,4,0)
        $sw.ToolTip = "{0}: {1}" -f $k, $Theme[$k]
        $swatchPanel.Children.Add($sw) > $null
    }
    $grid.Children.Add($swatchPanel) > $null

    $card.Child = $grid

    # Click -> save preference + launch Invoke-PatchGUI.ps1 + close Gallery.
    $card.Add_MouseLeftButtonUp({
        $tag = $this.Tag
        $prefs = Get-PatchPreferences
        $prefs.theme = $tag.Key
        Set-PatchPreferences -Preferences $prefs

        if (Test-Path $script:PatchGUI) {
            Start-Process -FilePath 'powershell.exe' `
                          -ArgumentList '-NoProfile', '-File', $script:PatchGUI
        }
        $window = [System.Windows.Window]::GetWindow($this)
        if ($window) { $window.Close() }
    })

    # Hover effect -- border thickens + switches to the theme's accent color
    $card.Add_MouseEnter({
        $this.BorderBrush = [Windows.Media.BrushConverter]::new().ConvertFrom($this.Tag.AccentHex)
        $this.BorderThickness = [System.Windows.Thickness]::new(2)
    })
    $card.Add_MouseLeave({
        $this.BorderBrush = [Windows.Media.BrushConverter]::new().ConvertFrom($this.Tag.Theme.Border)
        $this.BorderThickness = [System.Windows.Thickness]::new(1)
    })

    return $card
}


# Build cards in DisplayOrder; any theme keys not in the order list are
# appended alphabetically at the end (safety net for new themes).
$orderedKeys = @($script:DisplayOrder)
$remaining = @($script:AllThemes.Keys | Where-Object { $_ -notin $orderedKeys } | Sort-Object)
$allKeys = $orderedKeys + $remaining

foreach ($key in $allKeys) {
    if (-not $script:AllThemes.ContainsKey($key)) { continue }
    $theme = $script:AllThemes[$key]
    $card = New-ThemeCard -Key $key -Theme $theme
    $pnlCards.Children.Add($card) > $null
}


# ============================================================================
#  Show window
# ============================================================================

$window.ShowDialog() > $null
