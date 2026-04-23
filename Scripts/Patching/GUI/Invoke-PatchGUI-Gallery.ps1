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
        Version 2.1.0
    .PARAMETER DryRun
        Forwarded from the main GUI when the theme picker is opened from
        an in-flight DryRun session. The flag is passed back to
        Invoke-PatchGUI.ps1 on theme-card click so DryRun mode survives
        the relaunch. Not intended for direct CLI use.
    .PARAMETER Mode
        Forwarded from the main GUI's live toggle state (Patch or
        Version) at the moment the theme picker was opened. Passed back
        to Invoke-PatchGUI.ps1 on theme-card click so the user returns
        to the mode they were in. Not intended for direct CLI use.
    .EXAMPLE
        .\Invoke-PatchGUI-Gallery.ps1
#>

param(
    [switch]$DryRun,
    [ValidateSet('Patch','Version')]
    [string]$Mode
)

# Capture the session flags into script scope so the card-click closure
# can read them when spawning the new Invoke-PatchGUI process.
$script:SessionDryRun = [bool]$DryRun.IsPresent
$script:SessionMode   = $Mode

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase


# ============================================================================
#  Load themes + shared helpers
# ============================================================================
# Invoke-PatchGUI.Shared.ps1 defines Get-PatchPreferences /
# Set-PatchPreferences and sets $script:PrefsDir / $script:PrefsPath
# into our scope.

. (Join-Path $PSScriptRoot 'Invoke-PatchGUI.Shared.ps1')

$script:ThemesPath = Join-Path $PSScriptRoot 'Themes.psd1'
$script:AllThemes  = Import-PowerShellDataFile -Path $script:ThemesPath
$script:PatchGUI   = Join-Path $PSScriptRoot 'Invoke-PatchGUI.ps1'

# Display order -- arranged for visual balance in a 4-column grid:
#   - CobaltSlate + Day paired at top-left (identity pairing)
#   - Light themes placed diagonally (R1C2 and R2C4) so they don't
#     form a vertical stripe
#   - Warm/cool accents alternate horizontally within each row
#   - Teal/cyan adjacency broken in row 3 (red Tanium sits between
#     Carbon Teal and Cyber Command)
#   - CyberCommand anchors the bottom-right corner -- its
#     high-contrast neon-on-black is the loudest theme in the set
#     and works as a visual bookend opposite Cobalt Slate.
$script:DisplayOrder = @(
    'CobaltSlate',     'CobaltSlateDay', 'DarkForest',     'TokyoNight',
    'UltraDarkViolet', 'NavyAnalytics',  'Meridian',       'Quartz',
    'Monochrome',      'CarbonTeal',     'TaniumInspired', 'CyberCommand'
)

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
    MaxHeight="900"
    SizeToContent="Height"
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
            <WrapPanel Name="pnlCards" Orientation="Horizontal"
                       HorizontalAlignment="Center"/>
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

# Build a Brush (solid or linear gradient) from a theme's token set.
# Graceful cascade: use the *Stops array if present, else the *Color
# solid if present, else the named fallback token (usually 'Blue').
# Surfacing PipeStops / TitleStops / PipeColor on the gallery card
# lets themes like Tanium-Inspired (red pipe) and Tokyo Night (cyan
# to pink gradient) preview accurately instead of flattening to Blue.
function New-ThemeBrush {
    param(
        [hashtable]$Theme,
        [string]$StopsKey,
        [string]$SolidKey,
        [string]$FallbackKey,
        [ValidateSet('Vertical','Horizontal')]
        [string]$Orientation = 'Vertical'
    )

    if ($StopsKey -and $Theme.ContainsKey($StopsKey) -and $Theme[$StopsKey]) {
        $brush = New-Object System.Windows.Media.LinearGradientBrush
        if ($Orientation -eq 'Horizontal') {
            $brush.StartPoint = [Windows.Point]::new(0, 0.5)
            $brush.EndPoint   = [Windows.Point]::new(1, 0.5)
        } else {
            $brush.StartPoint = [Windows.Point]::new(0.5, 0)
            $brush.EndPoint   = [Windows.Point]::new(0.5, 1)
        }
        $stops = @($Theme[$StopsKey])
        for ($i = 0; $i -lt $stops.Count; $i++) {
            $offset = if ($stops.Count -eq 1) { 0 } else { $i / ($stops.Count - 1) }
            $color  = [Windows.Media.ColorConverter]::ConvertFromString($stops[$i])
            $stop   = New-Object System.Windows.Media.GradientStop -ArgumentList $color, $offset
            $brush.GradientStops.Add($stop) > $null
        }
        return $brush
    }
    if ($SolidKey -and $Theme.ContainsKey($SolidKey) -and $Theme[$SolidKey]) {
        return [Windows.Media.BrushConverter]::new().ConvertFrom($Theme[$SolidKey])
    }
    return [Windows.Media.BrushConverter]::new().ConvertFrom($Theme[$FallbackKey])
}

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
    # Top accent bar surfaces the theme's brand pipe: PipeStops for
    # themes with a vertical gradient (Tokyo Night), PipeColor for
    # themes with a distinct brand color (Tanium red, Dark Forest
    # copper), else the Green action accent. Vertical orientation
    # matches the schema and the bar's own vertical geometry.
    $bar.Background = New-ThemeBrush -Theme $Theme -StopsKey 'PipeStops' `
                                     -SolidKey 'PipeColor' -FallbackKey 'Green' `
                                     -Orientation 'Vertical'
    $bar.Margin = [System.Windows.Thickness]::new(0,2,10,0)
    $topPanel.Children.Add($bar) > $null

    $titlePanel = New-Object System.Windows.Controls.StackPanel

    $name = New-Object System.Windows.Controls.TextBlock
    $name.Text = $Theme.Name
    $name.FontSize = 16
    $name.FontWeight = 'SemiBold'
    $name.Foreground = [Windows.Media.BrushConverter]::new().ConvertFrom($Theme.Text)
    $titlePanel.Children.Add($name) > $null

    $vibe = New-Object System.Windows.Controls.TextBlock
    $vibe.Text = $Theme.Vibe
    $vibe.FontSize = 11
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

    # Title rendering honors TitleStyle so the mini preview matches
    # what the theme will actually produce when selected:
    #   'split'  -> "Invoke" dim + "Patch" bold, no hyphen
    #              (Monochrome, Carbon Teal)
    #   default  -> "Invoke-Patch" single TextBlock; uses TitleStops
    #              vertical gradient if present (Tokyo Night), else
    #              solid Theme.Blue (everyone else)
    $titleStyle = if ($Theme.ContainsKey('TitleStyle')) { $Theme.TitleStyle } else { 'solid' }

    if ($titleStyle -eq 'split') {
        $dimHex = if ($Theme.ContainsKey('TitleDimColor')) { $Theme.TitleDimColor } else { $Theme.SubText }

        $ptitleDim = New-Object System.Windows.Controls.TextBlock
        $ptitleDim.Text = 'Invoke'
        $ptitleDim.FontWeight = 'Normal'
        $ptitleDim.FontSize = 13
        $ptitleDim.Foreground = [Windows.Media.BrushConverter]::new().ConvertFrom($dimHex)
        $ptitleDim.VerticalAlignment = 'Center'
        $previewInner.Children.Add($ptitleDim) > $null

        $ptitle = New-Object System.Windows.Controls.TextBlock
        $ptitle.Text = 'Patch'
        $ptitle.FontWeight = 'Bold'
        $ptitle.FontSize = 13
        $ptitle.Foreground = [Windows.Media.BrushConverter]::new().ConvertFrom($Theme.Blue)
        $ptitle.VerticalAlignment = 'Center'
        $previewInner.Children.Add($ptitle) > $null
    }
    else {
        $ptitle = New-Object System.Windows.Controls.TextBlock
        $ptitle.Text = 'Invoke-Patch'
        $ptitle.FontWeight = 'Bold'
        $ptitle.FontSize = 13
        $ptitle.Foreground = New-ThemeBrush -Theme $Theme -StopsKey 'TitleStops' `
                                            -FallbackKey 'Blue' `
                                            -Orientation 'Vertical'
        $ptitle.VerticalAlignment = 'Center'
        $previewInner.Children.Add($ptitle) > $null
    }

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
            # Forward session flags so DryRun and the live Patch/Version
            # toggle state both survive the theme-change relaunch.
            $argList = @('-NoProfile', '-File', $script:PatchGUI)
            if ($script:SessionDryRun) { $argList += '-DryRun' }
            if ($script:SessionMode)   { $argList += @('-Mode', $script:SessionMode) }
            Start-Process -FilePath 'powershell.exe' -ArgumentList $argList
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
