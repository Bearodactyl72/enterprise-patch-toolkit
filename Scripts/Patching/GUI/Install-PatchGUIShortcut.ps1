# DOTS formatting comment

<#
    .SYNOPSIS
        Creates desktop + Start Menu shortcuts that launch Invoke-PatchGUI.
    .DESCRIPTION
        Run during Setup.ps1 so admins get the GUI as a clickable shortcut
        on their desktop and Start Menu. Can also be run standalone to
        re-create the shortcuts if they get deleted, or to retarget them
        after the repo moves to a new location.

        The shortcut launches:
          powershell.exe -NoProfile -ExecutionPolicy Bypass
                         -WindowStyle Hidden
                         -File "<repo>\Scripts\Patching\GUI\Invoke-PatchGUI.ps1"

        -NoProfile is used because the GUI only needs the profile inside
        its background runspace (which dot-sources it explicitly). This
        makes shortcut launches faster than CLI launches and avoids
        profile side effects on the short-lived host shell.

        -WindowStyle Hidden + shortcut WindowStyle=7 (Minimized) together
        suppress the PowerShell console window so only the WPF window
        appears. There's still a brief flicker on startup; eliminating
        it entirely would require a conhost/wscript wrapper and isn't
        worth the complexity.

        Written by Skyler Werner
        Date: 2026/04/21
        Version 1.0.0
    .PARAMETER RepoRoot
        Root of the Remediation-Script-Library repo. Defaults to the
        grandparent of this script (Scripts/Patching/GUI -> repo root).
    .PARAMETER Name
        Display name / file base used for both shortcut files. Default
        'Invoke-Patch' so the icon label matches the CLI command name
        admins already know.
    .PARAMETER SkipDesktop
        Skip creating the desktop shortcut.
    .PARAMETER SkipStartMenu
        Skip creating the Start Menu shortcut.
    .EXAMPLE
        .\Install-PatchGUIShortcut.ps1
    .EXAMPLE
        .\Install-PatchGUIShortcut.ps1 -Name 'Patching GUI' -SkipStartMenu
#>

[CmdletBinding()]
param(
    [string]$RepoRoot,
    [string]$Name = 'Invoke-Patch',
    [switch]$SkipDesktop,
    [switch]$SkipStartMenu
)


# ============================================================================
#  Resolve paths
# ============================================================================

if (-not $RepoRoot) {
    # Script lives at <repo>\Scripts\Patching\GUI\. Three Split-Parents
    # to climb back out to the repo root.
    $RepoRoot = Split-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -Parent
}

$guiScript = Join-Path $RepoRoot 'Scripts\Patching\GUI\Invoke-PatchGUI.ps1'
if (-not (Test-Path $guiScript)) {
    throw "Could not find GUI script at: $guiScript"
}

$powershellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
if (-not (Test-Path $powershellExe)) {
    throw "Could not find powershell.exe at: $powershellExe"
}


# ============================================================================
#  Shortcut factory
# ============================================================================

function New-PatchGUIShortcut {
    param(
        [Parameter(Mandatory)][string]$LinkPath,
        [Parameter(Mandatory)][string]$Target,
        [Parameter(Mandatory)][string]$Arguments,
        [string]$IconLocation,
        [string]$WorkingDirectory,
        [string]$Description
    )

    $parent = Split-Path $LinkPath -Parent
    if (-not (Test-Path $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }

    $wshShell = New-Object -ComObject WScript.Shell
    $shortcut = $wshShell.CreateShortcut($LinkPath)
    $shortcut.TargetPath       = $Target
    $shortcut.Arguments        = $Arguments
    $shortcut.WorkingDirectory = $WorkingDirectory
    $shortcut.Description      = $Description
    $shortcut.IconLocation     = $IconLocation

    # 7 = SW_SHOWMINNOACTIVE. Combined with powershell -WindowStyle Hidden
    # this keeps the PowerShell console out of the taskbar / Alt-Tab.
    $shortcut.WindowStyle = 7

    $shortcut.Save()

    # Release the COM reference so the file handle closes promptly
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shortcut) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wshShell) | Out-Null
}


# ============================================================================
#  Build args and create the shortcuts
# ============================================================================

$arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$guiScript`""
$workDir   = Split-Path $guiScript -Parent
$descr     = 'Launches the WPF patching and version-audit GUI (Invoke-Patch / Invoke-Version front-end).'
$icon      = "$powershellExe,0"

$linkName  = "$Name.lnk"
$created   = @()

if (-not $SkipDesktop) {
    $desktopLink = Join-Path ([Environment]::GetFolderPath('Desktop')) $linkName
    New-PatchGUIShortcut -LinkPath $desktopLink `
                         -Target $powershellExe `
                         -Arguments $arguments `
                         -IconLocation $icon `
                         -WorkingDirectory $workDir `
                         -Description $descr
    $created += $desktopLink
}

if (-not $SkipStartMenu) {
    $startMenuDir = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs'
    $startLink    = Join-Path $startMenuDir $linkName
    New-PatchGUIShortcut -LinkPath $startLink `
                         -Target $powershellExe `
                         -Arguments $arguments `
                         -IconLocation $icon `
                         -WorkingDirectory $workDir `
                         -Description $descr
    $created += $startLink
}


# ============================================================================
#  Report
# ============================================================================

if ($created.Count -gt 0) {
    Write-Host ''
    Write-Host 'Created shortcuts:' -ForegroundColor Cyan
    foreach ($path in $created) {
        Write-Host "  $path"
    }
    Write-Host ''
    Write-Host 'To pin to the taskbar: right-click the desktop shortcut and choose "Pin to taskbar".' -ForegroundColor Yellow
    Write-Host ''
}
