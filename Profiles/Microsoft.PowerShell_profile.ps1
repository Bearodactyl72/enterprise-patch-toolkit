# DOTS formatting comment

# About PowerShell Profiles:
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_profiles?view=powershell-7.1

Set-Location C:\

$ScriptPaths = @()
$ModulePaths = @()
if (Test-Path "$env:APPDATA\Patching" -PathType Container) {
    $String = Get-Content "$env:APPDATA\Patching\Paths.txt"
    foreach ($Line in $String) {
        if ($Line -match "ScriptPath :") {
            $ScriptPaths += $Line.Replace("ScriptPath : ","")
        }
        if ($Line -match "ModulePath :") {
            $ModulePaths += $Line.Replace("ModulePath : ","")
        }
    }

    if ($ModulePaths.Count -gt 0) {
        $ModuleFolders = Get-ChildItem $ModulePaths -Directory

        foreach ($Folder in $ModuleFolders) {
            Import-Module $Folder.FullName
        }
    }

    if ($ScriptPaths.Count -eq 0) {
        Write-Warning "No script paths found in Paths.txt"
    }

    foreach ($ScriptPath in $ScriptPaths) {
        . "$ScriptPath\Patching\Invoke-Patch.ps1"
        . "$ScriptPath\Patching\Invoke-Version.ps1"

        # --- Test runner ---
        $repoRoot = Split-Path $ScriptPath -Parent
        $testsPath = Join-Path $repoRoot 'Tests'
        if (Test-Path "$testsPath\Invoke-AllTests.ps1") {
            . "$testsPath\Invoke-AllTests.ps1"
        }

        # --- Discoverability ---
        $repoRootForHelp = Split-Path $ScriptPath -Parent
        Set-Variable -Name '_RepoRoot' -Value $repoRootForHelp -Scope Script

        function Get-PatchingCommands {
            <#
                .SYNOPSIS
                    Lists all available patching commands, modules, and scripts.
            #>
            $divider = ('-' * 60)

            Write-Host ''
            Write-Host 'MODULES' -ForegroundColor Cyan
            Write-Host $divider
            $loadedModules = @(Get-Module |
                Where-Object { $_.Path -like "$script:_RepoRoot*" })
            if ($loadedModules.Count -gt 0) {
                foreach ($m in $loadedModules) {
                    $cmds = @($m.ExportedCommands.Keys)
                    Write-Host "  $($m.Name)" -ForegroundColor Yellow -NoNewline
                    Write-Host "  ->  $($cmds -join ', ')"
                }
            }
            else {
                Write-Host '  (none loaded)'
            }

            Write-Host ''
            Write-Host 'DOT-SOURCED COMMANDS' -ForegroundColor Cyan
            Write-Host $divider
            Write-Host '  Invoke-Patch'
            Write-Host '  Invoke-Version'
            Write-Host '  Invoke-AllTests  (if Tests\ present)'

            Write-Host ''
            Write-Host 'GUI' -ForegroundColor Cyan
            Write-Host $divider
            Write-Host '  Invoke-PatchGUI [-Theme <key>] [-Mode Patch|Version] [-DryRun]'

            Write-Host ''
            Write-Host 'IMPORT-EXPORT WRAPPERS' -ForegroundColor Cyan
            Write-Host $divider
            Write-Host '  Export-PackageFlat [-ReferenceExport <path>]'
            Write-Host '  Import-PackageFlat [-Undo] [-WhatIf]'
            Write-Host '  Export-Package'
            Write-Host '  Import-Package'

            $utilPath = Join-Path $script:_RepoRoot 'Scripts\Utility'
            if (Test-Path $utilPath) {
                Write-Host ''
                Write-Host 'UTILITY SCRIPTS  (run manually)' -ForegroundColor Cyan
                Write-Host $divider
                $categories = Get-ChildItem $utilPath -Directory
                foreach ($cat in $categories) {
                    Write-Host "  [$($cat.Name)]" -ForegroundColor Yellow
                    $scripts = @(Get-ChildItem $cat.FullName -Filter '*.ps1')
                    foreach ($s in $scripts) {
                        Write-Host "    $($s.BaseName)"
                    }
                }
            }

            $uninstallPath = Join-Path $script:_RepoRoot 'Scripts\Patching\AppUninstalls'
            if (Test-Path $uninstallPath) {
                Write-Host ''
                Write-Host 'APP UNINSTALL SCRIPTS' -ForegroundColor Cyan
                Write-Host $divider
                $scripts = @(Get-ChildItem $uninstallPath -Filter '*.ps1' |
                    Where-Object { $_.Name -ne '_Template.ps1' })
                foreach ($s in $scripts) {
                    Write-Host "    $($s.BaseName)"
                }
            }

            Write-Host ''
        }

        # --- Patching GUI wrapper ---
        $GuiPath = Join-Path $ScriptPath 'Patching\GUI'
        if (Test-Path $GuiPath) {
            Set-Variable -Name '_GuiPath' -Value $GuiPath -Scope Script

            function Invoke-PatchGUI {
                <#
                    .SYNOPSIS
                        Launches the WPF patching GUI (Invoke-Patch / Invoke-Version front-end).
                    .DESCRIPTION
                        Opens the themed WPF GUI for running patching or version
                        audits without touching the command line. Mode slider in
                        the header swaps between Invoke-Patch and Invoke-Version
                        at runtime.
                    .PARAMETER Theme
                        One-off theme key override (e.g. 'TokyoNight'). See the
                        Themes.psd1 alongside the GUI for available keys, or run
                        Invoke-PatchGUI-Gallery to pick visually.
                    .PARAMETER Mode
                        Starting mode: 'Patch' (default) or 'Version'.
                    .PARAMETER DryRun
                        Populates the GUI with mock data for UI exercise without
                        running real patching or version queries.
                    .EXAMPLE
                        Invoke-PatchGUI
                    .EXAMPLE
                        Invoke-PatchGUI -Mode Version -Theme TokyoNight
                    .EXAMPLE
                        Invoke-PatchGUI -DryRun
                #>
                [CmdletBinding()]
                param(
                    [string]$Theme,
                    [ValidateSet('Patch','Version')]
                    [string]$Mode = 'Patch',
                    [switch]$DryRun
                )
                $params = @{}
                if ($Theme)  { $params['Theme']  = $Theme }
                if ($Mode)   { $params['Mode']   = $Mode }
                if ($DryRun) { $params['DryRun'] = $true }
                & "$script:_GuiPath\Invoke-PatchGUI.ps1" @params
            }
        }

        # --- Import-Export wrapper functions ---
        $ImportExportPath = Join-Path (Split-Path $ScriptPath -Parent) 'Import-Export'
        if (Test-Path $ImportExportPath) {
            # Store path for use inside the script blocks
            Set-Variable -Name '_ImportExportPath' -Value $ImportExportPath -Scope Script

            function Export-PackageFlat {
                param(
                    [Alias('Diff', 'Since')]
                    [string]$ReferenceExport
                )
                $params = @{}
                if ($ReferenceExport) { $params['ReferenceExport'] = $ReferenceExport }
                & "$script:_ImportExportPath\Export-Package-Flat.ps1" @params
            }

            function Import-PackageFlat {
                [CmdletBinding(SupportsShouldProcess)]
                param(
                    [switch]$Undo
                )
                $params = @{}
                if ($Undo)                { $params['Undo']    = $true }
                if ($WhatIfPreference)    { $params['WhatIf']  = $true }
                & "$script:_ImportExportPath\Import-Package-Flat.ps1" @params
            }

            function Export-Package {
                & "$script:_ImportExportPath\Export-Package.ps1"
            }

            function Import-Package {
                & "$script:_ImportExportPath\Import-Package.ps1"
            }
        }
    }
}
else {
    Write-Host "Modules have not been imported."
    Write-Host "Scripts have not been initialized."
}

Remove-Variable String, Line, ModuleFolders, Folder, ScriptPaths, ModulePaths,
    repoRoot, testsPath, ImportExportPath, GuiPath, repoRootForHelp -ErrorAction SilentlyContinue

function Get-Thotfix {Write-Host '*bonk* go to horny jail'}

function Invoke-PowerShell {
    powershell -nologo
    Invoke-PowerShell
}

function Restart-PowerShell {
    if ($host.Name -eq 'ConsoleHost') {
        exit
    }
    Write-Warning 'Only usable while in the PowerShell console host'
}

Set-Alias -Name 'reload' -Value 'Restart-PowerShell'

$parentProcessId = (Get-WmiObject Win32_Process -Filter "ProcessId = $PID").ParentProcessId
$parentProcessName = (Get-WmiObject Win32_Process -Filter "ProcessId = $parentProcessId").ProcessName

if ($host.Name -eq 'ConsoleHost') {
    if (-not($parentProcessName -eq 'powershell.exe')) {
        # Invoke-PowerShell
    }
}


# --- Session variables ---
$myPC = $env:CLIENTNAME
