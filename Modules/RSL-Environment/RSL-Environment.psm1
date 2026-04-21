# DOTS formatting comment

<#
    .SYNOPSIS
        Loads environment-specific configuration for the Remediation Script
        Library and resolves the active network profile.
    .DESCRIPTION
        Exposes Import-RSLEnvironment and Get-RSLActiveNetwork.

        Import-RSLEnvironment reads Config\Environment.psd1 if present, else
        falls back to Config\Environment.example.psd1 so the repo is still
        runnable on a fresh clone. The loaded hashtable is cached in a
        script-scope variable; subsequent calls reuse it unless -Force is
        passed.

        Get-RSLActiveNetwork matches $env:USERDNSDOMAIN against each
        Networks[*].DomainFqdn entry (case-insensitive) and returns the
        winning entry. If nothing matches, it returns $null and the caller
        should treat the host as a workgroup / non-domain machine.

        Written by Skyler Werner
        Date: 2026/04/22
        Version 1.0.0
#>

$script:RSLConfig = $null

function Import-RSLEnvironment {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [string]
        $RepoRoot,

        [switch]
        $Force
    )

    if ($script:RSLConfig -and -not $Force) {
        return $script:RSLConfig
    }

    if (-not $RepoRoot) {
        $RepoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    }

    $configDir  = Join-Path $RepoRoot 'Config'
    $realPath   = Join-Path $configDir 'Environment.psd1'
    $examplePath = Join-Path $configDir 'Environment.example.psd1'

    $chosenPath = $null
    if (Test-Path $realPath) {
        $chosenPath = $realPath
    }
    elseif (Test-Path $examplePath) {
        $chosenPath = $examplePath
        Write-Warning "Environment.psd1 not found; using Environment.example.psd1 (placeholder values)."
    }
    else {
        throw "No environment config found. Expected $realPath or $examplePath."
    }

    $script:RSLConfig = Import-PowerShellDataFile -Path $chosenPath
    $script:RSLConfig
}

function Get-RSLActiveNetwork {
    [CmdletBinding()]
    param()

    if (-not $script:RSLConfig) {
        [void](Import-RSLEnvironment)
    }

    $domain = $env:USERDNSDOMAIN
    if (-not $domain) {
        return $null
    }

    foreach ($net in $script:RSLConfig.Networks) {
        if ($net.DomainFqdn -and ($domain -ieq $net.DomainFqdn)) {
            return $net
        }
    }

    $null
}

Export-ModuleMember -Function Import-RSLEnvironment, Get-RSLActiveNetwork
