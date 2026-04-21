# DOTS formatting comment

<#
    .SYNOPSIS
        Manages the global $TargetMachines list for remediation scripts.
    .DESCRIPTION
        Provides Import-TargetMachines (from file) and Enter-TargetMachines
        (from pasted input) to populate $Global:TargetMachines.  Both functions
        pipe through Format-ComputerList -ToUpper for consistent sanitization.

        Written by Skyler Werner
        Date: 2026/04/13
        Version 1.0.0
#>

$TargetMachinesFile = "$env:USERPROFILE\Desktop\Lists\Target_Machines.txt"

# -------------------------------------------------------------------
#  Import-TargetMachines  -  read from the standard text file on disk
# -------------------------------------------------------------------
function Import-TargetMachines {
    [CmdletBinding()]
    param([switch]$Silent)

    if (Test-Path $TargetMachinesFile) {
        $Global:TargetMachines = @(Get-Content $TargetMachinesFile |
            Format-ComputerList -ToUpper)

        if (-not $Silent) {
            Write-Host "Loaded $($Global:TargetMachines.Count) target machines." -ForegroundColor Green
        }
    }
    else {
        if (-not $Silent) {
            Write-Warning "Target machines file not found: $TargetMachinesFile"
        }
        $Global:TargetMachines = @()
    }
}

# -------------------------------------------------------------------
#  Enter-TargetMachines  -  paste / type a list into the console
# -------------------------------------------------------------------
function Enter-TargetMachines {
    [CmdletBinding()]
    param([switch]$Append)

    Write-Host "Paste or type machine names (one per line). Press Enter on an empty line to finish." -ForegroundColor Cyan

    $lines = [System.Collections.ArrayList]::new()
    while ($true) {
        $entry = Read-Host
        if ([string]::IsNullOrWhiteSpace($entry)) { break }
        # Some console hosts deliver a multi-line paste as one Read-Host
        # response.  Split on line breaks so each machine is a separate entry.
        foreach ($segment in ($entry -split '\r?\n')) {
            if (-not [string]::IsNullOrWhiteSpace($segment)) {
                $null = $lines.Add($segment)
            }
        }
    }

    if ($lines.Count -eq 0) {
        Write-Warning "No input received."
        return
    }

    $parsed = @($lines | Format-ComputerList -ToUpper)

    if ($Append -and $Global:TargetMachines.Count -gt 0) {
        $Global:TargetMachines = @(
            @($Global:TargetMachines) + $parsed | Sort-Object -Unique
        )
    }
    else {
        $Global:TargetMachines = $parsed
    }

    Write-Host "Parsed $($parsed.Count) machines. Total in list: $($Global:TargetMachines.Count)" -ForegroundColor Green
    $Global:TargetMachines
}

# --- Alias for backward compatibility ---
Set-Alias -Name 'Load-TargetMachines' -Value 'Import-TargetMachines'

# --- Auto-load on module import ---
Import-TargetMachines -Silent

Export-ModuleMember -Function Import-TargetMachines, Enter-TargetMachines `
                    -Alias Load-TargetMachines
