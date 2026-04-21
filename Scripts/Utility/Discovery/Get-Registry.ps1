# DOTS formatting comment

function Get-Registry {
    <#
        .SYNOPSIS
            Queries registry uninstall keys on remote machines for a given software name.
        .DESCRIPTION
            Thin wrapper around Get-RegistryKey that accepts a computer list and
            software name, then displays results sorted by key presence and version.
        .PARAMETER ComputerName
            One or more computer names to query. Accepts pipeline input.
        .PARAMETER SoftwareName
            Display name pattern to match in the registry uninstall keys.
        .EXAMPLE
            $list = Get-Content "$env:USERPROFILE\Desktop\Lists\Target_Machines.txt"
            Get-Registry -ComputerName $list -SoftwareName "Google Chrome"
        .EXAMPLE
            Get-Registry -ComputerName "PC01","PC02" -SoftwareName "Adobe Acrobat"
        .NOTES
            Written by Skyler Werner
            Version: 2.0.0
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [string[]]
        $ComputerName,

        [Parameter(Mandatory, Position = 1)]
        [string]
        $SoftwareName,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        $collectedNames = @()
    }

    process {
        foreach ($name in $ComputerName) {
            if ($name.Length -gt 0) {
                $collectedNames += $name
            }
        }
    }

    end {
        $targets = @(Format-ComputerList $collectedNames -ToUpper)
        if ($targets.Count -eq 0) {
            Write-Warning "No valid computer names provided."
            return
        }

        Write-Host "Checking $SoftwareName registry keys on $($targets.Count) computer(s)..."

        $registryResults = Get-RegistryKey -ComputerName $targets -SoftwareName $SoftwareName

        $sorted = $registryResults |
            Select-Object ComputerName, KeyPresent, DisplayName, Version |
            Sort-Object KeyPresent, DisplayName, Version

        $sorted | Format-Table -AutoSize | Out-Host
        if ($PassThru) { return $sorted }
    }
}
