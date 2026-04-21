# DOTS formatting comment

<#
    .SYNOPSIS
        Formats a list of machines, removing domain prefixes, suffixes, and blank entries.
    .DESCRIPTION
        Formats a list of machines, removing domain prefixes, suffixes, and blank entries.
        Written by Skyler Werner
        Date: 2025/10/10
        Modified: 2025/11/28
        Version 1.0.1
#>

function Format-ComputerList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        $InputObject,
        [Parameter(Position = 1)]
        [Switch]
        $ToUpper
    )

    begin {
        $collected = [System.Collections.ArrayList]::new()
    }

    process {
        if ($InputObject -is [array]) {
            foreach ($item in $InputObject) { $null = $collected.Add($item) }
        }
        else {
            $null = $collected.Add($InputObject)
        }
    }

    end {
        $list = @($collected)

        # Strip NetBIOS-prefix, FQDN-prefix, and FQDN-suffix for the active
        # network profile. No-op on a non-domain / unmatched host, or when
        # the RSL-Environment loader module is not available.
        $activeNetwork = if (Get-Command Get-RSLActiveNetwork -ErrorAction SilentlyContinue) {
            Get-RSLActiveNetwork
        } else { $null }

        if ($activeNetwork) {
            $prefix1 = $activeNetwork.DomainShort + "\"
            $prefix2 = $activeNetwork.DomainFqdn  + "\"
            $suffix  = "." + $activeNetwork.DomainFqdn
            $list = ($list).Replace($prefix1,"")
            $list = ($list).Replace($prefix2,"")
            $list = ($list).Replace($suffix,"")
        }

        # Removes duplicates from list
        $listTrimUniq = $list.Trim() | Sort-Object -Unique
        $listClean =  $listTrimUniq | Where-Object { $_ -ne "" }

        if ($ToUpper) {
            $listClean.ToUpper()
        }
        else {
            $listClean
        }
    }
}

Export-ModuleMember Format-ComputerList
