# DOTS formatting comment

function Test-ConnectionAsJob {
<#
.SYNOPSIS
    Tests connectivity to one or more computers in parallel using background jobs.

.DESCRIPTION
    Pings each computer concurrently via Test-Connection -AsJob and returns a
    uniform object with ComputerName, IPV4Address, and Reachable properties.

    For lists exceeding 1000 targets a warning is displayed and execution is
    delayed 15 seconds to allow cancellation before the network burst.

.PARAMETER ComputerName
    One or more hostnames or IP addresses to test.

.INPUTS
    System.String[]

.OUTPUTS
    PSCustomObject with properties: ComputerName, IPV4Address, Reachable

.EXAMPLE
    Test-ConnectionAsJob -ComputerName 'SERVER01', 'SERVER02'

.EXAMPLE
    Get-Content .\machines.txt | Test-ConnectionAsJob

.NOTES
    Author:   Skyler Werner
    Created:  2021-05-05
    Modified: 2026-03-16
    Version:  1.1.0
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $ComputerName
    )

    begin {
        $allComputers = [System.Collections.Generic.List[string]]::new()
    }

    process {
        foreach ($name in $ComputerName) {
            $allComputers.Add($name)
        }
    }

    end {
        if ($allComputers.Count -gt 1000) {
            Write-Warning (
                "You are trying to connect to more than 1000 machines. " +
                "This may cause a large spike in network utilization. " +
                "You have 10 seconds to cancel before it continues automatically..."
            )
            Start-Sleep -Seconds 10
        }

        $allComputers |
            ForEach-Object { Test-Connection -ComputerName $_ -Count 1 -AsJob } |
            Get-Job |
            Receive-Job -Wait |
            Select-Object @{
                Name       = 'ComputerName'
                Expression = { $_.Address }
            },
            @{
                Name       = 'IPV4Address'
                Expression = {
                    if ($null -ne $_.IPV4Address) {
                        $_.IPV4Address
                    }
                    elseif ($null -ne $_.ProtocolAddress) {
                        $_.ProtocolAddress
                    }
                    elseif ($_.Address -match '\.') {
                        $_.Address
                    }
                    else {
                        $null
                    }
                }
            },
            @{
                Name       = 'Reachable'
                Expression = { $_.StatusCode -eq 0 }
            } |
            Tee-Object -Variable pingResults

        Get-Job | Remove-Job -Force
    }
}

Export-ModuleMember -Function Test-ConnectionAsJob
