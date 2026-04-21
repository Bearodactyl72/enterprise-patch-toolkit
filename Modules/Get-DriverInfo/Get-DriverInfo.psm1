# DOTS formatting comment

<#
    .SYNOPSIS
        Retrieves driver information from remote machines using runspaces.
    .DESCRIPTION
        Queries installed driver information on one or more remote machines in parallel.
        Supports two search modes: .inf drivers (via Get-WindowsDriver) and loaded
        drivers (via driverquery). Uses Invoke-RunspacePool for concurrent execution.

        Written by Skyler Werner and Alec Barrett
        Date: 2024/02/28
        Modified: 2026/03/30
        Version 2.0.0
#>

function Get-DriverInfo {
    <#
        .SYNOPSIS
            Retrieves driver information from remote machines using runspaces.
        .DESCRIPTION
            Queries installed driver information on one or more remote machines in
            parallel. Supports two search modes based on the DriverName input:

            - .inf drivers: Uses Get-WindowsDriver -Online to query the driver
              store. Returns structured objects with no text parsing required.

            - Loaded drivers (.sys, display name, or path): Uses driverquery /v /fo csv
              to find matching entries by module name, display name, or path. Supports
              wildcard patterns.

            Uses Invoke-RunspacePool for concurrent execution and Invoke-Command for
            remote queries.
        .PARAMETER ComputerName
            One or more target machine names. Accepts pipeline input.
        .PARAMETER DriverName
            One or more driver names to search for. Can be .inf filenames (e.g.
            "oem12.inf"), .sys filenames, display names, or file paths.
            Wildcards are supported for non-.inf searches.
        .PARAMETER ThrottleLimit
            Maximum concurrent runspaces. Default 50.
        .PARAMETER TimeoutMinutes
            Minutes before a runspace is stopped. Default 5.
        .EXAMPLE
            $list = Get-Content "$env:USERPROFILE\Desktop\Lists\Target_Machines.txt"
            Get-DriverInfo -ComputerName $list -DriverName "oem12.inf"
        .EXAMPLE
            Get-DriverInfo -ComputerName "PC001","PC002" -DriverName "e1d65x64.sys"
        .EXAMPLE
            Get-DriverInfo -ComputerName "PC001" -DriverName "Intel*"
        .NOTES
            Written by Skyler Werner and Alec Barrett
            Version: 2.0.0
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [string[]]
        $ComputerName,

        [Parameter(Mandatory, Position = 1)]
        [string[]]
        $DriverName,

        [Parameter()]
        [ValidateRange(1, 300)]
        [int]
        $ThrottleLimit = 50,

        [Parameter()]
        [ValidateRange(1, 120)]
        [int]
        $TimeoutMinutes = 5
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

    # --- Input cleanup ---
    $targets = @(Format-ComputerList $collectedNames -ToUpper)
    if ($targets.Count -eq 0) {
        Write-Warning "No valid computer names provided."
        return
    }


    # --- Connectivity check ---
    Write-Host ""
    Write-Host "Checking for online machines..."

    $pingResults = Test-ConnectionAsJob -ComputerName $targets

    $onlineList  = @()
    $allResults  = @()

    foreach ($ping in $pingResults) {
        if ($ping.Reachable -eq $true) {
            $onlineList += $ping.ComputerName
        }
        else {
            $allResults += [PSCustomObject][ordered]@{
                ComputerName = $ping.ComputerName
                Status       = "Offline"
                DriverType   = $null
                FileName     = $null
                DisplayName  = $null
                Version      = $null
                Date         = $null
                Vendor       = $null
                ClassName    = $null
                Path         = $null
                Comment      = $null
            }
        }
    }

    if ($onlineList.Count -eq 0) {
        Write-Host "No online machines found." -ForegroundColor Yellow
        $allResults | Sort-Object ComputerName | Format-Table -AutoSize | Out-Host
        return
    }


    # --- Scriptblock executed in each runspace (one per machine) ---
    $driverInfoBlock = {

        $computer    = $args[0]
        $driverNames = [string[]]($args[1] | ForEach-Object { "$_" })

        $PhaseTracker[$computer] = "Querying Drivers"

        try {
            $remoteResults = Invoke-Command -ComputerName $computer -ErrorAction Stop `
                -ArgumentList @(, $driverNames) -ScriptBlock {

                param($searchNames)
                $searchNames = [string[]]($searchNames | ForEach-Object { "$_" })

                $results = @()

                foreach ($driverName in $searchNames) {

                    # -----------------------------------------------------------
                    #  .inf driver handling -- Get-WindowsDriver
                    # -----------------------------------------------------------
                    if ($driverName -match '\.inf$') {

                        $storeDrivers = @(Get-WindowsDriver -Online |
                            Where-Object { $_.OriginalFileName -like "*$driverName" -or
                                           $_.Driver -like "*$driverName" })

                        if ($storeDrivers.Count -gt 0) {
                            foreach ($drv in ($storeDrivers | Sort-Object Driver)) {
                                $drvDate = $null
                                if ($drv.Date) {
                                    $drvDate = $drv.Date.ToString("yyyy.MM.dd")
                                }

                                $results += [PSCustomObject][ordered]@{
                                    DriverType  = "INF (Driver Store)"
                                    FileName    = $drv.Driver
                                    DisplayName = $drv.ProviderName
                                    Version     = $drv.Version
                                    Date        = $drvDate
                                    Vendor      = $drv.ProviderName
                                    ClassName   = $drv.ClassName
                                    Path        = $drv.OriginalFileName
                                    Comment     = $null
                                }
                            }
                        }
                        else {
                            $results += [PSCustomObject][ordered]@{
                                DriverType  = "INF"
                                FileName    = $driverName
                                DisplayName = $null
                                Version     = $null
                                Date        = $null
                                Vendor      = $null
                                ClassName   = $null
                                Path        = $null
                                Comment     = "No driver found"
                            }
                        }
                    }

                    # -----------------------------------------------------------
                    #  .sys / display name / path -- driverquery output
                    # -----------------------------------------------------------
                    else {
                        $useWildcard = $driverName -match '\*'

                        $allDrivers = driverquery.exe /v /fo csv | ConvertFrom-CSV

                        if ($useWildcard) {
                            $matched = @($allDrivers | Where-Object {
                                $_."Module Name"  -like $driverName -or
                                $_."Display Name" -like $driverName -or
                                $_.Path            -like $driverName
                            })
                        }
                        else {
                            $matched = @($allDrivers | Where-Object {
                                $_."Module Name"  -eq $driverName -or
                                $_."Display Name" -eq $driverName -or
                                $_.Path            -eq $driverName
                            })
                        }

                        if ($matched.Count -eq 0) {
                            $results += [PSCustomObject][ordered]@{
                                DriverType  = "System"
                                FileName    = $driverName
                                DisplayName = $null
                                Version     = $null
                                Date        = $null
                                Vendor      = $null
                                ClassName   = $null
                                Path        = $null
                                Comment     = "No driver found"
                            }
                        }
                        else {
                            foreach ($drv in $matched) {
                                $fileVer = $null
                                if ($drv.Path -and (Test-Path $drv.Path)) {
                                    $fileVer = (Get-Item $drv.Path).VersionInfo.FileVersion
                                }

                                $results += [PSCustomObject][ordered]@{
                                    DriverType  = "System"
                                    FileName    = $drv."Module Name"
                                    DisplayName = $drv."Display Name"
                                    Version     = $fileVer
                                    Date        = $null
                                    Vendor      = $null
                                    ClassName   = $null
                                    Path        = $drv.Path
                                    Comment     = $null
                                }
                            }
                        }
                    }
                }

                return $results
            }

            # Map remote results to output objects
            foreach ($r in @($remoteResults)) {
                [PSCustomObject][ordered]@{
                    ComputerName = $computer
                    Status       = "Online"
                    DriverType   = $r.DriverType
                    FileName     = $r.FileName
                    DisplayName  = $r.DisplayName
                    Version      = $r.Version
                    Date         = $r.Date
                    Vendor       = $r.Vendor
                    ClassName    = $r.ClassName
                    Path         = $r.Path
                    Comment      = $r.Comment
                }
            }
        }
        catch {
            [PSCustomObject][ordered]@{
                ComputerName = $computer
                Status       = "Online"
                DriverType   = $null
                FileName     = $null
                DisplayName  = $null
                Version      = $null
                Date         = $null
                Vendor       = $null
                ClassName    = $null
                Path         = $null
                Comment      = "Failed: $_"
            }
        }
    }


    # --- Build argument sets (one per machine) ---
    $argumentSets = @(
        foreach ($machine in $onlineList) {
            , @($machine, $DriverName)
        }
    )


    # --- Execute via Invoke-RunspacePool ---
    $runspaceParams = @{
        ScriptBlock    = $driverInfoBlock
        ArgumentList   = $argumentSets
        ThrottleLimit  = $ThrottleLimit
        TimeoutMinutes = $TimeoutMinutes
        ActivityName   = "Get Driver Info"
    }

    $poolResults = Invoke-RunspacePool @runspaceParams


    # --- Post-processing: normalize timed-out/failed results ---
    foreach ($result in $poolResults) {
        if ($result -isnot [PSCustomObject]) { continue }

        if ($null -ne $result.PSObject.Properties['Status']) {
            $allResults += $result
            continue
        }

        # Timeout guard -- result has no Status property
        $allResults += [PSCustomObject][ordered]@{
            ComputerName = $result.ComputerName
            Status       = "Online"
            DriverType   = $null
            FileName     = $null
            DisplayName  = $null
            Version      = $null
            Date         = $null
            Vendor       = $null
            ClassName    = $null
            Path         = $null
            Comment      = if ($null -ne $result.Comment) { $result.Comment } else { "Task Stopped" }
        }
    }


    # --- Output ---
    $allResults | Sort-Object -Property @(
        @{ Expression = "Status";       Descending = $true  }
        @{ Expression = "Comment";      Descending = $false }
        @{ Expression = "ComputerName"; Descending = $false }
        @{ Expression = "FileName";     Descending = $false }
    ) | Format-Table ComputerName, Status, DriverType, FileName, DisplayName, Version, Date, Comment -AutoSize | Out-Host


    } # end end
}

Export-ModuleMember Get-DriverInfo
