# DOTS formatting comment

function Remove-OldDrivers {
    <#
        .SYNOPSIS
            Removes outdated third-party driver packages from remote machines.
        .DESCRIPTION
            Queries each remote machine for third-party drivers via DISM, identifies
            driver packages where multiple versions of the same .inf file are installed,
            and removes all but the newest version using pnputil.

            Supports -WhatIf for dry-run mode, which enumerates which drivers would be
            removed without actually deleting them.

            Uses Invoke-RunspacePool for parallel execution across many machines.
        .PARAMETER ComputerName
            One or more target machine names. Accepts pipeline input.
        .PARAMETER ThrottleLimit
            Maximum concurrent runspaces. Default 50.
        .PARAMETER Timeout
            Timeout in minutes before a runspace is stopped. Default 15.
        .EXAMPLE
            $list = Get-Content "$env:USERPROFILE\Desktop\Lists\Target_Machines.txt"
            Remove-OldDrivers -ComputerName $list -WhatIf
        .EXAMPLE
            Remove-OldDrivers -ComputerName "PC001","PC002" -ThrottleLimit 10
        .NOTES
            Written by Skyler Werner
            Version: 2.0.0
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [string[]]
        $ComputerName,

        [Parameter()]
        [ValidateRange(1, 300)]
        [int]
        $ThrottleLimit = 50,

        [Parameter()]
        [Alias("TimeoutMinutes")]
        [ValidateRange(1, 120)]
        [int]
        $Timeout = 15,

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
            ComputerName   = $ping.ComputerName
            Status         = "Offline"
            DuplicatesFound = $null
            DriversRemoved = $null
            StaleDrivers   = $null
            Comment        = $null
        }
    }
}

if ($onlineList.Count -eq 0) {
    Write-Host "No online machines found." -ForegroundColor Yellow
    $allResults | Sort-Object ComputerName | Format-Table -AutoSize | Out-Host
    if ($PassThru) { return ($allResults | Sort-Object ComputerName) }
    return
}


# --- Capture WhatIf state for passing into runspaces ---
$isWhatIf = $WhatIfPreference


# --- Scriptblock executed in each runspace (one per machine) ---
$deleteDriversBlock = {

    $computer = $args[0]
    $dryRun   = $args[1]

    $PhaseTracker[$computer] = "Querying Drivers"

    try {
        $remoteResult = Invoke-Command -ComputerName $computer -ErrorAction Stop -ArgumentList $dryRun -ScriptBlock {
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseUsingScopeModifierInNewRunspaces', '')]
            param($IsDryRun)

            # --- Parse DISM driver output ---
            $dismOut = dism /online /get-drivers
            $lines   = $dismOut | Select-Object -Skip 10

            $operation = "theName"
            $drivers   = @()

            foreach ($line in $lines) {
                $txt = $($line.Split(':'))[1]

                switch ($operation) {
                    'theName' {
                        $drvName   = $txt
                        $operation = 'theFileName'
                    }
                    'theFileName' {
                        $fileName  = $txt.Trim()
                        $operation = 'theEntr'
                    }
                    'theEntr' {
                        $entr      = $txt.Trim()
                        $operation = 'theClassName'
                    }
                    'theClassName' {
                        $className = $txt.Trim()
                        $operation = 'theVendor'
                    }
                    'theVendor' {
                        $vendor    = $txt.Trim()
                        $operation = 'theDate'
                    }
                    'theDate' {
                        $parts = $txt.Split('/')
                        $date  = "{0:d4}.{1:d2}.{2:d2}" -f [int]$parts[2], [int]$parts[0], [int]$parts[1].Trim()
                        $operation = 'theVersion'
                    }
                    'theVersion' {
                        $version   = $txt.Trim()
                        $operation = 'theNull'

                        $drivers += [PSCustomObject][ordered]@{
                            Name      = $drvName
                            FileName  = $fileName
                            Vendor    = $vendor
                            Date      = $date
                            ClassName = $className
                            Version   = $version
                            Entr      = $entr
                        }
                    }
                    'theNull' {
                        $operation = 'theName'
                    }
                }
            }

            # --- Find driver packages with multiple versions ---
            $grouped = $drivers | Group-Object FileName | Where-Object { $_.Count -gt 1 }

            if ($null -eq $grouped -or @($grouped).Count -eq 0) {
                return [PSCustomObject]@{
                    DuplicatesFound = 0
                    Removed         = 0
                    StaleDrivers    = @()
                    Comment         = "No duplicate drivers found"
                }
            }

            # --- For each duplicate group, keep newest, identify the rest ---
            $staleList = @()
            $removed   = 0
            $errors    = @()

            foreach ($group in @($grouped)) {
                $sorted   = $group.Group | Sort-Object Date -Descending
                $newest   = $sorted[0]
                $toDelete = @($sorted | Select-Object -Skip 1)

                foreach ($old in $toDelete) {
                    $infName = $old.Name.Trim()

                    $staleList += "{0} [{1}] {2} {3} (keeping {4})" -f `
                        $infName, $old.ClassName, $old.Vendor, $old.Date, $newest.Date

                    if (-not $IsDryRun) {
                        $output = & pnputil.exe /delete-driver $infName
                        if ($LASTEXITCODE -eq 0) {
                            $removed++
                        }
                        else {
                            $errors += "$infName (exit $LASTEXITCODE)"
                        }
                    }
                }
            }

            $comment = $null
            if ($IsDryRun) {
                $comment = "WhatIf - no drivers removed"
            }
            elseif ($errors.Count -gt 0) {
                $comment = "Errors: " + ($errors -join "; ")
            }

            [PSCustomObject]@{
                DuplicatesFound = $staleList.Count
                Removed         = $removed
                StaleDrivers    = $staleList
                Comment         = $comment
            }
        }

        [PSCustomObject][ordered]@{
            ComputerName    = $computer
            Status          = "Online"
            DuplicatesFound = $remoteResult.DuplicatesFound
            DriversRemoved  = $remoteResult.Removed
            StaleDrivers    = @($remoteResult.StaleDrivers)
            Comment         = $remoteResult.Comment
        }
    }
    catch {
        [PSCustomObject][ordered]@{
            ComputerName    = $computer
            Status          = "Online"
            DuplicatesFound = $null
            DriversRemoved  = $null
            StaleDrivers    = @()
            Comment         = "Failed: $_"
        }
    }
}


# --- Build argument sets (one per machine) ---
$argumentSets = @(
    foreach ($machine in $onlineList) {
        , @($machine, $isWhatIf)
    }
)


# --- Execute via Invoke-RunspacePool ---
$runspaceParams = @{
    ScriptBlock    = $deleteDriversBlock
    ArgumentList   = $argumentSets
    ThrottleLimit  = $ThrottleLimit
    TimeoutMinutes = $Timeout
    ActivityName   = "Delete Old Drivers"
}

$poolResults = Invoke-RunspacePool @runspaceParams


# --- Post-processing ---
foreach ($result in $poolResults) {
    if ($result -isnot [PSCustomObject]) { continue }

    if ($null -ne $result.PSObject.Properties['Status']) {
        $allResults += $result
        continue
    }

    $allResults += [PSCustomObject][ordered]@{
        ComputerName    = $result.ComputerName
        Status          = "Online"
        DuplicatesFound = $null
        DriversRemoved  = $null
        StaleDrivers    = @()
        Comment         = if ($null -ne $result.Comment) { $result.Comment } else { "Task Stopped" }
    }
}


# --- Output summary table ---
$allResults | Sort-Object -Property @(
    @{ Expression = "Status";          Descending = $true  }
    @{ Expression = "DuplicatesFound"; Descending = $true  }
    @{ Expression = "DriversRemoved";  Descending = $true  }
    @{ Expression = "Comment";         Descending = $true  }
    @{ Expression = "ComputerName";    Descending = $false }
) | Format-Table ComputerName, Status, DuplicatesFound, DriversRemoved, Comment -AutoSize | Out-Host


# --- Detail section: list stale drivers per machine ---
$withStale = @($allResults | Where-Object { $_.StaleDrivers.Count -gt 0 })

if ($withStale.Count -gt 0) {
    Write-Host ""
    if ($isWhatIf) {
        Write-Host "Stale drivers that would be removed:" -ForegroundColor Yellow
    }
    else {
        Write-Host "Stale drivers removed:" -ForegroundColor Yellow
    }
    Write-Host ""

    foreach ($machine in ($withStale | Sort-Object ComputerName)) {
        Write-Host "  $($machine.ComputerName)" -ForegroundColor Cyan
        foreach ($drv in $machine.StaleDrivers) {
            Write-Host "    $drv"
        }
        Write-Host ""
    }
}

    if ($PassThru) { return $allResults }

    } # end
}
