# DOTS formatting comment

function Remove-UserRegistryKey {
    <#
        .SYNOPSIS
            Removes user-profile (HKU) registry keys left behind by uninstalled software.
        .DESCRIPTION
            Searches HKEY_USERS SID-based registry hives on remote machines for keys matching
            a software name pattern, then removes them. Used to clean up remnants after software
            is removed from the approved list and uninstalls don't execute cleanly.

            Supports targeting specific AD users (-UserName) or scanning all user profiles
            on the remote machine (-AllUsers).

            Uses Invoke-RunspacePool for parallel execution across many machines.
            Supports -WhatIf for dry-run mode.
        .PARAMETER ComputerName
            One or more target machine names. Accepts pipeline input.
        .PARAMETER SoftwareName
            Regex pattern(s) matched against registry key names and DisplayName properties
            (via -match). Use -ExactMatch to switch to literal -eq comparison.
        .PARAMETER UserName
            One or more AD usernames whose SIDs are resolved via Get-ADUser before
            connecting to remote machines. Mutually exclusive with -AllUsers.
        .PARAMETER AllUsers
            Enumerate all user SIDs from the remote machine's HKEY_USERS hive.
            Mutually exclusive with -UserName.
        .PARAMETER RegistrySubPath
            Optional sub-path(s) under the user's SID hive to narrow the search scope.
            For example: "SOFTWARE\WebEx", "SOFTWARE\Cisco".
            If omitted, searches both {SID}\SOFTWARE and {SID}_Classes.
        .PARAMETER ExactMatch
            When specified, matches SoftwareName using -eq instead of -match.
        .PARAMETER ThrottleLimit
            Maximum concurrent runspaces. Default 50.
        .PARAMETER Timeout
            Timeout in minutes before a runspace is stopped. Default 5.
        .EXAMPLE
            $list = Get-Content "$env:USERPROFILE\Desktop\Lists\Target_Machines.txt"
            Remove-UserRegistryKey -ComputerName $list -SoftwareName "WebEx" -UserName "john.doe" -WhatIf
        .EXAMPLE
            Remove-UserRegistryKey -ComputerName "PC001" -SoftwareName "Cisco" -AllUsers
        .EXAMPLE
            Remove-UserRegistryKey -ComputerName "PC001" `
                -SoftwareName "WebEx" -AllUsers -RegistrySubPath "SOFTWARE\WebEx"
        .NOTES
            Written by Skyler Werner
            Version: 2.0.0
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [string[]]
        $ComputerName,

        [Parameter(Mandatory, Position = 1)]
        [string[]]
        $SoftwareName,

        [Parameter(Mandatory, ParameterSetName = "ByUser")]
        [string[]]
        $UserName,

        [Parameter(Mandatory, ParameterSetName = "AllUsers")]
        [switch]
        $AllUsers,

        [Parameter()]
        [string[]]
        $RegistrySubPath,

        [Parameter()]
        [switch]
        $ExactMatch,

        [Parameter()]
        [ValidateRange(1, 300)]
        [int]
        $ThrottleLimit = 50,

        [Parameter()]
        [Alias("TimeoutMinutes")]
        [ValidateRange(1, 120)]
        [int]
        $Timeout = 5,

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


# --- SID resolution (ByUser mode) ---
$sidMap = @{}  # SID -> UserName

if ($PSCmdlet.ParameterSetName -eq "ByUser") {
    Write-Host ""
    Write-Host "Resolving user SIDs..."

    foreach ($user in $UserName) {
        try {
            $adUser = Get-ADUser $user -ErrorAction Stop
            $sid = $adUser.SID.Value
            $sidMap[$sid] = $user
            Write-Host "  $user -> $sid"
        }
        catch {
            Write-Host "  Failed to resolve SID for $user : $_" -ForegroundColor Red
        }
    }

    if ($sidMap.Count -eq 0) {
        Write-Host "No valid users resolved. Exiting." -ForegroundColor Yellow
        return
    }
}


    # --- Connectivity check ---
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
            UserName     = $null
            SID          = $null
            KeysFound    = $null
            KeysRemoved  = $null
            Comment      = $null
        }
    }
}

if ($onlineList.Count -eq 0) {
    Write-Host "No online machines found." -ForegroundColor Yellow
    $allResults | Sort-Object ComputerName | Format-Table -AutoSize | Out-Host
    return
}


# --- Capture WhatIf state for passing into runspaces ---
$isWhatIf = $WhatIfPreference


# --- Thread-safe dictionary for partial results ---
$partialResults = [System.Collections.Concurrent.ConcurrentDictionary[string, hashtable]]::new()


# --- Determine mode and SID data to pass ---
$isAllUsers   = $PSCmdlet.ParameterSetName -eq "AllUsers"
$sidMapArray  = if (-not $isAllUsers) {
    # Convert hashtable to array of key-value pairs for serialization
    @($sidMap.GetEnumerator() | ForEach-Object { @($_.Key, $_.Value) })
} else { @() }


# --- Scriptblock executed in each runspace (one per machine) ---
$cleanupScriptBlock = {

    $computer    = $args[0]
    $swNames     = $args[1]
    $exact       = $args[2]
    $subPaths    = $args[3]
    $dryRun      = $args[4]
    $enumAll     = $args[5]
    $sidPairs    = $args[6]
    $partialData = $args[7]

    # Save computer name immediately so timeout still yields a result
    $partialData[$computer] = @{ ComputerName = $computer }

    $PhaseTracker[$computer] = "HKU Cleanup"

    $remoteArgs = @($swNames, $exact, $subPaths, $dryRun, $enumAll, $sidPairs)

    try {
        $remoteResult = Invoke-Command -ComputerName $computer -ArgumentList $remoteArgs -ErrorAction Stop -ScriptBlock {
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseUsingScopeModifierInNewRunspaces', '')]
            param($SoftwareNames, $UseExact, $SubPaths, $IsDryRun, $EnumAll, $SidPairs)

            # --- Well-known SIDs to skip when enumerating ---
            $skipSids = @('.DEFAULT', 'S-1-5-18', 'S-1-5-19', 'S-1-5-20')

            # --- Build SID -> UserName map ---
            $sidUserMap = @{}
            if (-not $EnumAll) {
                for ($i = 0; $i -lt $SidPairs.Count; $i += 2) {
                    $sidUserMap[$SidPairs[$i]] = $SidPairs[$i + 1]
                }
            }

            # --- Determine target SIDs ---
            $targetSids = @()

            if ($EnumAll) {
                # Enumerate all user SIDs from HKEY_USERS
                $hkuChildren = Get-ChildItem "Registry::HKEY_USERS" -ErrorAction SilentlyContinue
                foreach ($child in $hkuChildren) {
                    $sidName = $child.PSChildName
                    # Skip well-known SIDs and _Classes duplicates
                    if ($sidName -in $skipSids) { continue }
                    if ($sidName -match '_Classes$') { continue }
                    if ($sidName -notmatch '^S-1-5-21-') { continue }
                    $targetSids += $sidName
                }
            }
            else {
                $targetSids = @($sidUserMap.Keys)
            }

            # --- Search and remove for each SID ---
            $results = @()

            foreach ($sid in $targetSids) {

                # Build search paths for this SID
                $searchPaths = @()

                if ($null -ne $SubPaths -and $SubPaths.Count -gt 0) {
                    foreach ($sp in $SubPaths) {
                        $searchPaths += "Registry::HKEY_USERS\$sid\$sp"
                    }
                }
                else {
                    $searchPaths += "Registry::HKEY_USERS\$sid\SOFTWARE"
                    $searchPaths += "Registry::HKEY_USERS\${sid}_Classes"
                }

                # Find matching keys
                $foundKeys = @()

                foreach ($searchPath in $searchPaths) {
                    if (-not (Test-Path $searchPath -ErrorAction SilentlyContinue)) { continue }

                    $children = Get-ChildItem $searchPath -ErrorAction SilentlyContinue -Force
                    if ($null -eq $children) { continue }

                    foreach ($child in $children) {
                        $keyName = $child.PSChildName
                        $isMatch = $false

                        # Match against key name
                        foreach ($swName in $SoftwareNames) {
                            if ($UseExact) {
                                if ($keyName -eq $swName) { $isMatch = $true; break }
                            }
                            else {
                                if ($keyName -match $swName) { $isMatch = $true; break }
                            }
                        }

                        # Also check DisplayName property if available
                        if (-not $isMatch) {
                            $props = Get-ItemProperty $child.PSPath -ErrorAction SilentlyContinue
                            if ($null -ne $props.DisplayName) {
                                foreach ($swName in $SoftwareNames) {
                                    if ($UseExact) {
                                        if ($props.DisplayName -eq $swName) { $isMatch = $true; break }
                                    }
                                    else {
                                        if ($props.DisplayName -match $swName) { $isMatch = $true; break }
                                    }
                                }
                            }
                        }

                        if ($isMatch) {
                            $foundKeys += $child.PSPath
                        }
                    }
                }

                # Skip this SID if no matching keys found
                if ($foundKeys.Count -eq 0) { continue }

                # Remove matching keys
                $removedCount = $null
                $comment      = $null

                if (-not $IsDryRun) {
                    $removedCount = 0
                    foreach ($keyPath in $foundKeys) {
                        try {
                            Remove-Item -Path $keyPath -Recurse -Force -ErrorAction Stop
                            $removedCount++
                        }
                        catch {
                            if ($null -eq $comment) { $comment = "" }
                            else { $comment += "; " }
                            $comment += "Failed to remove key: $_"
                        }
                    }
                }

                # Resolve username for output
                $userName = if ($sidUserMap.ContainsKey($sid)) { $sidUserMap[$sid] } else { $null }

                $results += [PSCustomObject]@{
                    UserName    = $userName
                    SID         = $sid
                    KeysFound   = $foundKeys.Count
                    KeysRemoved = $removedCount
                    Comment     = $comment
                }
            }

            # Return all results for this machine
            $results
        }

        # Emit one output row per user result from the remote machine
        if ($null -eq $remoteResult -or @($remoteResult).Count -eq 0) {
            # No matching keys found on any SID -- no rows emitted
            return
        }

        foreach ($userResult in @($remoteResult)) {
            [PSCustomObject][ordered]@{
                ComputerName = $computer
                Status       = "Online"
                UserName     = $userResult.UserName
                SID          = $userResult.SID
                KeysFound    = $userResult.KeysFound
                KeysRemoved  = $userResult.KeysRemoved
                Comment      = $userResult.Comment
            }
        }
    }
    catch {
        [PSCustomObject][ordered]@{
            ComputerName = $computer
            Status       = "Online"
            UserName     = $null
            SID          = $null
            KeysFound    = $null
            KeysRemoved  = $null
            Comment      = "Failed: $_"
        }
    }
}


# --- Build argument sets (one per machine) ---
$argumentSets = @(
    foreach ($machine in $onlineList) {
        , @(
            $machine,
            $SoftwareName,
            [bool]$ExactMatch,
            $RegistrySubPath,
            $isWhatIf,
            $isAllUsers,
            $sidMapArray,
            $partialResults
        )
    }
)


# --- Execute via Invoke-RunspacePool ---
$runspaceParams = @{
    ScriptBlock    = $cleanupScriptBlock
    ArgumentList   = $argumentSets
    ThrottleLimit  = $ThrottleLimit
    TimeoutMinutes = $Timeout
    ActivityName   = "HKU Registry Cleanup"
}

$poolResults = Invoke-RunspacePool @runspaceParams


# --- Post-processing ---
foreach ($result in $poolResults) {
    if ($result -isnot [PSCustomObject]) { continue }

    # Results with a Status property came from our scriptblock
    if ($null -ne $result.PSObject.Properties['Status']) {
        $allResults += $result
        continue
    }

    # Incomplete results from Invoke-RunspacePool (timed out / failed)
    $allResults += [PSCustomObject][ordered]@{
        ComputerName = $result.ComputerName
        Status       = "Online"
        UserName     = $null
        SID          = $null
        KeysFound    = $null
        KeysRemoved  = $null
        Comment      = if ($null -ne $result.Comment) { $result.Comment } else { "Task Stopped" }
    }
}


# --- Output results ---
$selectProps = @(
    "ComputerName",
    "Status",
    "UserName",
    "SID",
    "KeysFound",
    "KeysRemoved",
    "Comment"
)

    $sorted = $allResults | Select-Object $selectProps | Sort-Object -Property @(
        @{ Expression = "Status";       Descending = $true  }
        @{ Expression = "KeysFound";    Descending = $true  }
        @{ Expression = "ComputerName"; Descending = $false }
        @{ Expression = "UserName";     Descending = $false }
    )

    $sorted | Format-Table -AutoSize | Out-Host
    if ($PassThru) { return $sorted }

    } # end
}
