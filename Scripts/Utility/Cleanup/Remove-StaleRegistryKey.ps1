# DOTS formatting comment

function Remove-StaleRegistryKey {
    <#
        .SYNOPSIS
            Removes stale registry uninstall keys left behind by previous software versions.
        .DESCRIPTION
            Finds all registry uninstall entries matching a software name pattern on remote
            machines, determines which key belongs to the currently-installed version (by
            checking the actual file version on disk), and removes the rest.

            Uses Invoke-RunspacePool for parallel execution across many machines.
            Supports -WhatIf for dry-run mode.
        .PARAMETER ComputerName
            One or more target machine names. Accepts pipeline input.
        .PARAMETER SoftwareName
            Regex pattern(s) matched against DisplayName (via -match).
            Use -ExactMatch to switch to literal -eq comparison.
        .PARAMETER FilePath
            Path to the software executable on the remote machine, used to determine the
            currently-installed version via file version info. If omitted, falls back to
            InstallLocation from the registry key. If neither is available, the key with
            the highest DisplayVersion is treated as valid.
        .PARAMETER RegistryPath
            Registry hive path(s) to search. Defaults to both 32-bit and 64-bit
            Uninstall hives under HKLM.
        .PARAMETER ExactMatch
            When specified, matches SoftwareName against DisplayName using -eq instead
            of -match.
        .PARAMETER ThrottleLimit
            Maximum concurrent runspaces. Default 50.
        .PARAMETER Timeout
            Timeout in minutes before a runspace is stopped. Default 5.
        .EXAMPLE
            $list = Get-Content "$env:USERPROFILE\Desktop\Lists\Target_Machines.txt"
            Remove-StaleRegistryKey -ComputerName $list `
                -SoftwareName "Google Chrome" `
                -FilePath "C:\Program Files\Google\Chrome\Application\chrome.exe" `
                -ExactMatch -WhatIf
        .EXAMPLE
            Remove-StaleRegistryKey -ComputerName "PC001","PC002" `
                -SoftwareName "Adobe Reader" `
                -FilePath "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
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

        [Parameter(Position = 2)]
        [string]
        $FilePath,

        [Parameter()]
        [string[]]
        $RegistryPath = @(
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
        ),

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
            ComputerName     = $ping.ComputerName
            Status           = "Offline"
            SoftwareName     = $null
            SoftwarePath     = $null
            InstalledVersion = $null
            ValidRegVersion  = $null
            StaleRegVersions = $null
            KeysRemoved      = $null
            Comment          = $null
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


# --- Scriptblock executed in each runspace (one per machine) ---
$cleanupScriptBlock = {

    $computer    = $args[0]
    $swNames     = $args[1]
    $regKeys     = $args[2]
    $exact       = $args[3]
    $filePath    = $args[4]
    $dryRun      = $args[5]
    $partialData = $args[6]

    # Save computer name immediately so timeout still yields a result
    $partialData[$computer] = @{ ComputerName = $computer }

    $PhaseTracker[$computer] = "Registry Check"

    $remoteArgs = @($swNames, $regKeys, $exact, $filePath, $dryRun)

    try {
        $remoteResult = Invoke-Command -ComputerName $computer -ArgumentList $remoteArgs -ErrorAction Stop -ScriptBlock {
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseUsingScopeModifierInNewRunspaces', '')]
            param($SoftwareNames, $RegistryKeys, $UseExact, $ExeFilePath, $IsDryRun)

            # --- Find all matching registry keys ---
            $matchedKeys = @()

            foreach ($regKey in $RegistryKeys) {
                $children = Get-ChildItem $regKey -ErrorAction SilentlyContinue -Force
                if ($null -eq $children) { continue }

                $props = Get-ItemProperty $children.PSPath -ErrorAction SilentlyContinue

                foreach ($prop in $props) {
                    if ($null -eq $prop.DisplayName) { continue }

                    $isMatch = $false
                    foreach ($swName in $SoftwareNames) {
                        if ($UseExact) {
                            if ($prop.DisplayName -eq $swName) { $isMatch = $true; break }
                        }
                        else {
                            if ($prop.DisplayName -match $swName) { $isMatch = $true; break }
                        }
                    }

                    if ($isMatch) {
                        $matchedKeys += [PSCustomObject]@{
                            DisplayName    = $prop.DisplayName
                            DisplayVersion = $prop.DisplayVersion
                            PSPath         = $prop.PSPath
                            InstallLocation = $prop.InstallLocation
                        }
                    }
                }
            }

            # --- Edge case: no matching keys ---
            if ($matchedKeys.Count -eq 0) {
                return [PSCustomObject]@{
                    SoftwareName     = $null
                    SoftwarePath     = $null
                    InstalledVersion = $null
                    ValidRegVersion  = $null
                    StaleVersions    = $null
                    Removed          = $null
                    Comment          = "No matching registry keys found"
                }
            }

            # --- Safety: abort if multiple distinct software names matched ---
            $distinctNames = @($matchedKeys | ForEach-Object { $_.DisplayName } | Sort-Object -Unique)
            if ($distinctNames.Count -gt 1) {
                $nameList = $distinctNames -join ", "
                return [PSCustomObject]@{
                    SoftwareName     = $nameList
                    SoftwarePath     = $null
                    InstalledVersion = $null
                    ValidRegVersion  = $null
                    StaleVersions    = $null
                    Removed          = $null
                    Comment          = "Multiple software matched: $nameList. Use -ExactMatch or refine -SoftwareName"
                }
            }

            # --- Edge case: only one key, nothing to clean ---
            if ($matchedKeys.Count -eq 1) {
                return [PSCustomObject]@{
                    SoftwareName     = $matchedKeys[0].DisplayName
                    SoftwarePath     = $null
                    InstalledVersion = $null
                    ValidRegVersion  = $matchedKeys[0].DisplayVersion
                    StaleVersions    = $null
                    Removed          = $null
                    Comment          = $null
                }
            }

            # --- Determine installed file version ---
            $installedVersion = $null
            $resolvedPath     = $null
            $comment          = $null

            # Try provided FilePath first
            if ($ExeFilePath -and (Test-Path -LiteralPath $ExeFilePath -ErrorAction SilentlyContinue)) {
                $fileItem = Get-Item -LiteralPath $ExeFilePath -ErrorAction SilentlyContinue
                if ($null -ne $fileItem) {
                    $resolvedPath = $fileItem.FullName
                    $ver = $fileItem.VersionInfo.FileVersionRaw
                    if ($null -eq $ver) { $ver = $fileItem.VersionInfo.ProductVersion }
                    if ($null -eq $ver) { $ver = $fileItem.VersionInfo.FileVersion }
                    if ($null -ne $ver) {
                        $installedVersion = [string]$ver
                    }
                }
            }

            # Fallback: try InstallLocation from registry keys
            if ($null -eq $installedVersion) {
                foreach ($key in $matchedKeys) {
                    if ($key.InstallLocation -and (Test-Path -LiteralPath $key.InstallLocation -ErrorAction SilentlyContinue)) {
                        # Look for exe files in the install location
                        $exes = Get-ChildItem $key.InstallLocation -Filter "*.exe" -ErrorAction SilentlyContinue |
                                Select-Object -First 1
                        if ($null -ne $exes) {
                            $resolvedPath = $exes.FullName
                            $ver = $exes.VersionInfo.FileVersionRaw
                            if ($null -eq $ver) { $ver = $exes.VersionInfo.ProductVersion }
                            if ($null -eq $ver) { $ver = $exes.VersionInfo.FileVersion }
                            if ($null -ne $ver) {
                                $installedVersion = [string]$ver
                                break
                            }
                        }
                    }
                }
            }

            # --- Classify keys as valid or stale ---
            $validKey   = $null
            $staleKeys  = @()

            if ($null -ne $installedVersion) {
                # Match file version against DisplayVersion
                foreach ($key in $matchedKeys) {
                    if ($key.DisplayVersion -eq $installedVersion) {
                        $validKey = $key
                    }
                    else {
                        $staleKeys += $key
                    }
                }

                # Safety: if file version matches NO key, do not remove anything
                if ($null -eq $validKey) {
                    return [PSCustomObject]@{
                        SoftwareName     = $matchedKeys[0].DisplayName
                        SoftwarePath     = $resolvedPath
                        InstalledVersion = $installedVersion
                        ValidRegVersion  = $null
                        StaleVersions    = @($matchedKeys | ForEach-Object { $_.DisplayVersion })
                        Removed          = $null
                        Comment          = "File version $installedVersion does not match any registry key"
                    }
                }
            }
            else {
                # No file version available -- fall back to highest version
                $comment = "No file path available; used highest version as valid"

                $sorted = $matchedKeys | Sort-Object {
                    try { [version]$_.DisplayVersion } catch { [version]"0.0" }
                } -Descending

                $validKey  = $sorted[0]
                $staleKeys = @($sorted | Select-Object -Skip 1)
            }

            # --- Remove stale keys ---
            $removedCount = $null
            if ($staleKeys.Count -gt 0 -and -not $IsDryRun) {
                $removedCount = 0
                foreach ($stale in $staleKeys) {
                    try {
                        Remove-Item -Path $stale.PSPath -Recurse -Force -ErrorAction Stop
                        $removedCount++
                    }
                    catch {
                        if ($null -eq $comment) { $comment = "" }
                        else { $comment += "; " }
                        $comment += "Failed to remove $($stale.DisplayVersion): $_"
                    }
                }
            }

            # --- Return result ---
            [PSCustomObject]@{
                SoftwareName     = $validKey.DisplayName
                SoftwarePath     = $resolvedPath
                InstalledVersion = $installedVersion
                ValidRegVersion  = $validKey.DisplayVersion
                StaleVersions    = @($staleKeys | ForEach-Object { $_.DisplayVersion })
                Removed          = $removedCount
                Comment          = $comment
            }
        }

        # Build output object
        [PSCustomObject][ordered]@{
            ComputerName     = $computer
            Status           = "Online"
            SoftwareName     = $remoteResult.SoftwareName
            SoftwarePath     = $remoteResult.SoftwarePath
            InstalledVersion = $remoteResult.InstalledVersion
            ValidRegVersion  = $remoteResult.ValidRegVersion
            StaleRegVersions = $remoteResult.StaleVersions
            KeysRemoved      = $remoteResult.Removed
            Comment          = $remoteResult.Comment
        }
    }
    catch {
        [PSCustomObject][ordered]@{
            ComputerName     = $computer
            Status           = "Online"
            SoftwareName     = $null
            SoftwarePath     = $null
            InstalledVersion = $null
            ValidRegVersion  = $null
            StaleRegVersions = $null
            KeysRemoved      = $null
            Comment          = "Failed: $_"
        }
    }
}


# --- Build argument sets (one per machine) ---
$argumentSets = @(
    foreach ($machine in $onlineList) {
        , @(
            $machine,
            $SoftwareName,
            $RegistryPath,
            [bool]$ExactMatch,
            $FilePath,
            $isWhatIf,
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
    ActivityName   = "Stale Registry Cleanup"
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
        ComputerName     = $result.ComputerName
        Status           = "Online"
        SoftwareName     = $null
        SoftwarePath     = $null
        InstalledVersion = $null
        ValidRegVersion  = $null
        StaleRegVersions = $null
        KeysRemoved      = $null
        Comment          = if ($null -ne $result.Comment) { $result.Comment } else { "Task Stopped" }
    }
}


# --- Output results ---
$selectProps = @(
    "ComputerName",
    "Status",
    "SoftwareName",
    "SoftwarePath",
    "InstalledVersion",
    "ValidRegVersion",
    @{ Name = 'StaleRegVersions'; Expression = { $_.StaleRegVersions -join ', ' } },
    "KeysRemoved",
    "Comment"
)

    $sorted = $allResults | Select-Object $selectProps | Sort-Object -Property @(
        @{ Expression = "Status";           Descending = $true  }
        @{ Expression = "KeysRemoved";      Descending = $true  }
        @{ Expression = "StaleRegVersions"; Descending = $true  }
        @{ Expression = "ComputerName";     Descending = $false }
    )

    $sorted | Format-Table -AutoSize | Out-Host
    if ($PassThru) { return $sorted }

    } # end
}
