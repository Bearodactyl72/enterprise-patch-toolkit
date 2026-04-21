# DOTS formatting comment

<#
    .SYNOPSIS
        Queries registry uninstall keys on remote machines for matching software entries.
    .DESCRIPTION
        Searches the registry uninstall hives on one or more remote machines in parallel
        for entries matching a software name pattern. Uses Invoke-RunspacePool for concurrent
        execution and Invoke-Command for remote registry queries.

        Written by Skyler Werner
        Date: 2021/05/05
        Modified: 2026/03/17
        Version 2.0.0
#>

function Get-RegistryKey {
    [CmdletBinding()]
    param(
        # Target machine(s) to query
        [Parameter(Mandatory, Position = 0)]
        [string[]]
        $ComputerName,

        # Software name pattern to match against DisplayName (regex via -match)
        [Parameter(Mandatory, Position = 1)]
        [string[]]
        $SoftwareName,

        # Registry key paths to search (defaults to both Uninstall hives)
        [Parameter(Position = 2)]
        [string[]]
        $RegistryKey = @(
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
        ),

        # Include the full registry key object(s) in the output
        [Parameter()]
        [switch]
        $IncludeKey,

        # Maximum concurrent runspaces (default 50)
        [Parameter()]
        [ValidateRange(1, 300)]
        [int]
        $ThrottleLimit = 50,

        # Timeout in minutes before runspaces are stopped (default 5)
        [Parameter()]
        [Alias("TimeoutMinutes")]
        [ValidateRange(1, 120)]
        [int]
        $Timeout = 5
    )


    process {

        # --- Thread-safe dictionary for partial results from timed-out runspaces ---
        $partialResults = [System.Collections.Concurrent.ConcurrentDictionary[string, hashtable]]::new()


        # --- Scriptblock executed in each runspace (one per machine) ---
        $registryScriptBlock = {

            $computer     = $args[0]
            $swNames      = $args[1]
            $regKeys      = $args[2]
            $wantFullKey  = $args[3]
            $partialData  = $args[4]

            # Save computer name immediately so timeout still yields a result
            $partialData[$computer] = @{ ComputerName = $computer }

            $PhaseTracker[$computer] = "Registry Check"

            $remoteArgs = @($swNames, $regKeys, $wantFullKey)

            try {
                $remoteResult = Invoke-Command -ComputerName $computer -ArgumentList $remoteArgs -ErrorAction Stop -ScriptBlock {
                    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseUsingScopeModifierInNewRunspaces', '')]
                    param($SoftwareNames, $RegistryKeys, $IncludeFullKey)

                    $displayNames = @()
                    $versions     = @()
                    $fullKeys     = @()

                    foreach ($regKey in $RegistryKeys) {
                        $children = Get-ChildItem $regKey -ErrorAction SilentlyContinue -Force
                        if ($null -eq $children) { continue }

                        $props = Get-ItemProperty $children.PSPath -ErrorAction SilentlyContinue

                        foreach ($prop in $props) {
                            foreach ($swName in $SoftwareNames) {
                                if ($prop.DisplayName -match $swName) {
                                    $displayNames += $prop.DisplayName

                                    if ($null -ne $prop.DisplayVersion) {
                                        $versions += $prop.DisplayVersion
                                    }

                                    if ($IncludeFullKey) {
                                        $fullKeys += $prop
                                    }

                                    break  # Don't match the same entry against multiple patterns
                                }
                            }
                        }
                    }

                    [PSCustomObject]@{
                        DisplayName = $displayNames
                        Version     = $versions
                        FullKeys    = $fullKeys
                    }
                }

                # Unwrap single-element arrays for cleaner output
                $nameOut = if ($remoteResult.DisplayName.Count -eq 1) { $remoteResult.DisplayName[0] }
                           elseif ($remoteResult.DisplayName.Count -gt 1) { $remoteResult.DisplayName }
                           else { $null }

                $verOut  = if ($remoteResult.Version.Count -eq 1) { $remoteResult.Version[0] }
                           elseif ($remoteResult.Version.Count -gt 1) { $remoteResult.Version }
                           else { $null }

                $result = [ordered]@{
                    ComputerName = $computer
                    KeyPresent   = ($remoteResult.DisplayName.Count -gt 0)
                    DisplayName  = $nameOut
                    Version      = $verOut
                }

                if ($wantFullKey) {
                    $keyOut = if ($remoteResult.FullKeys.Count -eq 1) { $remoteResult.FullKeys[0] }
                              elseif ($remoteResult.FullKeys.Count -gt 1) { $remoteResult.FullKeys }
                              else { $null }
                    $result.RegistryKey = $keyOut
                }

                [PSCustomObject]$result
            }
            catch {
                $result = [ordered]@{
                    ComputerName = $computer
                    KeyPresent   = $null
                    DisplayName  = $null
                    Version      = $null
                    Comment      = "Failed: $_"
                }

                if ($wantFullKey) {
                    $result.RegistryKey = $null
                }

                [PSCustomObject]$result
            }
        }


        # --- Build argument sets (one per machine) ---
        $argumentSets = @(
            foreach ($machine in $ComputerName) {
                , @(
                    $machine,
                    $SoftwareName,
                    $RegistryKey,
                    [bool]$IncludeKey,
                    $partialResults
                )
            }
        )


        # --- Execute via Invoke-RunspacePool ---
        $runspaceParams = @{
            ScriptBlock    = $registryScriptBlock
            ArgumentList   = $argumentSets
            ThrottleLimit  = $ThrottleLimit
            TimeoutMinutes = $Timeout
            ActivityName   = "Registry Check"
        }

        $poolResults = Invoke-RunspacePool @runspaceParams


        # --- Post-processing: handle timed-out / incomplete results ---
        foreach ($result in $poolResults) {
            if ($result -isnot [PSCustomObject]) { continue }

            # Results with a KeyPresent property came from the scriptblock
            if ($null -ne $result.PSObject.Properties['KeyPresent']) {
                $result
                continue
            }

            # Incomplete results from Invoke-RunspacePool (timed out / failed)
            $computerKey = $result.ComputerName
            $partial = $null
            if ($null -ne $computerKey) {
                $partialResults.TryGetValue($computerKey, [ref]$partial) > $null
            }

            $incompleteResult = [ordered]@{
                ComputerName = $computerKey
                KeyPresent   = $null
                DisplayName  = $null
                Version      = $null
                Comment      = if ($null -ne $result.Comment) { $result.Comment } else { "Task Stopped" }
            }

            if ([bool]$IncludeKey) {
                $incompleteResult.RegistryKey = $null
            }

            [PSCustomObject]$incompleteResult
        }

    } # End process

} # End function


Export-ModuleMember -Function Get-RegistryKey
