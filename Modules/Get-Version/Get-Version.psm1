# DOTS formatting comment

<#
    .SYNOPSIS
        Retrieves file version info from remote machines using runspaces.
    .DESCRIPTION
        Queries file version information from one or more remote machines in parallel.
        Uses Invoke-RunspacePool for concurrent execution and Invoke-Command for remote
        file queries. Supports user-profile paths (USER token), hidden files, and both
        FileVersion and ProductVersion modes.

        Written by Skyler Werner
        Date: 2026/03/17
        Version 2.0.0
#>

function Get-Version {
    [CmdletBinding(DefaultParameterSetName = "Path")]
    param(
        # Target machine(s) to query
        [Parameter(Mandatory, Position = 0)]
        [string[]]
        $ComputerName,

        # File path(s) to check. Supports wildcards. Use USER token for user-profile paths.
        [Parameter(Mandatory, Position = 1, ParameterSetName = "Path")]
        [string[]]
        $Path,

        # Literal file path(s) to check. Wildcards are not expanded.
        [Parameter(Mandatory, Position = 1, ParameterSetName = "LiteralPath")]
        [string[]]
        $LiteralPath,

        # Use ProductVersion priority instead of FileVersionRaw
        [Parameter()]
        [switch]
        $Product,

        # Access hidden files (passed through to Get-Item / Get-ChildItem)
        [Parameter()]
        [switch]
        $Force,

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

        # Resolve which path set to use
        $targetPaths = if ($PSCmdlet.ParameterSetName -eq "LiteralPath") { $LiteralPath } else { $Path }
        $useLiteral  = $PSCmdlet.ParameterSetName -eq "LiteralPath"


        # --- Thread-safe dictionary for partial results from timed-out runspaces ---
        $partialResults = [System.Collections.Concurrent.ConcurrentDictionary[string, hashtable]]::new()


        # --- Scriptblock executed in each runspace (one per machine) ---
        $versionScriptBlock = {

            $computer       = $args[0]
            $paths          = $args[1]
            $useProduct     = $args[2]
            $useForce       = $args[3]
            $literal        = $args[4]
            $partialData    = $args[5]

            # Save computer name immediately so timeout still yields a result
            $partialData[$computer] = @{ ComputerName = $computer }

            $PhaseTracker[$computer] = "Version Check"

            # Build force parameter splat
            $forceParam = @{}
            if ($useForce) { $forceParam.Force = $true }

            # Build path parameter splat for the remote scriptblock arguments
            $remoteArgs = @($paths, $useProduct, $useForce, $literal)

            try {
                $remoteResult = Invoke-Command -ComputerName $computer -ArgumentList $remoteArgs -ErrorAction Stop -ScriptBlock {
                    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseUsingScopeModifierInNewRunspaces', '')]
                    param($PathList, $IsProduct, $IsForce, $IsLiteral)

                    $versions      = @()
                    $resolvedPaths = @()
                    $installed     = $false

                    # Build splat for -Force
                    $fSplat = @{}
                    if ($IsForce) { $fSplat.Force = $true }

                    foreach ($filePath in $PathList) {

                        # --- USER path handling ---
                        if ($filePath -cmatch 'USER') {

                            $userDirs = (Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue @fSplat).Name
                            $excludeUsers = @('Public', 'ADMINI~1')
                            $userDirs = $userDirs | Where-Object {
                                ($_ -notin $excludeUsers) -and ($_ -notmatch 'svc\d*\$')
                            }

                            foreach ($userName in $userDirs) {
                                $resolvedUserPath = $filePath.Replace('USER', $userName)

                                # Use Get-Item or Get-Item -LiteralPath based on mode
                                if ($IsLiteral) {
                                    $exists = Test-Path -LiteralPath $resolvedUserPath
                                    if (-not $exists) { continue }
                                    $items = @(Get-Item -LiteralPath $resolvedUserPath -ErrorAction SilentlyContinue @fSplat)
                                }
                                else {
                                    $exists = Test-Path -Path $resolvedUserPath
                                    if (-not $exists) { continue }
                                    $items = @(Get-Item -Path $resolvedUserPath -ErrorAction SilentlyContinue @fSplat)
                                }

                                foreach ($item in $items) {
                                    if ($null -eq $item) { continue }

                                    # File detection (archive attribute or no mode flags)
                                    if (($item.Mode -match 'a') -or ($item.Mode -eq '------')) {
                                        $installed = $true
                                        $resolvedPaths += $item.FullName

                                        # Version priority
                                        if ($IsProduct) {
                                            $ver = $item.VersionInfo.ProductVersion
                                            if ($null -eq $ver) { $ver = $item.VersionInfo.FileVersionRaw }
                                        }
                                        else {
                                            $ver = $item.VersionInfo.FileVersionRaw
                                            if ($null -eq $ver) { $ver = $item.VersionInfo.ProductVersion }
                                        }
                                        if ($null -eq $ver) { $ver = $item.VersionInfo.FileVersion }

                                        if ($null -ne $ver) {
                                            if ($ver.GetType().Name -match 'string') {
                                                try { $ver = [version]($ver.Replace(',', '.')) } catch {}
                                            }
                                            $versions += $ver
                                        }
                                    }
                                    # Directory detection
                                    elseif ($item.Mode -match 'd') {
                                        $installed = $true
                                        $resolvedPaths += $item.FullName
                                    }
                                }
                            }
                        }

                        # --- System path handling ---
                        else {
                            if ($IsLiteral) {
                                $exists = Test-Path -LiteralPath $filePath
                                if (-not $exists) { continue }
                                $items = @(Get-Item -LiteralPath $filePath -ErrorAction SilentlyContinue @fSplat)
                            }
                            else {
                                $exists = Test-Path -Path $filePath
                                if (-not $exists) { continue }
                                $items = @(Get-Item -Path $filePath -ErrorAction SilentlyContinue @fSplat)
                            }

                            foreach ($item in $items) {
                                if ($null -eq $item) { continue }

                                if (($item.Mode -match 'a') -or ($item.Mode -eq '------')) {
                                    $installed = $true
                                    $resolvedPaths += $item.FullName

                                    # Version priority
                                    if ($IsProduct) {
                                        $ver = $item.VersionInfo.ProductVersion
                                        if ($null -eq $ver) { $ver = $item.VersionInfo.FileVersionRaw }
                                    }
                                    else {
                                        $ver = $item.VersionInfo.FileVersionRaw
                                        if ($null -eq $ver) { $ver = $item.VersionInfo.ProductVersion }
                                    }
                                    if ($null -eq $ver) { $ver = $item.VersionInfo.FileVersion }

                                    if ($null -ne $ver) {
                                        if ($ver.GetType().Name -match 'string') {
                                            try { $ver = [version]($ver.Replace(',', '.')) } catch {}
                                        }
                                        $versions += $ver
                                    }
                                }
                                elseif ($item.Mode -match 'd') {
                                    $installed = $true
                                    $resolvedPaths += $item.FullName
                                }
                            }
                        }
                    }

                    # Return results from remote machine
                    [PSCustomObject]@{
                        Version      = $versions
                        Installed    = $installed
                        ResolvedPath = $resolvedPaths
                    }
                }

                # Build output object from remote result
                $versionOut  = if ($remoteResult.Version.Count -eq 1) { $remoteResult.Version[0] } elseif ($remoteResult.Version.Count -gt 1) { $remoteResult.Version } else { $null }
                $pathOut     = if ($remoteResult.ResolvedPath.Count -eq 1) { $remoteResult.ResolvedPath[0] } elseif ($remoteResult.ResolvedPath.Count -gt 1) { $remoteResult.ResolvedPath } else { $null }

                [PSCustomObject]@{
                    ComputerName = $computer
                    Installed    = $remoteResult.Installed
                    Version      = $versionOut
                    ResolvedPath = $pathOut
                }
            }
            catch {
                [PSCustomObject]@{
                    ComputerName = $computer
                    Installed    = $null
                    Version      = $null
                    ResolvedPath = $null
                    Comment      = "Failed: $_"
                }
            }
        }


        # --- Build argument sets (one per machine) ---
        $argumentSets = @(
            foreach ($machine in $ComputerName) {
                , @(
                    $machine,
                    $targetPaths,
                    [bool]$Product,
                    [bool]$Force,
                    $useLiteral,
                    $partialResults
                )
            }
        )


        # --- Execute via Invoke-RunspacePool ---
        $runspaceParams = @{
            ScriptBlock    = $versionScriptBlock
            ArgumentList   = $argumentSets
            ThrottleLimit  = $ThrottleLimit
            TimeoutMinutes = $Timeout
            ActivityName   = "Version Check"
        }

        $poolResults = Invoke-RunspacePool @runspaceParams


        # --- Post-processing: handle timed-out / incomplete results ---
        foreach ($result in $poolResults) {
            if ($result -isnot [PSCustomObject]) { continue }

            # Results with a Version or Installed property came from the scriptblock
            if ($null -ne $result.PSObject.Properties['Installed']) {
                $result
                continue
            }

            # Incomplete results from Invoke-RunspacePool (timed out / failed)
            $computerKey = $result.ComputerName
            $partial = $null
            if ($null -ne $computerKey) {
                $partialResults.TryGetValue($computerKey, [ref]$partial) > $null
            }

            [PSCustomObject]@{
                ComputerName = $computerKey
                Installed    = $null
                Version      = $null
                ResolvedPath = $null
                Comment      = if ($null -ne $result.Comment) { $result.Comment } else { "Task Stopped" }
            }
        }

    } # End process

} # End function


Export-ModuleMember -Function Get-Version
