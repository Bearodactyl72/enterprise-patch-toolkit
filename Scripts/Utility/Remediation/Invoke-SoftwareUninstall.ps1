# DOTS formatting comment

function Invoke-SoftwareUninstall {
    <#
        .SYNOPSIS
            Discovers and optionally uninstalls software on remote machines via
            registry uninstall keys.
        .DESCRIPTION
            Searches all four registry uninstall hives (HKLM/HKCU, 32/64-bit)
            on remote machines for entries matching a software name pattern.

            By default runs in discovery mode: returns what was found along with
            warnings about potential uninstall issues (missing exe, no silent
            flag, missing uninstall string, HKCU-only keys).

            With -Uninstall, executes the uninstall string for each matched
            entry. MSI entries get quiet/norestart flags automatically. EXE
            entries prefer QuietUninstallString when available.

            With -RemoveKeys, deletes the matched registry keys without running
            any uninstall (useful for cleaning stale entries).

            Uses Invoke-RunspacePool for concurrent execution.
        .PARAMETER ComputerName
            One or more computer names to target. Accepts pipeline input.
        .PARAMETER SoftwareName
            Display name pattern to match. Uses -match (regex) by default.
        .PARAMETER ExactMatch
            If specified, uses -eq instead of -match for name comparison.
        .PARAMETER Uninstall
            If specified, runs the uninstall string for each matched entry.
        .PARAMETER RemoveKeys
            If specified, deletes the matched registry keys without uninstalling.
        .PARAMETER ThrottleLimit
            Maximum concurrent machines. Default: 32
        .PARAMETER TimeoutMinutes
            Minutes before a machine's task is auto-stopped. Default: 10
        .EXAMPLE
            Invoke-SoftwareUninstall -ComputerName "PC01" -SoftwareName "Java"
            # Discovery mode: shows all matching entries with warnings.
        .EXAMPLE
            $list = Get-Content "$env:USERPROFILE\Desktop\Lists\Target_Machines.txt"
            Invoke-SoftwareUninstall -ComputerName $list -SoftwareName "^Java\s" -Uninstall
        .EXAMPLE
            Invoke-SoftwareUninstall -ComputerName "PC01" -SoftwareName "Old App" -ExactMatch -RemoveKeys
        .NOTES
            Written by Skyler Werner
            Version: 3.0.0
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [string[]]
        $ComputerName,

        [Parameter(Mandatory, Position = 1)]
        [string]
        $SoftwareName,

        [Parameter()]
        [switch]
        $ExactMatch,

        [Parameter()]
        [switch]
        $Uninstall,

        [Parameter()]
        [switch]
        $RemoveKeys,

        [Parameter()]
        [ValidateRange(1, 300)]
        [int]
        $ThrottleLimit = 32,

        [Parameter()]
        [ValidateRange(1, 120)]
        [int]
        $TimeoutMinutes = 10,

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

        # --- Sanitize input ---
        $targets = @(Format-ComputerList $collectedNames -ToUpper)
        if ($targets.Count -eq 0) {
            Write-Warning "No valid computer names provided."
            return
        }

        # --- Ping check ---
        Write-Host "Checking for online machines..."
        $pingResults = Test-ConnectionAsJob -ComputerName $targets
        $online  = @($pingResults | Where-Object { $_.Reachable } | Select-Object -ExpandProperty ComputerName)
        $offline = @($pingResults | Where-Object { -not $_.Reachable } | Select-Object -ExpandProperty ComputerName)

        $offlineResults = @()
        foreach ($pc in $offline) {
            Write-Host "$pc Offline" -ForegroundColor Red
            $offlineResults += [PSCustomObject]@{
                ComputerName    = $pc
                Status          = 'Offline'
                DisplayName     = $null
                Version         = $null
                UninstallType   = $null
                UninstallString = $null
                Action          = $null
                ExitCode        = $null
                Warnings        = $null
            }
        }

        if ($online.Count -eq 0) {
            Write-Warning "No machines responded to ping."
            if ($PassThru) { return $offlineResults }
        }

        # --- Build argument list ---
        $isWhatIf = $WhatIfPreference
        $argList = $online | ForEach-Object {
            ,@($_, $SoftwareName, [bool]$ExactMatch, [bool]$Uninstall, [bool]$RemoveKeys, $isWhatIf)
        }

        # --- Remote scriptblock ---
        $scriptBlock = {
            $computer    = $args[0]
            $swName      = [string]$args[1]
            $useExact    = [bool]$args[2]
            $doUninstall = [bool]$args[3]
            $doRemoveKey = [bool]$args[4]
            $whatIf      = $args[5]

            $results = @()

            try {
                $remoteOutput = Invoke-Command -ComputerName $computer -ArgumentList @($swName, $useExact, $doUninstall, $doRemoveKey, $whatIf) -ErrorAction Stop -ScriptBlock {
                    param($softwareName, $exactMatch, $uninstall, $removeKeys, $whatIf)

                    $output = @()

                    # --- Common silent flags for EXE-based uninstallers ---
                    $silentFlags = @(
                        '/S', '/s', '/silent', '/SILENT', '/quiet', '/QUIET',
                        '/qn', '/QN', '/norestart', '/NORESTART',
                        '-silent', '--silent', '-s', '--uninstall',
                        '/VERYSILENT', '/verysilent'
                    )

                    # --- Search all four uninstall hives ---
                    $regHives = @(
                        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'; Scope = 'HKLM' }
                        @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'; Scope = 'HKLM' }
                        @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'; Scope = 'HKCU' }
                        @{ Path = 'HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'; Scope = 'HKCU' }
                    )

                    foreach ($hive in $regHives) {
                        $children = @(Get-ChildItem $hive.Path -ErrorAction SilentlyContinue -Force)
                        if ($children.Count -eq 0) { continue }

                        $entries = @(Get-ItemProperty $children.PSPath -ErrorAction SilentlyContinue)

                        foreach ($entry in $entries) {
                            if ([string]::IsNullOrEmpty($entry.DisplayName)) { continue }

                            $matched = $false
                            if ($exactMatch) {
                                $matched = ($entry.DisplayName -eq $softwareName)
                            }
                            else {
                                $matched = ($entry.DisplayName -match $softwareName)
                            }
                            if (-not $matched) { continue }

                            # --- Determine uninstall type and string ---
                            $uString = $null
                            $uType   = 'Unknown'
                            $warnings = @()

                            $rawUninstall      = $entry.UninstallString
                            $rawQuietUninstall = $entry.QuietUninstallString

                            if ([string]::IsNullOrEmpty($rawUninstall) -and [string]::IsNullOrEmpty($rawQuietUninstall)) {
                                $warnings += 'No uninstall string'
                            }
                            elseif ($rawUninstall -match 'msiexec' -or $rawQuietUninstall -match 'msiexec') {
                                $uType = 'MSI'
                                # Extract the GUID from the uninstall string
                                $guidMatch = $null
                                $testStr = if (-not [string]::IsNullOrEmpty($rawUninstall)) { $rawUninstall } else { $rawQuietUninstall }
                                if ($testStr -match '\{[0-9A-Fa-f\-]+\}') {
                                    $guidMatch = $Matches[0]
                                }
                                if ($null -ne $guidMatch) {
                                    $uString = "msiexec.exe /X $guidMatch /quiet /norestart"
                                }
                                else {
                                    $uString = $rawUninstall
                                    $warnings += 'Could not extract MSI GUID'
                                }
                            }
                            else {
                                $uType = 'EXE'
                                # Prefer QuietUninstallString
                                if (-not [string]::IsNullOrEmpty($rawQuietUninstall)) {
                                    $uString = $rawQuietUninstall
                                }
                                else {
                                    $uString = $rawUninstall
                                }

                                # --- Validate: does the exe exist? ---
                                $exePath = $null
                                if ($uString -match '^"([^"]+)"') {
                                    $exePath = $Matches[1]
                                }
                                elseif ($uString -match '^([A-Za-z]:\\[^\s]+\.exe)') {
                                    $exePath = $Matches[1]
                                }

                                if ($null -ne $exePath -and -not (Test-Path $exePath)) {
                                    $warnings += "Exe not found: $exePath"
                                }

                                # --- Validate: silent flag present? ---
                                $hasSilent = $false
                                foreach ($flag in $silentFlags) {
                                    if ($uString -match [regex]::Escape($flag)) {
                                        $hasSilent = $true
                                        break
                                    }
                                }
                                if (-not $hasSilent) {
                                    $warnings += 'No silent flag'
                                }
                            }

                            if ($hive.Scope -eq 'HKCU') {
                                $warnings += 'HKCU key (may need user context)'
                            }

                            # --- Determine action ---
                            $action   = 'Discovered'
                            $exitCode = $null

                            if ($removeKeys -and $whatIf) {
                                $action = 'WhatIf'
                                $warnings += 'Would remove registry key'
                            }
                            elseif ($removeKeys) {
                                Remove-Item $entry.PSPath -Recurse -Force -ErrorAction SilentlyContinue
                                $action = 'KeyRemoved'
                            }
                            elseif ($uninstall -and $whatIf -and -not [string]::IsNullOrEmpty($uString)) {
                                $action = 'WhatIf'
                                $warnings += "Would run: $uString"
                            }
                            elseif ($uninstall -and -not [string]::IsNullOrEmpty($uString)) {
                                if ($uType -eq 'MSI') {
                                    # Parse out exe and arguments for Start-Process
                                    $msiArgs = $uString -replace '^msiexec\.exe\s*', ''
                                    $proc = Start-Process "msiexec.exe" -ArgumentList $msiArgs -Wait -PassThru 2>&1
                                    $exitCode = $proc.ExitCode
                                }
                                else {
                                    # EXE: run via cmd /c to handle varied quoting
                                    $proc = Start-Process "cmd.exe" -ArgumentList "/c $uString" -Wait -PassThru -WindowStyle Hidden 2>&1
                                    $exitCode = $proc.ExitCode
                                }
                                $action = 'Uninstalled'
                            }

                            $warnString = $null
                            if ($warnings.Count -gt 0) {
                                $warnString = $warnings -join '; '
                            }

                            $output += [PSCustomObject]@{
                                DisplayName     = $entry.DisplayName
                                Version         = $entry.DisplayVersion
                                UninstallType   = $uType
                                UninstallString = $uString
                                Action          = $action
                                ExitCode        = $exitCode
                                Warnings        = $warnString
                            }
                        }
                    }

                    if ($output.Count -eq 0) {
                        $output += [PSCustomObject]@{
                            DisplayName     = $null
                            Version         = $null
                            UninstallType   = $null
                            UninstallString = $null
                            Action          = 'None'
                            ExitCode        = $null
                            Warnings        = 'No matching entries found'
                        }
                    }

                    return $output
                }

                foreach ($item in @($remoteOutput)) {
                    $results += [PSCustomObject]@{
                        ComputerName    = $computer
                        Status          = 'Online'
                        DisplayName     = $item.DisplayName
                        Version         = "$($item.Version)"
                        UninstallType   = "$($item.UninstallType)"
                        UninstallString = "$($item.UninstallString)"
                        Action          = "$($item.Action)"
                        ExitCode        = $item.ExitCode
                        Warnings        = "$($item.Warnings)"
                    }
                }
            }
            catch {
                $errMsg = ($_.Exception.Message) -replace ',', ';'
                $results += [PSCustomObject]@{
                    ComputerName    = $computer
                    Status          = 'Online'
                    DisplayName     = $null
                    Version         = $null
                    UninstallType   = $null
                    UninstallString = $null
                    Action          = $null
                    ExitCode        = $null
                    Warnings        = "Failed: $errMsg"
                }
            }

            return $results
        }

        # --- Execute via RunspacePool ---
        $actionLabel = 'Scanning'
        if ($Uninstall)  { $actionLabel = 'Uninstalling' }
        if ($RemoveKeys) { $actionLabel = 'Removing keys for' }
        if ($isWhatIf -and ($Uninstall -or $RemoveKeys)) { $actionLabel = "WhatIf: $actionLabel" }
        Write-Host "$actionLabel '$SoftwareName' on $($online.Count) machine(s)..."

        $runspaceResults = @(Invoke-RunspacePool $scriptBlock $argList `
            -ThrottleLimit $ThrottleLimit `
            -TimeoutMinutes $TimeoutMinutes)

        # --- Normalize timed-out/failed results ---
        $onlineResults = foreach ($r in $runspaceResults) {
            if ($r -isnot [PSCustomObject]) { continue }
            if ($null -ne $r.PSObject.Properties['DisplayName']) {
                $r
            }
            else {
                [PSCustomObject]@{
                    ComputerName    = $r.ComputerName
                    Status          = 'Online'
                    DisplayName     = $null
                    Version         = $null
                    UninstallType   = $null
                    UninstallString = $null
                    Action          = $null
                    ExitCode        = $null
                    Warnings        = if ($null -ne $r.Comment) { $r.Comment } else { 'Task Stopped' }
                }
            }
        }
        $allResults = @($onlineResults) + @($offlineResults)

        $sorted = $allResults | Sort-Object -Property (
            @{Expression = 'Status'; Descending = $true},
            @{Expression = 'Action'; Descending = $true},
            @{Expression = 'ComputerName'; Descending = $false}
        )

        $sorted | Format-Table -AutoSize | Out-Host

        # --- Log results to Patch-Results (skipped when -WhatIf is used) ---
        if (-not $isWhatIf) {
            $resultsRoot = "$env:USERPROFILE\Desktop\Patch-Results"
            if (-not (Test-Path $resultsRoot -PathType Container)) {
                mkdir $resultsRoot -Force > $null
            }

            $dateStamp = Get-Date -Format "yyyy-MM-dd-HHmm"
            $safeName = ($SoftwareName -replace '[^a-zA-Z0-9._-]', '_')
            $actionTag = 'Discovery'
            if ($Uninstall)  { $actionTag = 'Uninstall' }
            if ($RemoveKeys) { $actionTag = 'RemoveKeys' }
            $fileName = "${actionTag}_${safeName}_${dateStamp}"

            $logProperties = @('ComputerName', 'Status', 'DisplayName', 'Version',
                'UninstallType', 'Action', 'ExitCode', 'Warnings')

            $sorted | Select-Object $logProperties |
                Export-Csv "$resultsRoot\$fileName.csv" -Force -NoTypeInformation

            Write-Host "Log saved to $resultsRoot\$fileName.csv"
        }

        if ($PassThru) { return $sorted }
    }
}
