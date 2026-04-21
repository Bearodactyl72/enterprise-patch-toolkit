# DOTS formatting comment

function Remove-VSCodeExtension {
    <#
        .SYNOPSIS
            Removes VS Code extensions from user profiles on remote machines.
        .DESCRIPTION
            Searches all user profiles on remote machines for VS Code extensions
            matching the specified patterns. For each affected user:

            1. If the user is actively logged in, creates an immediate scheduled
               task under their context to run 'code --force --uninstall-extension'
            2. If the user is logged off, creates an AtLogOn scheduled task that
               runs the uninstall at next login and self-deletes
            3. If code.exe cannot be found, falls back to direct folder deletion

            After CLI uninstall, verifies removal and cleans up any remaining
            extension folders as a fallback.

            Uses Invoke-RunspacePool for concurrent execution.
        .PARAMETER ComputerName
            One or more computer names to target. Accepts pipeline input.
        .PARAMETER ExtensionPattern
            One or more wildcard patterns matching extension folder names.
            Default: "github.copilot*"
        .PARAMETER ThrottleLimit
            Maximum concurrent machines. Default: 50
        .PARAMETER TimeoutMinutes
            Minutes before a machine's task is auto-stopped. Default: 10
        .EXAMPLE
            $list = Get-Content "$env:USERPROFILE\Desktop\Lists\VS_Code_Extension.txt"
            Remove-VSCodeExtension -ComputerName $list
        .EXAMPLE
            Remove-VSCodeExtension -ComputerName $list -ExtensionPattern "ms-python.python-*","tht13.python-*"
        .NOTES
            Written by Skyler Werner
            Version: 3.0.0
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [string[]]
        $ComputerName,

        [Parameter()]
        [string[]]
        $ExtensionPattern = @("github.copilot*"),

        [Parameter()]
        [ValidateRange(1, 300)]
        [int]
        $ThrottleLimit = 50,

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
        $online      = @($pingResults | Where-Object { $_.Reachable } | Select-Object -ExpandProperty ComputerName)
        $offline     = @($pingResults | Where-Object { -not $_.Reachable } | Select-Object -ExpandProperty ComputerName)

        $offlineResults = @()
        foreach ($pc in $offline) {
            Write-Host "$pc Offline" -ForegroundColor Red
            $offlineResults += [PSCustomObject]@{
                ComputerName  = $pc
                Status        = 'Offline'
                UserProfile   = $null
                Extensions    = $null
                Action        = $null
                Comment       = 'Offline'
            }
        }

        if ($online.Count -eq 0) {
            Write-Warning "No machines responded to ping."
            if ($PassThru) { return $offlineResults }
            return
        }

        # --- Build argument list ---
        $isWhatIf = $WhatIfPreference
        $argList = $online | ForEach-Object { ,@($_, $ExtensionPattern, $isWhatIf) }

        # --- Remote scriptblock ---
        $scriptBlock = {
            $computer = $args[0]
            $patterns = [string[]]($args[1] | ForEach-Object { "$_" })
            $whatIf   = $args[2]

            $results = @()

            try {
                $remoteOutput = Invoke-Command -ComputerName $computer -ScriptBlock {
                    param($extPatterns, $whatIf)

                    $output = @()

                    # --- Helper: extract extension ID from folder name ---
                    # Folder format is <extensionId>-<version> e.g. github.copilot-1.234.0
                    # Extension IDs can contain hyphens (github.copilot-chat)
                    # so we match the version as the last hyphen-separated segment
                    # that starts with digits.
                    function Get-ExtensionId {
                        param([string]$FolderName)
                        if ($FolderName -match '^(.+)-(\d+\..*)$') {
                            return $Matches[1]
                        }
                        return $FolderName
                    }

                    # --- Helper: find code.exe for a user profile ---
                    function Find-CodeExe {
                        param([string]$UserProfile)
                        # System install
                        $systemPaths = @(
                            "C:\Program Files\Microsoft VS Code\bin\code.cmd"
                            "C:\Program Files (x86)\Microsoft VS Code\bin\code.cmd"
                        )
                        foreach ($p in $systemPaths) {
                            if (Test-Path $p) { return $p }
                        }
                        # Per-user install
                        $userPath = "C:\Users\$UserProfile\AppData\Local\Programs\Microsoft VS Code\bin\code.cmd"
                        if (Test-Path $userPath) { return $userPath }
                        return $null
                    }

                    # --- Helper: get active user from quser ---
                    function Get-ActiveUser {
                        $quserLine = quser.exe 2>$null |
                            Where-Object { $_ -like '* Active *' } |
                            Select-Object -First 1
                        if ($null -eq $quserLine) { return $null }
                        $activeUser = $quserLine.Trim().Split(' ')[0]
                        if ($activeUser.StartsWith('>')) {
                            $activeUser = $activeUser.Substring(1)
                        }
                        return $activeUser
                    }

                    $activeUser = Get-ActiveUser

                    # --- Discover extensions per user ---
                    $userExtensions = @{}
                    foreach ($pattern in $extPatterns) {
                        $searchPath = "C:\Users\*\.vscode\extensions\$pattern"
                        $found = @(Get-ChildItem $searchPath -Directory -ErrorAction SilentlyContinue)
                        foreach ($dir in $found) {
                            # Extract username from path: C:\Users\<user>\.vscode\...
                            $userProfile = $dir.FullName.Split('\')[2]
                            if (-not $userExtensions.ContainsKey($userProfile)) {
                                $userExtensions[$userProfile] = @()
                            }
                            $userExtensions[$userProfile] += $dir
                        }
                    }

                    if ($userExtensions.Count -eq 0) {
                        $output += [PSCustomObject]@{
                            UserProfile = $null
                            Extensions  = $null
                            Action      = 'None'
                            Comment     = 'No matching extensions found'
                        }
                        return $output
                    }

                    # --- Process each user ---
                    foreach ($userProfile in $userExtensions.Keys) {
                        $extDirs = @($userExtensions[$userProfile])
                        $extIds = @($extDirs | ForEach-Object { Get-ExtensionId $_.Name }) | Sort-Object -Unique
                        $extNames = ($extIds -join ', ')
                        $codeExe = Find-CodeExe -UserProfile $userProfile

                        $userResult = [PSCustomObject]@{
                            UserProfile = $userProfile
                            Extensions  = $extNames
                            Action      = $null
                            Comment     = ''
                        }

                        $isActive = ($null -ne $activeUser -and $activeUser -eq $userProfile)

                        if ($null -eq $codeExe) {
                            # --- FALLBACK: No code.exe, delete folders directly ---
                            if ($whatIf) {
                                $userResult.Action = 'WhatIf'
                                $userResult.Comment = "Would delete $($extDirs.Count) folder(s) (no code.exe)"
                            }
                            else {
                                foreach ($dir in $extDirs) {
                                    Remove-Item $dir.FullName -Recurse -Force -ErrorAction SilentlyContinue
                                }
                                $userResult.Action = 'Deleted'
                                $userResult.Comment = "No code.exe found - deleted $($extDirs.Count) folder(s)"
                            }
                            $output += $userResult
                            continue
                        }

                        # Build uninstall command for all extensions
                        $uninstallParts = @()
                        foreach ($extId in $extIds) {
                            $uninstallParts += "`"$codeExe`" --force --uninstall-extension $extId"
                        }
                        $uninstallCmd = $uninstallParts -join ' & '

                        if ($isActive) {
                            # --- IMMEDIATE: user is logged in ---
                            if ($whatIf) {
                                $userResult.Action = 'WhatIf'
                                $userResult.Comment = "Would uninstall via CLI (user active): $extNames"
                            }
                            else {
                                $taskName = "Remove_VSCode_Ext_Immediate_$userProfile"
                                $exitCodeFile = "C:\Windows\Temp\vscode_ext_exit_$userProfile.txt"
                                Remove-Item $exitCodeFile -Force -ErrorAction SilentlyContinue

                                $cmdLine = "/c $uninstallCmd & echo %errorlevel% > `"$exitCodeFile`""
                                $taskAction = New-ScheduledTaskAction -Execute "cmd.exe" -Argument $cmdLine
                                $taskPrincipal = New-ScheduledTaskPrincipal -UserId $userProfile -LogonType Interactive

                                $taskParams = @{
                                    TaskName  = $taskName
                                    Action    = $taskAction
                                    Principal = $taskPrincipal
                                }
                                Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
                                Register-ScheduledTask @taskParams | Out-Null
                                Start-ScheduledTask -TaskName $taskName

                                # Wait for completion (up to 120 seconds)
                                $waitTimeout = 120
                                while (-not (Test-Path $exitCodeFile) -and $waitTimeout -gt 0) {
                                    Start-Sleep -Seconds 2
                                    $waitTimeout -= 2
                                }

                                $exitCode = $null
                                if (Test-Path $exitCodeFile) {
                                    $exitCode = (Get-Content $exitCodeFile -ErrorAction SilentlyContinue).Trim()
                                    Remove-Item $exitCodeFile -Force -ErrorAction SilentlyContinue
                                }

                                Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue

                                # Verify removal and clean up any remaining folders
                                $remaining = @()
                                foreach ($dir in $extDirs) {
                                    if (Test-Path $dir.FullName) {
                                        Remove-Item $dir.FullName -Recurse -Force -ErrorAction SilentlyContinue
                                        $remaining += $dir.Name
                                    }
                                }

                                $userResult.Action = 'Uninstalled'
                                if ($remaining.Count -gt 0) {
                                    $userResult.Comment = "CLI exit: $exitCode - fallback-deleted: $($remaining -join ', ')"
                                }
                                elseif ($null -ne $exitCode) {
                                    $userResult.Comment = "CLI exit: $exitCode"
                                }
                                else {
                                    $userResult.Comment = "CLI timed out - folders cleaned up"
                                }
                            }
                        }
                        else {
                            # --- SCHEDULED: user is logged off, run at next logon ---
                            if ($whatIf) {
                                $userResult.Action = 'WhatIf'
                                $userResult.Comment = "Would schedule logon task: $extNames"
                            }
                            else {
                                $taskName = "Remove_VSCode_Ext_OnLogon_$userProfile"
                                $logPath = "C:\Temp\Logs"
                                $logFile = "$logPath\Remove_VSCode_Extensions.log"

                                if (-not (Test-Path $logPath -PathType Container)) {
                                    mkdir $logPath -Force | Out-Null
                                }

                                $part1 = "echo %date% %time% - User '$userProfile' uninstalling VS Code extensions. >> `"$logFile`""
                                $part2 = $uninstallCmd
                                $part3 = "echo %date% %time% - Exit code %errorlevel%. >> `"$logFile`""
                                $part4 = "schtasks /delete /tn `"$taskName`" /f"
                                $cmdLine = "/c $part1 & $part2 & $part3 & $part4"

                                $taskAction = New-ScheduledTaskAction -Execute 'cmd.exe' -Argument $cmdLine
                                $taskTrigger = New-ScheduledTaskTrigger -AtLogOn -User $userProfile
                                $taskPrincipal = New-ScheduledTaskPrincipal -UserId $userProfile -LogonType Interactive

                                $taskParams = @{
                                    TaskName  = $taskName
                                    Action    = $taskAction
                                    Trigger   = $taskTrigger
                                    Principal = $taskPrincipal
                                    Force     = $true
                                }
                                Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
                                Register-ScheduledTask @taskParams | Out-Null

                                $userResult.Action = 'Scheduled'
                                $userResult.Comment = "Will uninstall at next logon"
                            }
                        }

                        $output += $userResult
                    }

                    return $output
                } -ArgumentList (,$patterns), $whatIf -ErrorAction Stop

                # Unpack remote results into per-user result objects
                foreach ($item in @($remoteOutput)) {
                    $results += [PSCustomObject]@{
                        ComputerName  = $computer
                        Status        = 'Online'
                        UserProfile   = $item.UserProfile
                        Extensions    = $item.Extensions
                        Action        = $item.Action
                        Comment       = "$($item.Comment)"
                    }
                }
            }
            catch {
                $errMsg = ($_.Exception.Message) -replace ',', ';'
                $results += [PSCustomObject]@{
                    ComputerName  = $computer
                    Status        = 'Online'
                    UserProfile   = $null
                    Extensions    = $null
                    Action        = $null
                    Comment       = "Failed: $errMsg"
                }
            }

            return $results
        }

        # --- Execute via RunspacePool ---
        if ($isWhatIf) {
            Write-Host "WhatIf: Scanning $($online.Count) machines for VS Code extensions..."
        }
        else {
            Write-Host "Removing VS Code extensions from $($online.Count) machines..."
        }
        $runspaceResults = @(Invoke-RunspacePool $scriptBlock $argList `
            -ThrottleLimit $ThrottleLimit `
            -TimeoutMinutes $TimeoutMinutes)

        # --- Normalize timed-out/failed results ---
        $onlineResults = foreach ($r in $runspaceResults) {
            if ($r -isnot [PSCustomObject]) { continue }
            if ($null -ne $r.PSObject.Properties['UserProfile']) {
                $r
            }
            else {
                [PSCustomObject]@{
                    ComputerName = $r.ComputerName
                    Status       = 'Online'
                    UserProfile  = $null
                    Extensions   = $null
                    Action       = $null
                    Comment      = if ($null -ne $r.Comment) { $r.Comment } else { 'Task Stopped' }
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
        if ($PassThru) { return $sorted }
    }
}
