# DOTS formatting comment

function Uninstall-UserBasedSoftware {
    <#
        .SYNOPSIS
            Uninstalls per-user software from remote machines using scheduled tasks.
        .DESCRIPTION
            Discovers user-installed software by scanning user profile directories on
            remote machines, resolves profile folder names to AD accounts, then
            uninstalls via user-context scheduled tasks.

            For active users, creates an immediate scheduled task that kills the
            process and runs the uninstaller. For logged-off users, creates an
            AtLogOn task that fires at next login and self-deletes.

            Software configuration (paths, exe names, arguments) is read from the
            companion .psd1 file alongside this script. The .psd1 contains entries
            for each supported software, selected via the -SoftwareName parameter.

            Uses Invoke-RunspacePool for concurrent execution.
        .PARAMETER SoftwareName
            The name of the software to uninstall, matching a key in the .psd1
            configuration file (e.g., Grammarly, Zoom, Spotify).
        .PARAMETER ComputerName
            One or more computer names to target. Accepts pipeline input.
        .PARAMETER ThrottleLimit
            Maximum concurrent machines. Default: 32
        .PARAMETER TimeoutMinutes
            Minutes before a machine's task is auto-stopped. Default: 30
        .EXAMPLE
            $list = Get-Content "$env:USERPROFILE\Desktop\Lists\Grammarly.txt"
            Uninstall-UserBasedSoftware -SoftwareName Grammarly -ComputerName $list
        .EXAMPLE
            "PC01","PC02" | Uninstall-UserBasedSoftware -SoftwareName Zoom
        .NOTES
            Written by Skyler Werner
            Version: 2.0.0
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]
        $SoftwareName,

        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [string[]]
        $ComputerName,

        [Parameter()]
        [ValidateRange(1, 300)]
        [int]
        $ThrottleLimit = 32,

        [Parameter()]
        [ValidateRange(1, 120)]
        [int]
        $TimeoutMinutes = 30,

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

        # --- Load configuration ---
        $configPath = Join-Path $PSScriptRoot 'Uninstall-UserBasedSoftware.psd1'
        if (-not (Test-Path $configPath)) {
            Write-Error "Configuration file not found: $configPath"
            return
        }
        $allConfigs = Import-PowerShellDataFile -Path $configPath

        if (-not $allConfigs.ContainsKey($SoftwareName)) {
            $validNames = ($allConfigs.Keys | Sort-Object) -join ', '
            Write-Error "Unknown software '$SoftwareName'. Valid options: $validNames"
            return
        }

        $config        = $allConfigs[$SoftwareName]
        $softwareName  = $SoftwareName
        $processName   = $config.ProcessName
        $discoveryPath = $config.DiscoveryPath
        $discoveryExe  = $config.DiscoveryExe
        $uninstallPath = $config.UninstallPath
        $uninstallExe  = $config.UninstallExe
        $uninstallArgs = $config.UninstallArgs

        Write-Host "Software: $softwareName"
        Write-Host "Config:   $configPath"

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
                ComputerName = $pc
                Status       = 'Offline'
                Version      = $null
                UserProfile  = $null
                AD_Username  = $null
                LogonStatus  = $null
                TaskType     = $null
                TaskExitCode = $null
                NewVersion   = $null
                Comment      = ''
            }
        }

        if ($online.Count -eq 0) {
            Write-Warning "No machines responded to ping."
            if ($PassThru) { return $offlineResults }
        }

        # --- Build argument list ---
        # Pack config as a hashtable so it survives runspace serialization
        $configHash = @{
            SoftwareName  = $softwareName
            ProcessName   = $processName
            DiscoveryPath = $discoveryPath
            DiscoveryExe  = $discoveryExe
            UninstallPath = $uninstallPath
            UninstallExe  = $uninstallExe
            UninstallArgs = $uninstallArgs
        }

        $isWhatIf = $WhatIfPreference
        $argList = $online | ForEach-Object { ,@($_, $configHash, $isWhatIf) }

        # --- Runspace scriptblock (one per machine) ---
        $scriptBlock = {
            $computer = $args[0]
            $cfg      = $args[1]
            $whatIf   = $args[2]

            # Unpack configuration
            $softwareName  = $cfg.SoftwareName
            $processName   = $cfg.ProcessName
            $discoveryPath = $cfg.DiscoveryPath
            $discoveryExe  = $cfg.DiscoveryExe
            $uninstallPath = $cfg.UninstallPath
            $uninstallExe  = $cfg.UninstallExe
            $uninstallArgs = $cfg.UninstallArgs

            Import-Module ActiveDirectory -ErrorAction SilentlyContinue

            $results = @()

            # =============================================================
            # PHASE 1: DISCOVERY - find software in user profiles
            # =============================================================
            $discoveryData = @()

            $baseObject = [PSCustomObject]@{
                ComputerName = $computer
                Status       = 'Online'
                Version      = $null
                UserProfile  = $null
                AD_Username  = $null
                LogonStatus  = $null
                TaskType     = $null
                TaskExitCode = $null
                NewVersion   = $null
                Comment      = ''
            }

            try {
                $remoteDiscovery = Invoke-Command -ComputerName $computer -ErrorAction Stop `
                    -ArgumentList $discoveryPath, $discoveryExe -ScriptBlock {
                    param($relPath, $discExe)

                    $appPath = "C:\Users\*\$relPath\$discExe"
                    $installations = @(Get-ChildItem -Path $appPath -ErrorAction SilentlyContinue)

                    if ($installations.Count -eq 0) {
                        return $null
                    }

                    foreach ($install in $installations) {
                        [PSCustomObject]@{
                            UserProfile = ($install.FullName.Split('\'))[2]
                            Version     = (Get-Item $install.FullName).VersionInfo.ProductVersion
                        }
                    }
                }

                if ($null -eq $remoteDiscovery) {
                    $obj = $baseObject.PSObject.Copy()
                    $obj.Version = 'Not Installed'
                    $discoveryData += $obj
                }
                else {
                    foreach ($item in @($remoteDiscovery)) {
                        $obj = $baseObject.PSObject.Copy()
                        $obj.UserProfile = "$($item.UserProfile)"
                        $obj.Version     = "$($item.Version)"
                        $discoveryData += $obj
                    }
                }
            }
            catch {
                $errMsg = ($_.Exception.Message) -replace ',', ';'
                $obj = $baseObject.PSObject.Copy()
                $obj.Comment = "Invoke-Command failed: $errMsg"
                return $obj
            }

            # =============================================================
            # PHASE 2: AD RESOLUTION - match profile folders to AD users
            # =============================================================
            foreach ($item in $discoveryData) {
                if ($null -eq $item.UserProfile -or $item.UserProfile -eq '') {
                    continue
                }

                try {
                    $directMatch = Get-ADUser -Filter "SamAccountName -eq '$($item.UserProfile)'" -ErrorAction SilentlyContinue
                    if ($directMatch) {
                        $item.AD_Username = $directMatch.SamAccountName
                    }
                    else {
                        $firstName = ($item.UserProfile -split '\.')[0]
                        $lastNameFromFolder = ($item.UserProfile -split '\.')[1]
                        $adFilter = "Name -like '$firstName*' -and Enabled -eq 'true'"
                        $plausibleMatches = @(Get-ADUser -Filter $adFilter -Properties Name, Surname)

                        if ($plausibleMatches.Count -gt 0) {
                            $perfectMatch = @($plausibleMatches | Where-Object { $_.Surname -eq $lastNameFromFolder })

                            if ($perfectMatch.Count -eq 1) {
                                $item.AD_Username = $perfectMatch[0].SamAccountName
                                $item.Comment = (
                                    "Profile '$($item.UserProfile)' matched AD user " +
                                    "'$($perfectMatch[0].Name)'"
                                )
                            }
                            elseif ($plausibleMatches.Count -eq 1) {
                                $item.AD_Username = $plausibleMatches[0].SamAccountName
                                $item.Comment = (
                                    "Profile '$($item.UserProfile)' matched AD user " +
                                    "'$($plausibleMatches[0].Name)'"
                                )
                            }
                            else {
                                $item.Comment = (
                                    "Orphaned: $($plausibleMatches.Count) AD matches " +
                                    "for '$firstName' but none with surname '$lastNameFromFolder'"
                                )
                            }
                        }
                        else {
                            $item.Comment = "Orphaned: No AD user found for '$firstName'"
                        }
                    }
                }
                catch {
                    $errMsg = ($_.Exception.Message) -replace ',', ';'
                    $item.Comment = "AD lookup error: $errMsg"
                }
            }

            # =============================================================
            # PHASE 3: REMEDIATION - uninstall via user-context sched tasks
            # =============================================================
            foreach ($item in $discoveryData) {
                if ($null -eq $item.AD_Username -or $item.AD_Username -eq '') {
                    $results += $item
                    continue
                }

                if ($whatIf) {
                    $item.TaskType = 'WhatIf'
                    $item.Comment  = "Would uninstall $softwareName for $($item.AD_Username)"
                    $results += $item
                    continue
                }

                $uninstallExePath = "C:\Users\$($item.UserProfile)\$uninstallPath\$uninstallExe"
                $discoveryExePath = "C:\Users\$($item.UserProfile)\$discoveryPath\$discoveryExe"

                try {
                    $remResult = Invoke-Command -ComputerName $computer -ErrorAction Stop `
                        -ArgumentList @(
                            $item.AD_Username,
                            $item.UserProfile,
                            $uninstallExePath,
                            $discoveryExePath,
                            $softwareName,
                            $processName,
                            $uninstallArgs
                        ) `
                        -ScriptBlock {
                        param(
                            $TargetADUser,
                            $TargetProfile,
                            $UninstallPath,
                            $DiscoveryPath,
                            $SwName,
                            $ProcName,
                            $UninstArgs
                        )

                        # --- Determine logon status ---
                        $activeUserLine = quser.exe 2>$null |
                            Where-Object { $_ -like '* Active *' } |
                            Select-Object -First 1

                        $activeUser = $null
                        if ($activeUserLine) {
                            $activeUser = $activeUserLine.Trim().Split(' ')[0]
                            if ($activeUser.StartsWith('>')) {
                                $activeUser = $activeUser.Substring(1)
                            }
                        }

                        $logonStatus = "Logged Off"
                        if ($activeUser -and ($activeUser -eq $TargetADUser)) {
                            $logonStatus = "Active"
                        }
                        elseif ($activeUser) {
                            $logonStatus = "Different user ($activeUser)"
                        }

                        $exitCode = $null
                        $taskType = $null

                        if ($logonStatus -eq "Active") {
                            # --- IMMEDIATE UNINSTALL ---
                            $taskName = "Uninstall_${SwName}_Immediate_$TargetADUser"
                            $exitCodeFile = "C:\Windows\Temp\exitcode_${SwName}_$TargetADUser.txt"
                            Remove-Item $exitCodeFile -Force -ErrorAction SilentlyContinue

                            $part1 = "taskkill /F /IM $ProcName /T"
                            $part2 = "`"$UninstallPath`" $UninstArgs"
                            $part3 = "echo %errorlevel% > `"$exitCodeFile`""
                            $cmdLine = "/c $part1 & $part2 & $part3"

                            $taskAction = New-ScheduledTaskAction -Execute "cmd.exe" -Argument $cmdLine
                            $taskPrincipal = New-ScheduledTaskPrincipal -UserId $TargetADUser -LogonType Interactive

                            $taskParams = @{
                                TaskName  = $taskName
                                Action    = $taskAction
                                Principal = $taskPrincipal
                            }
                            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
                            Register-ScheduledTask @taskParams | Out-Null
                            Start-ScheduledTask -TaskName $taskName

                            # Wait for completion (up to 300 seconds)
                            $waitTimeout = 300
                            while (-not (Test-Path $exitCodeFile) -and $waitTimeout -gt 0) {
                                Start-Sleep -Seconds 2
                                $waitTimeout -= 2
                            }

                            if (Test-Path $exitCodeFile) {
                                $exitCode = (Get-Content $exitCodeFile -ErrorAction SilentlyContinue).Trim()
                                Remove-Item $exitCodeFile -Force -ErrorAction SilentlyContinue
                            }
                            else {
                                $exitCode = "Timed out"
                            }

                            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
                            $taskType = "Immediate"
                        }
                        else {
                            # --- SCHEDULED UNINSTALL AT NEXT LOGON ---
                            $logPath = "C:\Temp\Logs"
                            $logFile = "$logPath\Uninstall_$SwName.log"
                            $taskName = "Uninstall_${SwName}_OnLogon_$TargetADUser"

                            if (-not (Test-Path $logPath -PathType Container)) {
                                mkdir $logPath -Force | Out-Null
                            }

                            $part1 = "echo %date% %time% - User '$TargetADUser' triggering uninstall. >> `"$logFile`""
                            $part2 = "taskkill /F /IM $ProcName /T"
                            $part3 = "`"$UninstallPath`" $UninstArgs"
                            $part4 = "echo %date% %time% - Exit code %errorlevel%. >> `"$logFile`""
                            $part5 = "schtasks /delete /tn `"$taskName`" /f"
                            $cmdLine = "/c $part1 & $part2 & $part3 & $part4 & $part5"

                            $taskAction = New-ScheduledTaskAction -Execute 'cmd.exe' -Argument $cmdLine
                            $taskTrigger = New-ScheduledTaskTrigger -AtLogOn -User $TargetADUser
                            $taskPrincipal = New-ScheduledTaskPrincipal -UserId $TargetADUser -LogonType Interactive

                            $taskParams = @{
                                TaskName  = $taskName
                                Action    = $taskAction
                                Trigger   = $taskTrigger
                                Principal = $taskPrincipal
                                Force     = $true
                            }
                            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
                            Register-ScheduledTask @taskParams | Out-Null

                            $exitCode = "Pending"
                            $taskType = "Scheduled"
                        }

                        # --- Check post-uninstall version ---
                        $versionAfter = "No change"
                        if ($taskType -eq 'Immediate') {
                            if (-not (Test-Path $DiscoveryPath)) {
                                $versionAfter = "Removed"
                            }
                        }

                        return [PSCustomObject]@{
                            LogonStatus  = $logonStatus
                            TaskType     = $taskType
                            TaskExitCode = $exitCode
                            NewVersion   = $versionAfter
                        }
                    }

                    $item.LogonStatus  = "$($remResult.LogonStatus)"
                    $item.TaskType     = "$($remResult.TaskType)"
                    $item.TaskExitCode = "$($remResult.TaskExitCode)"
                    $item.NewVersion   = "$($remResult.NewVersion)"
                }
                catch {
                    $errMsg = ($_.Exception.Message) -replace ',', ';'
                    $item.Comment = "Remediation failed: $errMsg"
                }

                $results += $item
            }

            return $results
        }

        # --- Execute via RunspacePool ---
        if ($isWhatIf) {
            Write-Host "WhatIf: Scanning $softwareName on $($online.Count) machine(s)..."
        }
        else {
            Write-Host "Processing $softwareName on $($online.Count) machine(s)..."
        }
        $runspaceResults = @(Invoke-RunspacePool $scriptBlock $argList `
            -ThrottleLimit $ThrottleLimit `
            -TimeoutMinutes $TimeoutMinutes)

        # --- Normalize timed-out/failed results ---
        $onlineResults = foreach ($r in $runspaceResults) {
            if ($r -isnot [PSCustomObject]) { continue }
            if ($null -ne $r.PSObject.Properties['Version']) {
                $r
            }
            else {
                [PSCustomObject]@{
                    ComputerName = $r.ComputerName
                    Status       = 'Online'
                    Version      = $null
                    UserProfile  = $null
                    AD_Username  = $null
                    LogonStatus  = $null
                    TaskType     = $null
                    TaskExitCode = $null
                    NewVersion   = $null
                    Comment      = if ($null -ne $r.Comment) { $r.Comment } else { 'Task Stopped' }
                }
            }
        }
        $allResults = @($onlineResults) + @($offlineResults)

        $reportProperties = @(
            'ComputerName', 'Status', 'Version', 'UserProfile', 'AD_Username',
            'LogonStatus', 'TaskType', 'TaskExitCode', 'NewVersion', 'Comment'
        )

        $sorted = $allResults | Select-Object $reportProperties | Sort-Object -Property (
            @{Expression = 'Status';       Descending = $true},
            @{Expression = 'LogonStatus';  Descending = $true},
            @{Expression = 'Version';      Descending = $false},
            @{Expression = 'Comment';      Descending = $false},
            @{Expression = 'ComputerName'; Descending = $false}
        )

        $sorted | Format-Table -Wrap | Out-Host

        # --- Log results to Patch-Results (skipped when -WhatIf is used) ---
        if (-not $isWhatIf) {
            $resultsRoot = "$env:USERPROFILE\Desktop\Patch-Results"
            if (-not (Test-Path $resultsRoot -PathType Container)) {
                mkdir $resultsRoot -Force > $null
            }

            $dateStamp = Get-Date -Format "yyyy-MM-dd-HHmm"
            $fileName = "${softwareName}_UserUninstall_${dateStamp}"

            $sorted | Export-Csv "$resultsRoot\$fileName.csv" -Force -NoTypeInformation
            Write-Host "Report saved to $resultsRoot\$fileName.csv"
        }

        if ($PassThru) { return $sorted }
    }
}
