# DOTS formatting comment

function Get-UserBasedUninstallStatus {
    <#
        .SYNOPSIS
            Checks the status of scheduled user-based software uninstalls.
        .DESCRIPTION
            Follow-up companion to Uninstall-UserBasedSoftware. Accepts one or
            more result CSVs from previous runs and checks each machine that had
            a scheduled (AtLogOn) task to determine whether:

            - The software was successfully removed
            - The scheduled task self-deleted as expected
            - The log file captured an exit code

            Uses the same .psd1 configuration file as the original function to
            locate discovery paths. The -SoftwareName parameter selects the
            configuration entry and validates that the CSVs match.

            When multiple CSVs contain the same ComputerName + AD_Username pair,
            only the entry from the newest CSV is checked.
        .PARAMETER SoftwareName
            The name of the software to check, matching a key in the .psd1
            configuration file (e.g., Grammarly, Zoom, Spotify). Must match
            the software targeted by the CSVs being checked.
        .PARAMETER CsvPath
            One or more paths to result CSV files from Uninstall-UserBasedSoftware.
            Accepts pipeline input. All CSVs must target the same software.
        .PARAMETER ThrottleLimit
            Maximum concurrent machines. Default: 32
        .PARAMETER TimeoutMinutes
            Minutes before a machine's task is auto-stopped. Default: 15
        .EXAMPLE
            Get-UserBasedUninstallStatus -SoftwareName Grammarly -CsvPath ".\Grammarly_UserUninstall_2026-04-14-0900.csv"
        .EXAMPLE
            $csvs = Get-ChildItem "$env:USERPROFILE\Desktop\Patch-Results\Grammarly*.csv"
            Get-UserBasedUninstallStatus -SoftwareName Grammarly -CsvPath $csvs
        .EXAMPLE
            dir .\*Grammarly*.csv | Get-UserBasedUninstallStatus -SoftwareName Grammarly
        .NOTES
            Written by Skyler Werner
            Version: 1.0.0
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]
        $SoftwareName,

        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('FullName', 'Path')]
        [string[]]
        $CsvPath,

        [Parameter()]
        [ValidateRange(1, 300)]
        [int]
        $ThrottleLimit = 32,

        [Parameter()]
        [ValidateRange(1, 120)]
        [int]
        $TimeoutMinutes = 15,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        $collectedPaths = @()
    }

    process {
        foreach ($p in $CsvPath) {
            if ($p.Length -gt 0) {
                $collectedPaths += $p
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
        $discoveryPath = $config.DiscoveryPath
        $discoveryExe  = $config.DiscoveryExe

        # --- Resolve and validate CSV files ---
        $csvFiles = @()
        foreach ($p in $collectedPaths) {
            $resolved = Resolve-Path $p -ErrorAction SilentlyContinue
            if ($null -eq $resolved) {
                Write-Warning "CSV not found: $p"
                continue
            }
            foreach ($r in @($resolved)) {
                $csvFiles += Get-Item $r.Path
            }
        }

        if ($csvFiles.Count -eq 0) {
            Write-Warning "No valid CSV files provided."
            return
        }

        # --- Parse software names from filenames and validate ---
        $mismatchedFiles = @()
        foreach ($file in $csvFiles) {
            $baseName = $file.BaseName
            # Expected pattern: SoftwareName_UserUninstall_yyyy-MM-dd-HHmm
            if ($baseName -match '^(.+)_UserUninstall_\d{4}-\d{2}-\d{2}-\d{4}$') {
                if ($Matches[1] -ne $softwareName) {
                    $mismatchedFiles += "$($file.Name) (targets $($Matches[1]))"
                }
            }
            else {
                Write-Warning (
                    "Filename '$($file.Name)' does not match expected pattern " +
                    "'SoftwareName_UserUninstall_yyyy-MM-dd-HHmm.csv'. Skipping."
                )
            }
        }

        if ($mismatchedFiles.Count -gt 0) {
            Write-Error (
                "These CSVs do not match -SoftwareName '$softwareName': " +
                ($mismatchedFiles -join ', ')
            )
            return
        }

        Write-Host "Software:  $softwareName"
        Write-Host "CSV files: $($csvFiles.Count)"

        # --- Import and deduplicate ---
        $allRows = @()
        foreach ($file in $csvFiles) {
            $fileDate = $file.LastWriteTime
            $rows = @(Import-Csv $file.FullName)
            foreach ($row in $rows) {
                $row | Add-Member -NotePropertyName '_FileDate' -NotePropertyValue $fileDate -Force
                $allRows += $row
            }
        }

        # Filter to scheduled/pending rows only
        $pendingRows = @($allRows | Where-Object {
            $_.TaskType -eq 'Scheduled' -and $_.TaskExitCode -eq 'Pending'
        })

        if ($pendingRows.Count -eq 0) {
            Write-Warning "No scheduled/pending entries found in the provided CSVs."
            return
        }

        # Deduplicate: keep newest CSV entry per ComputerName + AD_Username
        $grouped = $pendingRows | Group-Object { "$($_.ComputerName)|$($_.AD_Username)" }
        $deduped = @()
        foreach ($group in $grouped) {
            $newest = $group.Group | Sort-Object _FileDate | Select-Object -Last 1
            $deduped += $newest
        }

        Write-Host "Pending entries: $($deduped.Count) (after dedup from $($pendingRows.Count) total)"

        # --- Ping check ---
        $uniqueComputers = @($deduped | Select-Object -ExpandProperty ComputerName | Sort-Object -Unique)
        Write-Host "Checking for online machines..."
        $pingResults = Test-ConnectionAsJob -ComputerName $uniqueComputers
        $onlineSet  = @{}
        $offlineSet = @{}
        foreach ($pr in $pingResults) {
            if ($pr.Reachable) {
                $onlineSet[$pr.ComputerName] = $true
            }
            else {
                $offlineSet[$pr.ComputerName] = $true
            }
        }

        $offlineResults = @()
        $onlineEntries  = @()
        foreach ($entry in $deduped) {
            if ($offlineSet.ContainsKey($entry.ComputerName)) {
                Write-Host "$($entry.ComputerName) Offline" -ForegroundColor Red
                $offlineResults += [PSCustomObject]@{
                    ComputerName    = $entry.ComputerName
                    Status          = 'Offline'
                    UserProfile     = $entry.UserProfile
                    AD_Username     = $entry.AD_Username
                    OriginalVersion = $entry.Version
                    FollowUpResult  = 'Unreachable'
                    TaskExists      = $null
                    SoftwarePresent = $null
                    LogExitCode     = $null
                    Comment         = ''
                }
            }
            else {
                $onlineEntries += $entry
            }
        }

        if ($onlineEntries.Count -eq 0) {
            Write-Warning "No pending machines are online."
            $allResults = @($offlineResults)
        }
        else {
            # --- Group entries by computer for runspace processing ---
            $byComputer = $onlineEntries | Group-Object ComputerName

            $configHash = @{
                SoftwareName = $softwareName
                DiscoveryPath = $discoveryPath
                DiscoveryExe = $discoveryExe
            }

            $argList = $byComputer | ForEach-Object {
                $userEntries = @($_.Group | ForEach-Object {
                    @{
                        UserProfile     = $_.UserProfile
                        AD_Username     = $_.AD_Username
                        OriginalVersion = $_.Version
                    }
                })
                ,@($_.Name, $userEntries, $configHash)
            }

            # --- Runspace scriptblock ---
            $scriptBlock = {
                $computer    = $args[0]
                $userEntries = $args[1]
                $cfg         = $args[2]

                $swName      = $cfg.SoftwareName
                $discPath    = $cfg.DiscoveryPath
                $discExe     = $cfg.DiscoveryExe

                $results = @()

                foreach ($entry in $userEntries) {
                    $userProfile = "$($entry.UserProfile)"
                    $adUsername  = "$($entry.AD_Username)"
                    $origVer    = "$($entry.OriginalVersion)"

                    $baseObject = [PSCustomObject]@{
                        ComputerName    = $computer
                        Status          = 'Online'
                        UserProfile     = $userProfile
                        AD_Username     = $adUsername
                        OriginalVersion = $origVer
                        FollowUpResult  = $null
                        TaskExists      = $null
                        SoftwarePresent = $null
                        LogExitCode     = $null
                        Comment         = ''
                    }

                    try {
                        $remResult = Invoke-Command -ComputerName $computer -ErrorAction Stop `
                            -ArgumentList @($adUsername, $userProfile, $swName, $discPath, $discExe) `
                            -ScriptBlock {
                            param($TargetUser, $Profile, $SwName, $DiscPath, $DiscExe)

                            # --- Check scheduled task ---
                            $taskName = "Uninstall_${SwName}_OnLogon_$TargetUser"
                            $taskExists = $false
                            try {
                                $task = Get-ScheduledTask -TaskName $taskName -ErrorAction Stop
                                if ($null -ne $task) { $taskExists = $true }
                            }
                            catch {
                                $taskExists = $false
                            }

                            # --- Check software presence ---
                            $fullDiscPath = "C:\Users\$Profile\$DiscPath\$DiscExe"
                            $softwarePresent = Test-Path $fullDiscPath

                            # --- Check log file ---
                            # Log format (from the scheduled task cmd chain):
                            #   <date> <time> - User 'jsmith' triggering uninstall.
                            #   <date> <time> - Exit code 0.
                            # The user name is on the trigger line; the exit code
                            # follows on the next line. Find the last trigger line
                            # for this user and grab the exit code from the line after.
                            $logFile = "C:\Temp\Logs\Uninstall_$SwName.log"
                            $logExitCode = $null
                            if (Test-Path $logFile) {
                                $logLines = @(Get-Content $logFile -ErrorAction SilentlyContinue)
                                $lastTriggerIndex = -1
                                for ($i = 0; $i -lt $logLines.Count; $i++) {
                                    if ($logLines[$i] -like "*$TargetUser*triggering uninstall*") {
                                        $lastTriggerIndex = $i
                                    }
                                }
                                if ($lastTriggerIndex -ge 0 -and ($lastTriggerIndex + 1) -lt $logLines.Count) {
                                    $exitLine = $logLines[$lastTriggerIndex + 1]
                                    if ($exitLine -match 'Exit code (\d+)') {
                                        $logExitCode = $Matches[1]
                                    }
                                }
                            }

                            # --- Clean up orphaned task if software is gone ---
                            $cleanedUp = $false
                            if (-not $softwarePresent -and $taskExists) {
                                try {
                                    Unregister-ScheduledTask -TaskName $taskName `
                                        -Confirm:$false -ErrorAction Stop
                                    $cleanedUp = $true
                                }
                                catch {
                                    # Will be noted in comment
                                }
                            }

                            return [PSCustomObject]@{
                                TaskExists      = $taskExists
                                SoftwarePresent = $softwarePresent
                                LogExitCode     = $logExitCode
                                CleanedUp       = $cleanedUp
                            }
                        }

                        $baseObject.TaskExists      = "$($remResult.TaskExists)"
                        $baseObject.SoftwarePresent = "$($remResult.SoftwarePresent)"
                        $baseObject.LogExitCode     = "$($remResult.LogExitCode)"

                        $swGone   = $remResult.SoftwarePresent -eq $false
                        $taskGone = $remResult.TaskExists -eq $false

                        if ($swGone -and $taskGone) {
                            if ($null -ne $remResult.LogExitCode) {
                                $baseObject.FollowUpResult = 'Completed'
                            }
                            else {
                                $baseObject.FollowUpResult = 'Completed (no log)'
                            }
                        }
                        elseif ($swGone -and -not $taskGone) {
                            if ($remResult.CleanedUp) {
                                $baseObject.FollowUpResult = 'Completed - Orphaned Task Removed'
                            }
                            else {
                                $baseObject.FollowUpResult = 'Completed - Orphaned Task'
                                $baseObject.Comment = 'Failed to remove orphaned task'
                            }
                        }
                        elseif (-not $swGone -and $taskGone) {
                            if ($null -ne $remResult.LogExitCode) {
                                $baseObject.FollowUpResult = 'Failed'
                                $baseObject.Comment = "Exit code: $($remResult.LogExitCode)"
                            }
                            else {
                                $baseObject.FollowUpResult = 'Failed (no log)'
                            }
                        }
                        else {
                            $baseObject.FollowUpResult = 'Pending'
                            $baseObject.Comment = 'Task has not fired yet'
                        }
                    }
                    catch {
                        $errMsg = ($_.Exception.Message) -replace ',', ';'
                        $baseObject.FollowUpResult = 'Error'
                        $baseObject.Comment = "Invoke-Command failed: $errMsg"
                    }

                    $results += $baseObject
                }

                return $results
            }

            # --- Execute via RunspacePool ---
            Write-Host "Checking $($onlineEntries.Count) pending entries on $($byComputer.Count) machine(s)..."
            $runspaceResults = @(Invoke-RunspacePool $scriptBlock $argList `
                -ThrottleLimit $ThrottleLimit `
                -TimeoutMinutes $TimeoutMinutes)

            # --- Normalize timed-out/failed results ---
            $onlineResults = foreach ($r in $runspaceResults) {
                if ($r -isnot [PSCustomObject]) { continue }
                if ($null -ne $r.PSObject.Properties['FollowUpResult']) {
                    $r
                }
                else {
                    # Build a result for each user entry on this timed-out machine
                    $timedOutComputer = $r.ComputerName
                    $timedOutEntries = @($onlineEntries | Where-Object {
                        $_.ComputerName -eq $timedOutComputer
                    })
                    if ($timedOutEntries.Count -eq 0) {
                        [PSCustomObject]@{
                            ComputerName    = $timedOutComputer
                            Status          = 'Online'
                            UserProfile     = $null
                            AD_Username     = $null
                            OriginalVersion = $null
                            FollowUpResult  = $null
                            TaskExists      = $null
                            SoftwarePresent = $null
                            LogExitCode     = $null
                            Comment         = if ($null -ne $r.Comment) { $r.Comment } else { 'Task Stopped' }
                        }
                    }
                    else {
                        foreach ($te in $timedOutEntries) {
                            [PSCustomObject]@{
                                ComputerName    = $timedOutComputer
                                Status          = 'Online'
                                UserProfile     = $te.UserProfile
                                AD_Username     = $te.AD_Username
                                OriginalVersion = $te.Version
                                FollowUpResult  = $null
                                TaskExists      = $null
                                SoftwarePresent = $null
                                LogExitCode     = $null
                                Comment         = if ($null -ne $r.Comment) { $r.Comment } else { 'Task Stopped' }
                            }
                        }
                    }
                }
            }

            $allResults = @($onlineResults) + @($offlineResults)
        }

        # --- Sort and display ---
        $reportProperties = @(
            'ComputerName', 'Status', 'UserProfile', 'AD_Username',
            'OriginalVersion', 'FollowUpResult', 'TaskExists',
            'SoftwarePresent', 'LogExitCode', 'Comment'
        )

        $sorted = $allResults | Select-Object $reportProperties | Sort-Object -Property (
            @{Expression = 'Status';         Descending = $true},
            @{Expression = 'FollowUpResult'; Descending = $false},
            @{Expression = 'ComputerName';   Descending = $false}
        )

        $sorted | Format-Table -Wrap | Out-Host

        # --- Log results ---
        $resultsRoot = "$env:USERPROFILE\Desktop\Patch-Results"
        if (-not (Test-Path $resultsRoot -PathType Container)) {
            mkdir $resultsRoot -Force > $null
        }

        $dateStamp = Get-Date -Format "yyyy-MM-dd-HHmm"
        $fileName = "${softwareName}_FollowUp_${dateStamp}"

        $sorted | Export-Csv "$resultsRoot\$fileName.csv" -Force -NoTypeInformation
        Write-Host "Report saved to $resultsRoot\$fileName.csv"

        if ($PassThru) { return $sorted }
    }
}
