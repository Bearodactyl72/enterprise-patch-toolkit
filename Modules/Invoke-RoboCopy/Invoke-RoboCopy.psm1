# DOTS formatting comment

<#
    .SYNOPSIS
        Copies files or directories to multiple remote computers in parallel using robocopy and runspaces.
    .DESCRIPTION
        Stand-alone utility module for copying files or directories to one or more remote computers
        using robocopy over UNC paths. Uses a RunspacePool for parallel execution with live progress
        monitoring. Always copies recursively. Outputs a results table showing success or failure
        per computer.

        Written by Skyler Werner
        Date: 2026/03/16
        Version 1.0.0
#>

# Module-scoped variables for stale pool cleanup after Ctrl+C
$script:currentPool = $null
$script:currentRunspaces = $null


function Stop-RoboCopyRunspaceAsync {
    <#
        .SYNOPSIS
            Sends an async stop signal to a PowerShell instance and waits up to $TimeoutSeconds.
    #>
    param(
        [PowerShell]$PowerShell,
        [string]$Label = '',
        [int]$TimeoutSeconds = 15
    )

    try {
        $asyncResult = $PowerShell.BeginStop($null, $null)
        $stopped     = $asyncResult.AsyncWaitHandle.WaitOne(
            [TimeSpan]::FromSeconds($TimeoutSeconds)
        )

        if (-not $stopped -and $Label) {
            Write-Host "  Force-terminating $Label (not responding to stop signal)..." -ForegroundColor Red
        }
    }
    catch {
        # BeginStop can throw if the pipeline is already in a broken state
    }
}


function Invoke-RoboCopy {
    <#
        .SYNOPSIS
            Copies files or directories to multiple remote computers in parallel using robocopy.
        .DESCRIPTION
            Uses robocopy over admin share (UNC) paths with a RunspacePool for parallel execution.
            Always copies recursively (/E). Shows a live progress table during execution and outputs
            a final results table showing success or failure per computer.
        .PARAMETER ComputerName
            One or more target computer hostnames.
        .PARAMETER Path
            Source path to copy (supports wildcards).
        .PARAMETER LiteralPath
            Source path to copy (no wildcard expansion).
        .PARAMETER LocalizedDestination
            Destination written as a local path on each target machine (e.g. C:\Temp).
            Automatically converted to UNC (\\Computer\C$\Temp) for each target.
        .PARAMETER Timeout
            Minutes before a copy is considered timed out. Defaults based on source size.
        .PARAMETER Throttle
            Maximum concurrent copy operations. Default 20.
        .EXAMPLE
            Invoke-RoboCopy -ComputerName PC01,PC02,PC03 -Path "C:\Packages\App" -LocalizedDestination "C:\Temp"
        .EXAMPLE
            $computers = Get-Content .\computers.txt
            Invoke-RoboCopy -ComputerName $computers -LiteralPath "D:\Deploy\Patch[v2]" -LocalizedDestination "C:\Temp\Patches" -Throttle 30
    #>
    [CmdletBinding(DefaultParameterSetName = 'Path')]
    param(
        [Parameter(Mandatory, Position = 0)]
        [String[]]
        $ComputerName,

        [Parameter(Mandatory, Position = 1, ParameterSetName = 'Path')]
        [String]
        $Path,

        [Parameter(Mandatory, ParameterSetName = 'LiteralPath')]
        [String]
        $LiteralPath,

        [Parameter(Mandatory, Position = 2)]
        [Alias("Destination")]
        [String]
        $LocalizedDestination,

        [Parameter()]
        [ValidateRange(1, 120)]
        [Int32]
        $Timeout,

        [Parameter()]
        [ValidateRange(1, 100)]
        [Int32]
        $Throttle = 20,

        # Always overwrite destination files even if they appear identical
        [Parameter()]
        [Switch]
        $Force
    )


    process {

        #region --- Resolve Source Path ---

        if ($PSCmdlet.ParameterSetName -eq 'LiteralPath') {
            $sourcePath = $LiteralPath
            if (-not (Test-Path -LiteralPath $sourcePath)) {
                Write-Error "Source path not found: $sourcePath"
                return
            }
        }
        else {
            $sourcePath = $Path
            if (-not (Test-Path -Path $sourcePath)) {
                Write-Error "Source path not found: $sourcePath"
                return
            }
        }

        # Get the resolved full path for size calculation
        if ($PSCmdlet.ParameterSetName -eq 'LiteralPath') {
            $resolvedSource = (Get-Item -LiteralPath $sourcePath).FullName
        }
        else {
            $resolvedSource = (Resolve-Path -Path $sourcePath).Path
        }

        $itemFolder = Split-Path $resolvedSource -Leaf

        #endregion --- Resolve Source Path ---


        #region --- Calculate Source Size ---

        $sourceBytes = 0
        if (Test-Path $resolvedSource -PathType Container) {
            Get-ChildItem $resolvedSource -Recurse -File | ForEach-Object { $sourceBytes += $_.Length }
        }
        else {
            $sourceBytes = (Get-Item $resolvedSource).Length
        }

        if ($sourceBytes -lt 1MB) {
            $sizeDisplay = "{0:N1} KB" -f ($sourceBytes / 1KB)
        }
        elseif ($sourceBytes -lt 1GB) {
            $sizeDisplay = "{0:N1} MB" -f ($sourceBytes / 1MB)
        }
        else {
            $sizeDisplay = "{0:N2} GB" -f ($sourceBytes / 1GB)
        }

        # Dynamic timeout if not specified
        if (-not $PSBoundParameters.ContainsKey('Timeout')) {
            $divBy100MB = $sourceBytes / 104857600  # 100 MB
            if     ($divBy100MB -lt 0.1) { $Timeout = 5  }
            elseif ($divBy100MB -lt 1)   { $Timeout = 10 }
            elseif ($divBy100MB -lt 2)   { $Timeout = 20 }
            elseif ($divBy100MB -lt 5)   { $Timeout = 30 }
            elseif ($divBy100MB -lt 10)  { $Timeout = 45 }
            elseif ($divBy100MB -lt 50)  { $Timeout = 60 }
            else                         { $Timeout = 90 }
        }

        #endregion --- Calculate Source Size ---


        #region --- Stale Pool Cleanup ---

        if ($null -ne $script:currentPool) {
            Write-Warning "Cleaning up stale runspace pool from a previous interrupted run..."
            try {
                if ($null -ne $script:currentRunspaces) {
                    foreach ($rs in $script:currentRunspaces) {
                        Stop-RoboCopyRunspaceAsync -PowerShell $rs.PowerShell -TimeoutSeconds 5
                        try { $rs.PowerShell.Dispose() } catch {}
                    }
                }
                $script:currentPool.Close()
                $script:currentPool.Dispose()
            }
            catch {}
            $script:currentPool = $null
            $script:currentRunspaces = $null
        }

        #endregion --- Stale Pool Cleanup ---


        #region --- Build Robocopy ScriptBlock ---

        $roboCopyScript = {
            param($Computer, $SourcePath, $LocalDest, $ItemFolder, $ForceOverwrite)

            $result = [PSCustomObject]@{
                ComputerName = $Computer
                CopyComplete = $false
                Elapsed      = ''
                ExitCode     = $null
                Comment      = ''
            }

            # Update phase tracker if available
            if ($null -ne $PhaseTracker) {
                $PhaseTracker[$Computer] = "Copying..."
            }

            # Convert local destination to UNC path
            # e.g. C:\Temp -> \\COMPUTER\C$\Temp
            $uncDest = $LocalDest -replace '^([A-Za-z]):', ('\\{0}\$1$$' -f $Computer)
            $destPath = Join-Path $uncDest $ItemFolder

            # Remove stale file if a file exists where a directory is expected
            if (Test-Path $destPath -PathType Leaf) {
                Remove-Item $destPath -Force > $null
            }

            # Build robocopy arguments
            $robocopyArgs = @(
                "`"$SourcePath`""     # source directory
                "`"$destPath`""       # destination directory
                '/E'                  # copy subdirectories including empty ones (recurse)
                '/R:3'               # retry 3 times on failed copies
                '/W:5'               # wait 5 seconds between retries
                '/MT:4'              # multi-threaded copy (4 threads per target)
                '/NP'                # no progress percentage
                '/NDL'               # no directory listing
                '/NFL'               # no file listing
                '/NJH'               # no job header
                '/NJS'               # no job summary
            )

            if ($ForceOverwrite) {
                $robocopyArgs += '/IS'   # include same files (overwrite even if identical)
                $robocopyArgs += '/IT'   # include tweaked files (overwrite even if only timestamps differ)
            }

            $robocopyOutput = & robocopy @robocopyArgs 2>&1
            $robocopyExit   = $LASTEXITCODE
            $result.ExitCode = $robocopyExit

            # Robocopy exit codes: 0-7 = success (bitmask), 8+ = failure
            if ($robocopyExit -lt 8) {
                $result.CopyComplete = $true
                $result.Comment = switch ($robocopyExit) {
                    0 { "No changes - source and destination in sync" }
                    1 { "Files copied successfully" }
                    2 { "Extra files or directories detected at destination" }
                    3 { "Files copied, extras detected at destination" }
                    4 { "Mismatched files or directories detected" }
                    5 { "Files copied, mismatches detected" }
                    6 { "Extras and mismatches detected" }
                    7 { "Files copied, extras and mismatches detected" }
                }
            }
            else {
                $result.CopyComplete = $false
                $result.Comment = switch ($robocopyExit) {
                    8  { "Some files could not be copied (exit 8)" }
                    16 { "Fatal error - no files were copied (exit 16)" }
                    default { "Robocopy error (exit $robocopyExit)" }
                }
            }

            return $result
        }

        #endregion --- Build Robocopy ScriptBlock ---


        #region --- Create RunspacePool ---

        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

        $PhaseTracker = [Hashtable]::Synchronized(@{})
        $phaseEntry = [System.Management.Automation.Runspaces.SessionStateVariableEntry]::new(
            'PhaseTracker', $PhaseTracker, 'Synchronized hashtable for live phase tracking'
        )
        $iss.Variables.Add($phaseEntry)

        $pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(
            1, $Throttle, $iss, $Host
        )
        $pool.Open()
        $script:currentPool = $pool

        #endregion --- Create RunspacePool ---


        #region --- Submit Runspaces ---

        $runspaces = [System.Collections.Generic.List[PSCustomObject]]::new()

        Write-Host ""
        Write-Host "Invoke-RoboCopy " -ForegroundColor Cyan -NoNewline
        Write-Host "- Copying $sizeDisplay to $($ComputerName.Count) computer(s)"
        Write-Host "  ------------------------------------------------"
        Write-Host "  Source:      $resolvedSource"
        Write-Host "  Destination: $LocalizedDestination\$itemFolder"
        Write-Host "  Throttle:    $Throttle"
        Write-Host "  Timeout:     $Timeout min"
        Write-Host "  ------------------------------------------------"
        Write-Host ""

        foreach ($computer in $ComputerName) {
            $ps = [PowerShell]::Create()
            $ps.RunspacePool = $pool
            $ps.AddScript($roboCopyScript.ToString()) > $null
            $ps.AddArgument($computer) > $null
            $ps.AddArgument($resolvedSource) > $null
            $ps.AddArgument($LocalizedDestination) > $null
            $ps.AddArgument($itemFolder) > $null
            $ps.AddArgument($Force.IsPresent) > $null

            $handle = $ps.BeginInvoke()

            $runspaces.Add([PSCustomObject]@{
                Computer   = [string]$computer
                PowerShell = $ps
                Handle     = $handle
                StartTime  = [DateTime]::Now
                Completed  = $false
                TimedOut   = $false
            })
        }

        $script:currentRunspaces = $runspaces

        #endregion --- Submit Runspaces ---


        #region --- Monitoring Loop ---

        try {

            $total = $runspaces.Count
            $secondsElapsed = 0
            $isConsoleHost = $Host.Name -eq 'ConsoleHost'

            $nameWidth = [int]($runspaces.Computer | ForEach-Object { $_.Length } |
                Measure-Object -Maximum).Maximum
            $nameWidth = [math]::Max($nameWidth, 8)

            if ($isConsoleHost) {
                $progressTop = [Console]::CursorTop
                $progressEnd = $progressTop
            }

            while (@($runspaces | Where-Object { -not $_.Completed -and -not $_.TimedOut }).Count -gt 0) {

                # Mark newly completed
                foreach ($rs in $runspaces) {
                    if ($rs.Completed -or $rs.TimedOut) { continue }
                    if ($rs.Handle.IsCompleted) {
                        $rs.Completed = $true
                    }
                }

                $doneCount = @($runspaces | Where-Object { $_.Completed -or $_.TimedOut }).Count

                # Adaptive update interval
                $updateInterval = if ($secondsElapsed -lt 60) { 3 }
                                  elseif ($secondsElapsed -lt 300) { 10 }
                                  else { 30 }

                if ($isConsoleHost -and $secondsElapsed % $updateInterval -eq 0) {
                    [Console]::SetCursorPosition(0, $progressTop)

                    Write-Host ""
                    Write-Host "Waiting on " -NoNewline
                    Write-Host "RoboCopy " -ForegroundColor Magenta -NoNewline
                    Write-Host "tasks... " -NoNewline
                    Write-Host "($doneCount/$total Completed)" -NoNewline
                    Write-Host "  Timeout: $Timeout min" -ForegroundColor DarkGray -NoNewline
                    Write-Host ""
                    Write-Host ""
                    $renderLines = 3

                    $active = @($runspaces | Where-Object {
                        -not $_.Completed -and -not $_.TimedOut -and -not $_.Handle.IsCompleted
                    })
                    if ($active.Count -gt 0) {
                        $tableProps = @(
                            @{
                                Name       = 'Computer'
                                Expression = { $_.Computer }
                                Width      = [int]$nameWidth + 1
                            }
                            @{
                                Name       = 'Elapsed'
                                Expression = {
                                    $span = [DateTime]::Now - $_.StartTime
                                    if ($span.TotalMinutes -ge 1) {
                                        "{0}m {1:D2}s" -f [math]::Floor($span.TotalMinutes), $span.Seconds
                                    }
                                    else {
                                        "{0}s" -f [math]::Floor($span.TotalSeconds)
                                    }
                                }
                                Width      = 9
                            }
                            @{
                                Name       = 'Status'
                                Expression = { $PhaseTracker[$_.Computer] }
                            }
                        )
                        $sorted = $active | Sort-Object { $_.StartTime }  # earliest start = longest elapsed first
                        $tableStr = $sorted | Format-Table -Property $tableProps | Out-String
                        Write-Host $tableStr
                        $renderLines += ([regex]::Matches($tableStr, "`n")).Count + 1
                    }

                    # Clear leftover lines from previous taller render
                    $contentEnd = [Console]::CursorTop
                    $clearWidth = [Console]::BufferWidth - 1
                    for ($clr = $contentEnd; $clr -lt $progressEnd; $clr++) {
                        [Console]::SetCursorPosition(0, $clr)
                        [Console]::Write(" " * $clearWidth)
                    }
                    $progressEnd = $contentEnd
                    $progressTop = $contentEnd - $renderLines
                }
                elseif (-not $isConsoleHost -and $secondsElapsed % 5 -eq 0) {
                    $pct = if ($total -gt 0) { [math]::Floor(($doneCount / $total) * 100) } else { 0 }

                    $active = @($runspaces | Where-Object {
                        -not $_.Completed -and -not $_.TimedOut -and -not $_.Handle.IsCompleted
                    })
                    $activeNames = ($active | Select-Object -First 5 | ForEach-Object {
                        $phase = $PhaseTracker[$_.Computer]
                        if ($phase) { "$($_.Computer) [$phase]" } else { $_.Computer }
                    }) -join ', '
                    if ($active.Count -gt 5) {
                        $activeNames += " ... (+$($active.Count - 5) more)"
                    }

                    Write-Progress `
                        -Activity "Waiting on RoboCopy tasks..." `
                        -Status "$doneCount/$total Completed  |  Timeout: $Timeout min" `
                        -PercentComplete $pct `
                        -CurrentOperation "Running: $activeNames"
                }

                Start-Sleep -Seconds 1
                $secondsElapsed++


                # --- Timeout handling ---
                foreach ($rs in $runspaces) {
                    if ($rs.Completed -or $rs.TimedOut) { continue }

                    $elapsed = [DateTime]::Now - $rs.StartTime
                    if ($elapsed.TotalMinutes -lt $Timeout) { continue }

                    Write-Host "Time expired - Automatically stopping $($rs.Computer)..." -ForegroundColor Yellow
                    Stop-RoboCopyRunspaceAsync -PowerShell $rs.PowerShell -Label $rs.Computer
                    $rs.TimedOut = $true
                }
            }

            if (-not $isConsoleHost) {
                Write-Progress -Activity "Waiting on RoboCopy tasks..." -Completed
            }

            # Clear the progress area
            if ($isConsoleHost) {
                [Console]::SetCursorPosition(0, $progressTop)
                $clearWidth = [Console]::BufferWidth - 1
                for ($clr = $progressTop; $clr -lt $progressEnd; $clr++) {
                    [Console]::SetCursorPosition(0, $clr)
                    [Console]::Write(" " * $clearWidth)
                }
                [Console]::SetCursorPosition(0, $progressTop)
            }

            # Clean up phase tracker
            foreach ($rs in $runspaces) {
                $PhaseTracker.Remove($rs.Computer) > $null
            }

            Write-Host ""
            Write-Host "RoboCopy " -ForegroundColor Magenta -NoNewline
            Write-Host "tasks complete!"
            Write-Host ""


            #region --- Result Collection ---

            $results = [System.Collections.Generic.List[PSCustomObject]]::new()

            foreach ($rs in $runspaces) {
                $elapsed = [DateTime]::Now - $rs.StartTime

                # Format elapsed as friendly string for successful copies
                if ($elapsed.TotalMinutes -ge 1) {
                    $elapsedStr = "{0}m {1:D2}s" -f [math]::Floor($elapsed.TotalMinutes), $elapsed.Seconds
                }
                else {
                    $elapsedStr = "{0}s" -f [math]::Floor($elapsed.TotalSeconds)
                }

                if ($rs.TimedOut) {
                    $results.Add([PSCustomObject]@{
                        ComputerName = $rs.Computer
                        Elapsed      = ''
                        CopyComplete = $false
                        ExitCode     = $null
                        Comment      = "Exceeded $Timeout minute timeout"
                    })
                }
                elseif ($rs.Completed) {
                    $gotOutput = $false
                    try {
                        $output = $rs.PowerShell.EndInvoke($rs.Handle)
                        foreach ($item in $output) {
                            $gotOutput = $true
                            # Add Elapsed to the result object from the scriptblock
                            $item | Add-Member -NotePropertyName 'Elapsed' -NotePropertyValue $elapsedStr -Force
                            $results.Add($item)
                        }
                    }
                    catch {
                        $gotOutput = $true
                        $results.Add([PSCustomObject]@{
                            ComputerName = $rs.Computer
                            Elapsed      = ''
                            CopyComplete = $false
                            ExitCode     = $null
                            Comment      = "Task failed: $_"
                        })
                    }

                    if ($rs.PowerShell.HadErrors -and -not $gotOutput) {
                        $errMsg = ($rs.PowerShell.Streams.Error | Select-Object -First 1)
                        $results.Add([PSCustomObject]@{
                            ComputerName = $rs.Computer
                            Elapsed      = ''
                            CopyComplete = $false
                            ExitCode     = $null
                            Comment      = "Task error: $errMsg"
                        })
                    }
                }
            }

            # Sort: failures first (CopyComplete=False before True), then by ExitCode descending
            $results | Sort-Object @{Expression = 'CopyComplete'; Ascending = $true},
                                   @{Expression = 'ExitCode'; Descending = $true}

            #endregion --- Result Collection ---

            #endregion --- Monitoring Loop ---

        }

        finally {

            #region --- Cleanup ---

            foreach ($rs in $runspaces) {
                try { $rs.PowerShell.Dispose() } catch {}
            }
            try { $pool.Close() } catch {}
            try { $pool.Dispose() } catch {}
            $script:currentPool = $null
            $script:currentRunspaces = $null

            #endregion --- Cleanup ---
        }

    } # End process

} # End function


Export-ModuleMember -Function Invoke-RoboCopy
