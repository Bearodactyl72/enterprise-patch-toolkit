# DOTS formatting comment

<#
    .SYNOPSIS
        Executes a scriptblock across multiple targets using a RunspacePool for concurrent execution.
    .DESCRIPTION
        Generic concurrency engine that replaces all *AsJob modules. Creates a RunspacePool, submits
        one runspace per argument set, monitors progress, and collects results. By default, timed-out
        runspaces are automatically stopped. Use -ConfirmTimeout for interactive timeout prompts.
        Includes stale pool cleanup for interrupted runs (Ctrl+C safety).

        Written by Skyler Werner
        Date: 2026/03/04
        Version 2.0.1
#>

# Module-scoped variables for stale pool detection after Ctrl+C
$script:currentPool = $null
$script:currentRunspaces = $null


function Stop-RunspaceAsync {
    <#
        .SYNOPSIS
            Sends an async stop signal to a PowerShell instance and waits up to $TimeoutSeconds.
            If the pipeline doesn't respond, control returns anyway so Dispose() can clean up later.
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
        # BeginStop can throw if the pipeline is already in a broken state -- move on
    }
}


function Invoke-RunspacePool {
    [CmdletBinding()]
    param(
        # Scriptblock to execute in each runspace. Receives arguments via $args.
        [Parameter(Mandatory, Position = 0)]
        [ScriptBlock]
        $ScriptBlock,

        # Array of argument arrays. Each element is an array of arguments for one runspace invocation.
        # The first element of each inner array should be a computer name (used for display).
        [Parameter(Mandatory, Position = 1)]
        [Array]
        $ArgumentList,

        # Maximum concurrent runspaces
        [Parameter(Position = 2)]
        [ValidateRange(1, 300)]
        [Int32]
        $ThrottleLimit = 50,

        # Minutes before timeout handling triggers (default 1 hour, max 2 hours)
        [Parameter(Position = 3)]
        [ValidateRange(1, 240)]
        [Int32]
        $TimeoutMinutes = 60,

        # Prompt the user to confirm or dismiss each timed-out runspace instead of auto-stopping
        [Parameter()]
        [Alias("CT")]
        [Switch]
        $ConfirmTimeout,

        # Label for progress display (e.g. "Google Chrome")
        [Parameter()]
        [String]
        $ActivityName = "Runspace"
    )


    process {

        #region --- ArgumentList Guard ---
        # PS 5.1 can unwrap a single-element array during parameter binding (especially
        # via splatting), turning @( ,@(args) ) into just @(args). Detect and re-wrap.
        if ($ArgumentList.Count -gt 0 -and $ArgumentList[0] -isnot [Array]) {
            $ArgumentList = , $ArgumentList
        }
        #endregion --- ArgumentList Guard ---


        #region --- Stale Pool Cleanup ---

        if ($null -ne $script:currentPool) {
            Write-Warning "Cleaning up stale runspace pool from a previous interrupted run..."
            try {
                if ($null -ne $script:currentRunspaces) {
                    foreach ($rs in $script:currentRunspaces) {
                        Stop-RunspaceAsync -PowerShell $rs.PowerShell -TimeoutSeconds 5
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



        #region --- Helpers ---

        # Phase sort order -- higher number = further along in the pipeline
        $phaseOrder = @{
            "Pinging"       = 1
            "DNS Lookup"    = 2
            "Version Check" = 3
            "Copying Files" = 4
            "Patching"      = 5
            "Verifying"     = 6
        }

        function Format-Elapsed {
            param([TimeSpan]$Span)
            if ($Span.TotalMinutes -ge 1) {
                return "{0}m {1:D2}s" -f [math]::Floor($Span.TotalMinutes), $Span.Seconds
            }
            return "{0}s" -f [math]::Floor($Span.TotalSeconds)
        }

        #endregion --- Helpers ---



        #region --- Pool Creation ---

        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

        # Inject a synchronized hashtable into every runspace for live phase tracking.
        # Scriptblocks can write $PhaseTracker[$Computer] = "Phase..." to report status.
        $PhaseTracker = [Hashtable]::Synchronized(@{})
        $phaseEntry = [System.Management.Automation.Runspaces.SessionStateVariableEntry]::new(
            'PhaseTracker', $PhaseTracker, 'Synchronized hashtable for live phase tracking'
        )
        $iss.Variables.Add($phaseEntry)

        # Inject a synchronized hashtable for optional caller-defined status text.
        # Scriptblocks can write $StatusMessage['Text'] = "..." to append info to the
        # progress status line (e.g. "Avg Success: ~25s"). The module displays it but
        # does not interpret it -- callers own the content.
        $StatusMessage = [Hashtable]::Synchronized(@{})
        $statusEntry = [System.Management.Automation.Runspaces.SessionStateVariableEntry]::new(
            'StatusMessage', $StatusMessage, 'Synchronized hashtable for caller-defined status text'
        )
        $iss.Variables.Add($statusEntry)

        # Do NOT pass $Host to the pool. Sharing $Host between the pool and
        # the session corrupts the Out-Default formatter on pool disposal,
        # killing all output for the rest of the session. Our scriptblocks
        # are non-interactive workers that use $PhaseTracker for status --
        # they never call Write-Host, Read-Host, or any host-dependent API.
        $pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool($iss)
        $pool.SetMaxRunspaces($ThrottleLimit) > $null
        $pool.Open()
        $script:currentPool = $pool

        #endregion --- Pool Creation ---



        #region --- Runspace Submission ---

        $runspaces = [System.Collections.Generic.List[PSCustomObject]]::new()
        $total = $ArgumentList.Count
        # Wrap in array so nested function can mutate by reference
        $submitState = @{ Index = 0 }
        $scriptString = $ScriptBlock.ToString()

        # Drip-feed helper: submit new runspaces up to $ThrottleLimit in-flight.
        # Called once before the monitor loop (initial batch) and once per loop
        # iteration to refill slots as machines complete.
        function Submit-NextBatch {
            while ($submitState.Index -lt $total) {
                $inFlight = @($runspaces | Where-Object {
                    -not $_.Completed -and -not $_.TimedOut
                }).Count
                if ($inFlight -ge $ThrottleLimit) { break }

                $argSet = $ArgumentList[$submitState.Index]
                $submitState.Index++

                $ps = [PowerShell]::Create()
                $ps.RunspacePool = $pool

                # ScriptBlocks cannot cross runspace boundaries; pass as string
                $ps.AddScript($scriptString) > $null

                # Add each argument individually via AddArgument
                foreach ($arg in $argSet) {
                    $ps.AddArgument($arg) > $null
                }

                $handle = $ps.BeginInvoke()

                $runspaces.Add([PSCustomObject]@{
                    Computer   = [string]$argSet[0]
                    PowerShell = $ps
                    Handle     = $handle
                    StartTime  = [DateTime]::Now
                    Completed  = $false
                    TimedOut   = $false
                    Skipped    = $false
                })
            }
        }

        # Submit initial batch
        Submit-NextBatch

        $script:currentRunspaces = $runspaces

        #endregion --- Runspace Submission ---



        try {

            #region --- Monitoring Loop ---

            $skipAll = $false
            $secondsElapsed = 0
            $isConsoleHost = $Host.Name -eq 'ConsoleHost'

            # Fixed column width from ALL targets so alignment stays stable as machines complete
            $nameWidth = [int]($ArgumentList | ForEach-Object { ([string]$_[0]).Length } |
                Measure-Object -Maximum).Maximum
            $nameWidth = [math]::Max($nameWidth, 8)  # at least as wide as "Computer"

            if ($isConsoleHost) {
                $progressTop = [Console]::CursorTop
                $progressEnd = $progressTop
            }

            # IMPORTANT: Wrap Where-Object results with @() to guarantee array context.
            # In PS 5.1, Where-Object returning a single object yields a scalar whose
            # .Count property is $null, causing the while condition to evaluate as $false
            # and skip the monitoring loop entirely for single-runspace invocations.

            while (
                $submitState.Index -lt $total -or
                @($runspaces | Where-Object { -not $_.Completed -and -not $_.TimedOut }).Count -gt 0
            ) {

                # Mark newly completed runspaces
                foreach ($rs in $runspaces) {
                    if ($rs.Completed -or $rs.TimedOut) { continue }
                    if ($rs.Handle.IsCompleted) {
                        $rs.Completed = $true
                    }
                }

                # Refill slots as machines complete
                Submit-NextBatch
                $script:currentRunspaces = $runspaces

                $doneCount = @($runspaces | Where-Object { $_.Completed -or $_.TimedOut }).Count

                # Progress display -- 5s for first 5 min, then 10s
                $updateInterval = if ($secondsElapsed -lt 300) { 5 } else { 10 }

                if ($isConsoleHost -and $secondsElapsed % $updateInterval -eq 0) {
                    # --- Drift detection ---
                    # External output (Write-Host from runspaces, user clicks, or
                    # scroll) can move the cursor from where the last render left it.
                    # Detect and correct before rendering.
                    $cursorNow = [Console]::CursorTop
                    if ($cursorNow -gt $progressEnd) {
                        # External output pushed cursor past our area - clear the
                        # stale PUT and re-anchor below the new output.
                        $cw = [Console]::BufferWidth - 1
                        for ($clrDrift = $progressTop; $clrDrift -le $progressEnd; $clrDrift++) {
                            try {
                                [Console]::SetCursorPosition(0, $clrDrift)
                                [Console]::Write(' ' * $cw)
                            } catch { break }
                        }
                        $progressTop = $cursorNow
                        $progressEnd = $cursorNow
                    }
                    elseif ($cursorNow -lt $progressEnd) {
                        # Cursor moved up (user click or buffer scroll) - shift our
                        # tracked region by the same delta so the PUT stays in place.
                        $delta = $progressEnd - $cursorNow
                        $progressTop = [math]::Max(0, $progressTop - $delta)
                        $progressEnd = $cursorNow
                    }

                    [Console]::CursorVisible = $false
                    [Console]::SetCursorPosition(0, $progressTop)

                    $padWidth = [Console]::BufferWidth - 1

                    # Queued = not-yet-submitted + submitted-but-not-executing
                    # Active = currently running (has a PhaseTracker entry)
                    $pending = @($runspaces | Where-Object {
                        -not $_.Completed -and -not $_.TimedOut -and -not $_.Handle.IsCompleted
                    })
                    $activeCount    = @($pending | Where-Object { $PhaseTracker[$_.Computer] }).Count
                    $notSubmitted   = $total - $submitState.Index
                    $queuedCount    = ($pending.Count - $activeCount) + $notSubmitted
                    $completedCount = @($runspaces | Where-Object { $_.Completed }).Count
                    $failedCount    = @($runspaces | Where-Object { $_.TimedOut }).Count

                    # --- Line 1: default text, magenta software name ---
                    Write-Host (' ' * $padWidth)
                    $line1 = "Waiting on $ActivityName tasks...    Timeout: $TimeoutMinutes min    (Progress: $doneCount/$total)"
                    Write-Host "Waiting on " -NoNewline
                    Write-Host "$ActivityName " -ForegroundColor Magenta -NoNewline
                    Write-Host "tasks...    Timeout: $TimeoutMinutes min    " -NoNewline
                    Write-Host "(Progress: $doneCount/$total)".PadRight($padWidth - $line1.Length + "(Progress: $doneCount/$total)".Length)

                    # --- Runspace counts (verbose only) ---
                    $line2 = "Queued: $queuedCount  |  Active: $activeCount  |  Completed: $completedCount  |  Failed: $failedCount"
                    if ($StatusMessage['Text']) {
                        $line2 += "  |  $($StatusMessage['Text'])"
                    }
                    Write-Verbose $line2

                    Write-Host (' ' * $padWidth)

                    if ($pending.Count -gt 0) {
                        $tableProps = @(
                            @{
                                Name       = 'Computer'
                                Expression = { $_.Computer }
                                Width      = [int]$nameWidth + 1
                            }
                            @{
                                Name       = 'Elapsed'
                                Expression = { Format-Elapsed -Span ([DateTime]::Now - $_.StartTime) }
                                Width      = 8
                            }
                            @{
                                Name       = 'Status'
                                Expression = { $PhaseTracker[$_.Computer] }
                            }
                        )
                        $sorted = @($pending | Sort-Object @(
                            @{ Expression = {
                                $phase = $PhaseTracker[$_.Computer]
                                if ($phase) { $phaseOrder[$phase] } else { 0 }
                            }; Descending = $true }
                            @{ Expression = { $_.StartTime }; Ascending = $true }
                        ))

                        # --- Viewport cap ---
                        # Limit visible rows so the table fits within the terminal.
                        # Overhead: 2 header lines + 2 table header/divider + 1 summary + 2 pad
                        $maxRows = [Console]::WindowHeight - 7
                        if ($maxRows -lt 3) { $maxRows = 3 }

                        $hiddenCount = 0
                        if ($sorted.Count -gt $maxRows) {
                            $hiddenCount = $sorted.Count - $maxRows
                            $sorted = @($sorted | Select-Object -First $maxRows)
                        }

                        $tableLines = (($sorted | Format-Table -Property $tableProps | Out-String).TrimEnd()) -split "`n"
                        foreach ($tl in $tableLines) {
                            Write-Host ($tl.TrimEnd().PadRight($padWidth))
                        }

                        if ($hiddenCount -gt 0) {
                            Write-Host ("... +$hiddenCount more active".PadRight($padWidth)) -ForegroundColor DarkGray
                        }
                    }

                    # Clear leftover lines from a previous taller render
                    $contentEnd = [Console]::CursorTop
                    $clearWidth = [Console]::BufferWidth - 1
                    for ($clr = $contentEnd; $clr -lt $progressEnd; $clr++) {
                        [Console]::SetCursorPosition(0, $clr)
                        [Console]::Write(" " * $clearWidth)
                    }
                    # Park cursor at content end for accurate drift detection
                    [Console]::SetCursorPosition(0, $contentEnd)
                    $progressEnd = $contentEnd
                    [Console]::CursorVisible = $true
                }
                elseif (-not $isConsoleHost -and $secondsElapsed % 5 -eq 0) {
                    # ISE / Other: Write-Progress fallback every ~5 seconds
                    $pct = if ($total -gt 0) { [math]::Floor(($doneCount / $total) * 100) } else { 0 }

                    $pending = @($runspaces | Where-Object {
                        -not $_.Completed -and -not $_.TimedOut -and -not $_.Handle.IsCompleted
                    })
                    $activeCount    = @($pending | Where-Object { $PhaseTracker[$_.Computer] }).Count
                    $notSubmitted   = $total - $submitState.Index
                    $queuedCount    = ($pending.Count - $activeCount) + $notSubmitted
                    $completedCount = @($runspaces | Where-Object { $_.Completed }).Count
                    $failedCount    = @($runspaces | Where-Object { $_.TimedOut }).Count

                    $activeNames = ($pending | Select-Object -First 5 | ForEach-Object {
                        $phase = $PhaseTracker[$_.Computer]
                        if ($phase) { "$($_.Computer) [$phase]" } else { $_.Computer }
                    }) -join ', '
                    if ($pending.Count -gt 5) {
                        $activeNames += " ... (+$($pending.Count - 5) more)"
                    }

                    $statusParts = @()
                    $statusParts += "Queued: $queuedCount"
                    $statusParts += "Active: $activeCount"
                    $statusParts += "Completed: $completedCount"
                    $statusParts += "Failed: $failedCount"
                    if ($StatusMessage['Text']) { $statusParts += $StatusMessage['Text'] }
                    Write-Verbose ($statusParts -join '  |  ')

                    Write-Progress `
                        -Activity "Waiting on $ActivityName tasks..." `
                        -Status "Progress: $doneCount/$total  |  Timeout: $TimeoutMinutes min" `
                        -PercentComplete $pct `
                        -CurrentOperation "      Running: $activeNames"
                }

                Start-Sleep -Seconds 1
                $secondsElapsed++


                # Check for timeouts on active runspaces
                # Collect all newly timed-out runspaces first, then stop
                # them concurrently so total wait is ~15s, not 15s * N.
                $timedOutBatch = @(foreach ($rs in $runspaces) {
                    if ($rs.Completed -or $rs.TimedOut) { continue }
                    if ($rs.Skipped -or $skipAll) { continue }

                    $elapsed = [DateTime]::Now - $rs.StartTime
                    if ($elapsed.TotalMinutes -lt $TimeoutMinutes) { continue }
                    $rs
                })

                if ($timedOutBatch.Count -gt 0 -and -not $ConfirmTimeout.IsPresent) {
                    # Fire all stop signals concurrently
                    $stopHandles = @(foreach ($rs in $timedOutBatch) {
                        Write-Host "Time expired - Automatically stopping $($rs.Computer)..." -ForegroundColor Yellow
                        try {
                            [PSCustomObject]@{
                                Runspace    = $rs
                                AsyncResult = $rs.PowerShell.BeginStop($null, $null)
                            }
                        }
                        catch {
                            $rs.TimedOut = $true
                            $null
                        }
                    })

                    # Wait for each with a timeout
                    foreach ($sh in $stopHandles) {
                        if ($null -eq $sh) { continue }
                        try {
                            $stopped = $sh.AsyncResult.AsyncWaitHandle.WaitOne(
                                [TimeSpan]::FromSeconds(15)
                            )
                            if (-not $stopped) {
                                Write-Host "  Force-terminating $($sh.Runspace.Computer) (not responding to stop signal)..." -ForegroundColor Red
                            }
                        }
                        catch {}
                        $sh.Runspace.TimedOut = $true
                    }
                }

                # Interactive timeout handling (only when -ConfirmTimeout is specified)
                foreach ($rs in $timedOutBatch) {
                    if ($rs.TimedOut) { continue }
                    if (-not $ConfirmTimeout.IsPresent) { continue }

                    # Interactive timeout prompt (only when -ConfirmTimeout is specified)
                    $elapsed = [DateTime]::Now - $rs.StartTime
                    $mins = [math]::Round($elapsed.TotalMinutes, 1)
                    Write-Host ""
                    Write-Host "$($rs.Computer) is taking an unusually long time ($mins Minutes). Cancel the task?"
                    Write-Host ""

                    do {
                        $choice = Read-Host (
                            "[Y] Yes   " +
                            "[A] Yes for ALL machines   " +
                            "[N] No, don't ask me again for THIS machine   " +
                            "[D] No, don't ask me again for ANY machines   " +
                            "[R] Refresh"
                        )
                    }
                    until ($choice -match '^[YANDR]$')

                    switch ($choice) {
                        "Y" {
                            Write-Host "Stopping $($rs.Computer)..."
                            Stop-RunspaceAsync -PowerShell $rs.PowerShell -Label $rs.Computer
                            $rs.TimedOut = $true
                        }
                        "A" {
                            Write-Host "Stopping all running tasks..."

                            # Fire all stop signals concurrently first
                            $stopHandles = foreach ($activeRS in @($runspaces | Where-Object { -not $_.Completed -and -not $_.TimedOut })) {
                                try {
                                    [PSCustomObject]@{
                                        Runspace    = $activeRS
                                        AsyncResult = $activeRS.PowerShell.BeginStop($null, $null)
                                    }
                                }
                                catch {
                                    $activeRS.TimedOut = $true
                                    $null
                                }
                            }

                            # Wait for each with a timeout
                            foreach ($sh in $stopHandles) {
                                if ($null -eq $sh) { continue }
                                try {
                                    $stopped = $sh.AsyncResult.AsyncWaitHandle.WaitOne(
                                        [TimeSpan]::FromSeconds(15)
                                    )
                                    if (-not $stopped) {
                                        Write-Host "  Force-terminating $($sh.Runspace.Computer)..." -ForegroundColor Red
                                    }
                                }
                                catch {}
                                $sh.Runspace.TimedOut = $true
                            }
                        }
                        "N" {
                            $rs.Skipped = $true
                            Start-Sleep 1
                        }
                        "D" {
                            $skipAll = $true
                        }
                        "R" {
                            Write-Host "Refreshing..." -ForegroundColor Gray
                            Start-Sleep 3
                        }
                    }
                }
            }

            if (-not $isConsoleHost) {
                Write-Progress -Activity "Waiting on $ActivityName tasks..." -Completed
            }

            # Clear the progress area so stale table entries don't linger on screen
            if ($isConsoleHost) {
                # Adjust for any drift since the last render
                $cursorNow = [Console]::CursorTop
                if ($cursorNow -gt $progressEnd) {
                    # External output appeared below PUT - clear old PUT area
                    $cw = [Console]::BufferWidth - 1
                    for ($i = $progressTop; $i -le $progressEnd; $i++) {
                        try { [Console]::SetCursorPosition(0, $i); [Console]::Write(' ' * $cw) } catch { break }
                    }
                    [Console]::SetCursorPosition(0, $cursorNow)
                }
                else {
                    if ($cursorNow -lt $progressEnd) {
                        $delta = $progressEnd - $cursorNow
                        $progressTop = [math]::Max(0, $progressTop - $delta)
                        $progressEnd = $cursorNow
                    }
                    $clearWidth = [Console]::BufferWidth - 1
                    for ($clr = $progressTop; $clr -lt $progressEnd; $clr++) {
                        try {
                            [Console]::SetCursorPosition(0, $clr)
                            [Console]::Write(" " * $clearWidth)
                        } catch { break }
                    }
                    [Console]::SetCursorPosition(0, $progressTop)
                }
            }

            # Clean up phase tracker entries for completed machines
            foreach ($rs in $runspaces) {
                $PhaseTracker.Remove($rs.Computer) > $null
            }

            $completedCount = @($runspaces | Where-Object { $_.Completed }).Count
            $failedCount    = @($runspaces | Where-Object { $_.TimedOut }).Count

            Write-Host ""
            Write-Host "$ActivityName " -ForegroundColor Magenta -NoNewline
            $completeLine = "tasks complete!  $completedCount Completed  $failedCount Failed"
            if ($StatusMessage['Text']) {
                $completeLine += "  |  $($StatusMessage['Text'])"
            }
            Write-Host $completeLine
            Write-Host ""

            #endregion --- Monitoring Loop ---



            #region --- Result Collection ---

            # Emit results directly to the pipeline. Each object is written individually
            # so PS does not wrap them in a collection on the way out.

            foreach ($rs in $runspaces) {
                if ($rs.TimedOut) {
                    [PSCustomObject]@{
                        ComputerName = $rs.Computer
                        Status       = "Online"
                        Comment      = "Task Stopped"
                    }
                }
                elseif ($rs.Completed) {
                    $gotOutput = $false
                    try {
                        $output = $rs.PowerShell.EndInvoke($rs.Handle)

                        foreach ($item in $output) {
                            $gotOutput = $true
                            $item   # emit directly to pipeline
                        }
                    }
                    catch {
                        $gotOutput = $true
                        [PSCustomObject]@{
                            ComputerName = $rs.Computer
                            Comment      = "Task Failed: $_"
                        }
                    }

                    # Check for errors that didn't surface as terminating exceptions
                    if ($rs.PowerShell.HadErrors -and -not $gotOutput) {
                        $gotOutput = $true
                        $errMsg = ($rs.PowerShell.Streams.Error | Select-Object -First 1)
                        [PSCustomObject]@{
                            ComputerName = $rs.Computer
                            Comment      = "Task Error: $errMsg"
                        }
                    }

                    # Safety net: if a runspace completed but produced no output and
                    # no errors, emit a minimal result so the machine is never silently
                    # lost from the results. This can happen when an exception inside a
                    # catch block terminates the scriptblock before any output is written.
                    if (-not $gotOutput) {
                        [PSCustomObject]@{
                            ComputerName = $rs.Computer
                            Comment      = "Task produced no output (possible internal error)"
                        }
                    }
                }
            }

            #endregion --- Result Collection ---
        }

        finally {

            #region --- Cleanup ---

            # Separate responsive from stuck runspaces so Dispose() on a stuck
            # pipeline doesn't block the function for minutes.
            $stuckRS = [System.Collections.Generic.List[PSCustomObject]]::new()
            foreach ($rs in $runspaces) {
                $state = $rs.PowerShell.InvocationStateInfo.State
                if ($state -eq 'Completed' -or $state -eq 'Stopped' -or $state -eq 'Failed') {
                    try { $rs.PowerShell.Dispose() } catch {}
                }
                else {
                    $stuckRS.Add($rs)
                }
            }

            if ($stuckRS.Count -gt 0) {
                # Stop stuck runspaces synchronously before disposal.
                # NEVER use ThreadPool::QueueUserWorkItem to dispose a
                # host-bound RunspacePool -- disposing $Host on a background
                # thread crashes the PowerShell process in PS 5.1.
                Write-Host "Cleaning up $($stuckRS.Count) stuck task(s)..." -ForegroundColor DarkGray

                # Fire all stop signals concurrently
                $stopHandles = @(foreach ($rs in $stuckRS) {
                    try {
                        [PSCustomObject]@{
                            Runspace    = $rs
                            AsyncResult = $rs.PowerShell.BeginStop($null, $null)
                        }
                    }
                    catch { $null }
                })

                # Wait up to 10 seconds for each to respond
                foreach ($sh in $stopHandles) {
                    if ($null -eq $sh) { continue }
                    try {
                        $sh.AsyncResult.AsyncWaitHandle.WaitOne(
                            [TimeSpan]::FromSeconds(10)
                        ) > $null
                    }
                    catch {}
                }

                # Dispose individual PowerShell instances first
                foreach ($rs in $stuckRS) {
                    try { $rs.PowerShell.Dispose() } catch {}
                }
            }

            try { $pool.Close() } catch {}
            try { $pool.Dispose() } catch {}

            $script:currentPool = $null
            $script:currentRunspaces = $null

            #endregion --- Cleanup ---
        }

    } # End process

} # End function


Export-ModuleMember -Function Invoke-RunspacePool
