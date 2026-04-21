# DOTS formatting comment

function Get-LoggedInUser {
    <#
        .SYNOPSIS
            Queries remote machines for active and disconnected user sessions.
        .DESCRIPTION
            Runs quser.exe on each target machine via Invoke-Command and parses
            the output into structured objects. Uses Invoke-RunspacePool for
            concurrent execution and Test-ConnectionAsJob for pre-filtering
            offline machines.

            Each result includes the computer name, session type (console/RDP),
            session state (Active/Disc), idle time, and logon time.
        .PARAMETER ComputerName
            One or more computer names to query. Accepts pipeline input.
        .PARAMETER ThrottleLimit
            Maximum concurrent machines to query. Default: 50
        .PARAMETER TimeoutMinutes
            Minutes before a machine's query task is auto-stopped. Default: 5
        .EXAMPLE
            $list = Get-Content "$env:USERPROFILE\Desktop\Lists\Target_Machines.txt"
            Get-LoggedInUser -ComputerName $list
        .EXAMPLE
            Get-LoggedInUser -ComputerName "PC01","PC02" | Format-Table -AutoSize
        .EXAMPLE
            Get-LoggedInUser -ComputerName $list | Where-Object { $_.State -eq 'Active' }
        .NOTES
            Written by Skyler Werner
            Date: 2026/03/23
            Version: 2.0.0
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [string[]]
        $ComputerName,

        [Parameter()]
        [ValidateRange(1, 300)]
        [int]
        $ThrottleLimit = 50,

        [Parameter()]
        [ValidateRange(1, 120)]
        [int]
        $TimeoutMinutes = 5,

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
        $pingResults  = Test-ConnectionAsJob -ComputerName $targets
        $online       = @($pingResults | Where-Object { $_.Reachable } | Select-Object -ExpandProperty ComputerName)
        $offline      = @($pingResults | Where-Object { -not $_.Reachable } | Select-Object -ExpandProperty ComputerName)

        if ($online.Count -eq 0) {
            Write-Warning "No machines responded to ping."
            $offlineOnly = @(foreach ($pc in $offline) {
                [PSCustomObject]@{
                    ComputerName = $pc
                    Status       = 'Offline'
                    Username     = $null
                    SessionName  = $null
                    Id           = $null
                    State        = $null
                    IdleTime     = $null
                    LogonTime    = $null
                    Comment      = ''
                }
            })
            $offlineOnly | Format-Table -AutoSize | Out-Host
            if ($PassThru) { return $offlineOnly }
            return
        }

        # --- Build argument list ---
        $argList = $online | ForEach-Object { ,@($_) }

        # --- Remote scriptblock ---
        $scriptBlock = {
            $computer = $args[0]

            # --- Parse quser idle time into a human-readable string ---
            function ConvertTo-IdleString {
                param([string]$Raw)
                $Raw = $Raw.Trim()
                if ($Raw -eq '.' -or $Raw -eq 'none' -or $Raw -eq '') {
                    return 'None'
                }
                # days+HH:MM format
                if ($Raw -match '^(\d+)\+(\d+):(\d+)$') {
                    $d = [int]$Matches[1]; $h = [int]$Matches[2]; $m = [int]$Matches[3]
                    $parts = @()
                    if ($d -gt 0) { $parts += "${d}d" }
                    if ($h -gt 0) { $parts += "${h}h" }
                    if ($m -gt 0) { $parts += "${m}m" }
                    if ($parts.Count -eq 0) { return 'None' }
                    return $parts -join ' '
                }
                # H:MM or HH:MM format
                if ($Raw -match '^(\d+):(\d+)$') {
                    $h = [int]$Matches[1]; $m = [int]$Matches[2]
                    $parts = @()
                    if ($h -gt 0) { $parts += "${h}h" }
                    if ($m -gt 0) { $parts += "${m}m" }
                    if ($parts.Count -eq 0) { return 'None' }
                    return $parts -join ' '
                }
                # Bare number = minutes
                if ($Raw -match '^\d+$') {
                    $m = [int]$Raw
                    if ($m -eq 0) { return 'None' }
                    return "${m}m"
                }
                return $Raw
            }

            try {
                $quserOutput = Invoke-Command -ComputerName $computer -ScriptBlock {
                    quser.exe 2>&1
                } -ErrorAction Stop

                # quser returns error text when no user is logged in
                if ($null -eq $quserOutput -or @($quserOutput).Count -eq 0) {
                    [PSCustomObject]@{
                        ComputerName = $computer
                        Status       = 'Online'
                        Username     = $null
                        SessionName  = $null
                        Id           = $null
                        State        = 'No sessions'
                        IdleTime     = $null
                        LogonTime    = $null
                        Comment      = ''
                    }
                    return
                }

                # Convert to string array for reliable parsing
                $lines = @($quserOutput | ForEach-Object { "$_" })

                # Check if the output is an error message rather than session data
                $firstLine = $lines[0].Trim()
                if ($firstLine -match 'No User exists' -or $firstLine -match 'Error') {
                    [PSCustomObject]@{
                        ComputerName = $computer
                        Status       = 'Online'
                        Username     = $null
                        SessionName  = $null
                        Id           = $null
                        State        = 'No sessions'
                        IdleTime     = $null
                        LogonTime    = $null
                        Comment      = $firstLine
                    }
                    return
                }

                # Parse the header to find column positions
                # quser header: " USERNAME  SESSIONNAME  ID  STATE  IDLE TIME  LOGON TIME"
                # The ">" prefix on the current user's line shifts things by 1 char
                $header = $lines[0]

                $colUser    = $header.IndexOf('USERNAME')
                $colSession = $header.IndexOf('SESSIONNAME')
                $colId      = $header.IndexOf('ID')
                $colState   = $header.IndexOf('STATE')
                $colIdle    = $header.IndexOf('IDLE TIME')
                $colLogon   = $header.IndexOf('LOGON TIME')

                # If we couldn't find header columns, fall back to fixed positions
                if ($colUser -lt 0)    { $colUser    = 1 }
                if ($colSession -lt 0) { $colSession = 23 }
                if ($colId -lt 0)      { $colId      = 42 }
                if ($colState -lt 0)   { $colState   = 46 }
                if ($colIdle -lt 0)    { $colIdle    = 54 }
                if ($colLogon -lt 0)   { $colLogon   = 65 }

                $sessionLines = $lines | Select-Object -Skip 1

                foreach ($line in $sessionLines) {
                    if ([string]::IsNullOrWhiteSpace($line)) { continue }

                    # Pad the line so Substring doesn't throw on short lines
                    $padded = $line.PadRight($colLogon + 20)

                    $username    = $padded.Substring($colUser, $colSession - $colUser).Trim()
                    $sessionName = $padded.Substring($colSession, $colId - $colSession).Trim()
                    $id          = $padded.Substring($colId, $colState - $colId).Trim()
                    $state       = $padded.Substring($colState, $colIdle - $colState).Trim()
                    $idleRaw     = $padded.Substring($colIdle, $colLogon - $colIdle).Trim()
                    $logonTime   = $padded.Substring($colLogon).Trim()
                    $idleTime    = ConvertTo-IdleString $idleRaw

                    # The ">" prefix on the current user can push USERNAME
                    # into the leading space. Clean it up.
                    $username = $username.TrimStart('>', ' ')

                    [PSCustomObject]@{
                        ComputerName = $computer
                        Status       = 'Online'
                        Username     = $username
                        SessionName  = $sessionName
                        Id           = $id
                        State        = $state
                        IdleTime     = $idleTime
                        LogonTime    = $logonTime
                        Comment      = ''
                    }
                }
            }
            catch {
                $errMsg = ($_.Exception.Message).Trim()

                # "No User exists" is a normal quser error, not a real failure
                if ($errMsg -match 'No User exists') {
                    [PSCustomObject]@{
                        ComputerName = $computer
                        Status       = 'Online'
                        Username     = $null
                        SessionName  = $null
                        Id           = $null
                        State        = 'No sessions'
                        IdleTime     = $null
                        LogonTime    = $null
                        Comment      = ''
                    }
                }
                else {
                    [PSCustomObject]@{
                        ComputerName = $computer
                        Status       = 'Online'
                        Username     = $null
                        SessionName  = $null
                        Id           = $null
                        State        = $null
                        IdleTime     = $null
                        LogonTime    = $null
                        Comment      = $errMsg
                    }
                }
            }
        }

        # --- Execute via RunspacePool ---
        $results = @(Invoke-RunspacePool `
            -ScriptBlock    $scriptBlock `
            -ArgumentList   $argList `
            -ThrottleLimit  $ThrottleLimit `
            -TimeoutMinutes $TimeoutMinutes `
            -ActivityName   "Get-LoggedInUser"
        )

        # --- Normalize timed-out/failed results ---
        $normalizedResults = @()
        foreach ($r in $results) {
            if ($r -isnot [PSCustomObject]) { continue }
            if ($null -ne $r.PSObject.Properties['Username']) {
                $normalizedResults += $r
            }
            else {
                $normalizedResults += [PSCustomObject]@{
                    ComputerName = $r.ComputerName
                    Status       = 'Online'
                    Username     = $null
                    SessionName  = $null
                    Id           = $null
                    State        = $null
                    IdleTime     = $null
                    LogonTime    = $null
                    Comment      = if ($null -ne $r.Comment) { $r.Comment } else { 'Task Stopped' }
                }
            }
        }
        $results = $normalizedResults

        # --- Add offline machines to results ---
        foreach ($pc in $offline) {
            $results += [PSCustomObject]@{
                ComputerName = $pc
                Status       = 'Offline'
                Username     = $null
                SessionName  = $null
                Id           = $null
                State        = $null
                IdleTime     = $null
                LogonTime    = $null
                Comment      = ''
            }
        }


        # --- Summary output ---
        Write-Host ""
        Write-Host "========================================"
        Write-Host "  Get-LoggedInUser -- Results"
        Write-Host "========================================"
        Write-Host ""

        $withSessions    = @($results | Where-Object { $_.Username })
        $noSessions      = @($results | Where-Object { $_.Status -eq 'Online' -and -not $_.Username -and $_.State -eq 'No sessions' })
        $errored         = @($results | Where-Object { $_.Status -eq 'Online' -and -not $_.Username -and $_.State -ne 'No sessions' })
        $offlineResults  = @($results | Where-Object { $_.Status -eq 'Offline' })
        $uniqueOnline    = @($results | Where-Object { $_.Status -eq 'Online' } | Select-Object -ExpandProperty ComputerName -Unique)

        # Session detail table
        $sorted = $results | Sort-Object @(
            @{ Expression = 'Status';    Descending = $true  }
            @{ Expression = 'State';     Descending = $true  }
            @{ Expression = 'Username';  Descending = $false }
            @{ Expression = 'IdleTime';  Descending = $true  }
        )

        $sorted | Format-Table ComputerName, Status, Username, SessionName,
            Id, State, IdleTime, LogonTime, Comment -AutoSize | Out-Host

        # Totals
        Write-Host "========================================"
        Write-Host "  Totals"
        Write-Host "========================================"
        Write-Host ""
        Write-Host "  Machines targeted:    $($targets.Count)"
        Write-Host "    Online:             $($uniqueOnline.Count)"
        Write-Host "    Offline:            $($offlineResults.Count)"
        Write-Host ""
        Write-Host "    With sessions:      $($withSessions.Count) session(s) across $(
            @($withSessions | Select-Object -ExpandProperty ComputerName -Unique).Count
        ) machine(s)"
        Write-Host "    No sessions:        $($noSessions.Count)"
        if ($errored.Count -gt 0) {
            Write-Host "    Errors:             $($errored.Count)"
        }
        Write-Host ""

        if ($PassThru) { return $sorted }
    }
}
