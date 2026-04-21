# DOTS formatting comment

function Restart-Machine {
    <#
        .SYNOPSIS
            Schedules a graceful restart on remote machines with a user warning.
        .DESCRIPTION
            Issues a shutdown /r command to each remote machine with a configurable
            delay so users can save their work. Displays the estimated completion
            time after all commands have been sent.

            Use -Abort to cancel pending restarts that have not yet executed.
        .PARAMETER ComputerName
            One or more target machine names. Accepts pipeline input.
        .PARAMETER DelayMinutes
            Minutes to wait before restarting. Default 15.
        .PARAMETER Message
            Notification text displayed to the logged-in user.
        .PARAMETER Abort
            Cancel pending restarts instead of scheduling new ones. Only works
            if the restart delay has not yet elapsed.
        .EXAMPLE
            $list = Get-Content "$env:USERPROFILE\Desktop\Lists\Target_Machines.txt"
            Restart-Machine -ComputerName $list
        .EXAMPLE
            Restart-Machine -ComputerName "PC001","PC002" -DelayMinutes 5
        .EXAMPLE
            Restart-Machine -ComputerName "PC001","PC002" -Abort
        .NOTES
            Written by Skyler Werner
            Version: 2.0.0

            Shutdown flags reference:
              /r  Restart
              /m  Remote machine (\\name)
              /t  Delay in seconds
              /c  Message displayed to user (in quotes)
              /f  Force open applications to close after delay
              /a  Abort a pending shutdown (must run before delay elapses)
              /d  Reason code (p:4:1 = planned application maintenance)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [string[]]
        $ComputerName,

        [Parameter()]
        [ValidateRange(1, 120)]
        [int]
        $DelayMinutes = 15,

        [Parameter()]
        [string]
        $Message = "Your computer will restart in $DelayMinutes minutes to apply an update. Please save your files now.",

        [Parameter()]
        [switch]
        $Abort,

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

$delaySec = $DelayMinutes * 60


# --- Business hours confirmation for large runs ---
if (-not $Abort -and $targets.Count -ge 10) {
    $hour = (Get-Date).Hour
    if ($hour -ge 7 -and $hour -lt 17) {
        Write-Host ""
        Write-Host "WARNING: You are about to restart $($targets.Count) machines during business hours." -ForegroundColor Yellow
        $confirm = Read-Host "Type YES to continue"
        if ($confirm -ne "YES") {
            Write-Host "Aborted." -ForegroundColor Red
            return
        }
    }
}


# --- Send commands ---
$activity = if ($Abort) { "Cancelling Restarts" } else { "Scheduling Restarts" }

Write-Host ""

$allResults = [System.Collections.Generic.List[PSCustomObject]]::new()
$total      = $targets.Count
$current    = 0

foreach ($computer in $targets) {
    $current++
    Write-Progress -Activity $activity -Status "$current / $total" -PercentComplete ([math]::Round($current / $total * 100))

    if ($Abort) {
        & shutdown /a /m "\\$computer" 2>&1 | Out-Null
    }
    else {
        & shutdown /r /m "\\$computer" /t $delaySec /c $Message /f /d p:4:1 2>&1 | Out-Null
    }
    $exitCode = $LASTEXITCODE

    $status  = if ($exitCode -eq 0 -and $Abort) { "Cancelled" }
        elseif ($exitCode -eq 0) { "Queued" }
        else { "Failed" }
    $comment = if ($exitCode -ne 0) { "shutdown exit code $exitCode" } else { $null }

    $allResults.Add([PSCustomObject][ordered]@{
        ComputerName  = $computer
        Status        = $status
        ExitCode      = $exitCode
        Comment       = $comment
    })
}

Write-Progress -Activity $activity -Completed


# --- Estimated completion time (restart mode only) ---
if (-not $Abort) {
    $now      = Get-Date
    $estimate = $now + (New-TimeSpan -Seconds $delaySec)

    Write-Host ""
    Write-Host "Current Time:      $now"
    Write-Host "Estimated Restart: $estimate"
}

Write-Host ""


# --- Output results ---
    $sorted = $allResults | Sort-Object -Property @(
        @{ Expression = "Status";       Descending = $true  }
        @{ Expression = "ComputerName"; Descending = $false }
    )

    $sorted | Format-Table -AutoSize | Out-Host
    if ($PassThru) { return $sorted }

    } # end
}
