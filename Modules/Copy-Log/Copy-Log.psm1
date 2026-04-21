# DOTS formatting comment

<#
    .SYNOPSIS
        Copies log files from remote machines to a local folder on the admin's desktop.
    .DESCRIPTION
        Copies log files from remote machines to a centralized folder on the admin's
        desktop. By default expects logs named {ScriptName}-{ComputerName}-{DateTime}.txt
        on the remote machine. Use -Filter to override the filename pattern for scripts
        that use a different naming convention. Creates the destination folder if it
        does not exist.

        Written by Skyler Werner
        Date: 2026/03/26
        Version 1.1.0
    .PARAMETER ComputerName
        One or more remote machine names to copy logs from.
    .PARAMETER ScriptName
        The script name prefix used in log filenames. Builds the default filter
        pattern: {ScriptName}-{ComputerName}-*.txt
    .PARAMETER Filter
        A custom filename wildcard pattern (e.g. '*_Adobe_Reader_*.log'). When
        provided, this is used instead of the ScriptName-based pattern.
    .PARAMETER RemoteLogPath
        The folder path on the remote machine containing the log files.
    .PARAMETER DestinationFolder
        The folder name on the admin's desktop to copy logs into. Defaults to
        'Repair-Logs'.
#>

function Copy-Log {
    [CmdletBinding(DefaultParameterSetName = 'ByScriptName')]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string[]]$ComputerName,

        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = 'ByScriptName')]
        [string]$ScriptName,

        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = 'ByScriptName')]
        [Parameter(Mandatory = $true, ParameterSetName = 'ByFilter')]
        [string]$RemoteLogPath,

        [Parameter(Position = 3, ParameterSetName = 'ByScriptName')]
        [Parameter(ParameterSetName = 'ByFilter')]
        [string]$DestinationFolder = "Repair-Logs",

        [Parameter(Mandatory = $true, ParameterSetName = 'ByFilter')]
        [string]$Filter
    )

    # --- Build and verify local destination ---
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $localDest   = Join-Path $desktopPath $DestinationFolder

    if (-not (Test-Path $localDest)) {
        New-Item -Path $localDest -ItemType Directory -Force | Out-Null
        Write-Host "Created folder: $localDest"
    }

    # --- Convert remote path to UNC share format ---
    # e.g. C:\Windows\Logs\Repair -> Windows\Logs\Repair
    $remoteSuffix = $RemoteLogPath -replace '^[A-Za-z]:\\', ''
    $driveLetter  = ($RemoteLogPath -replace ':\\.*', '').ToUpper()

    $copied = 0
    $failed = 0

    foreach ($computer in $ComputerName) {
        $uncPath = "\\$computer\$driveLetter`$\$remoteSuffix"

        # Build file filter -- custom pattern or default ScriptName-Computer convention
        if ($Filter) {
            $fileFilter = $Filter
        }
        else {
            $fileFilter = "$ScriptName-$computer-*.txt"
        }

        try {
            if (-not (Test-Path $uncPath)) {
                Write-Warning "Path not found: $uncPath"
                $failed++
                continue
            }

            $logFiles = @(Get-ChildItem -Path $uncPath -Filter $fileFilter -ErrorAction Stop)

            if ($logFiles.Count -eq 0) {
                Write-Warning "No logs matching '$fileFilter' found on $computer"
                $failed++
                continue
            }

            foreach ($logFile in $logFiles) {
                Copy-Item -Path $logFile.FullName -Destination $localDest -Force -ErrorAction Stop
                $copied++
            }
        }
        catch {
            Write-Warning "Failed to copy logs from $computer -- $($_.Exception.Message)"
            $failed++
        }
    }

    # --- Summary ---
    Write-Host ""
    Write-Host "Log retrieval complete: $copied copied, $failed failed"
    Write-Host "Logs saved to: $localDest"
}

Export-ModuleMember Copy-Log
