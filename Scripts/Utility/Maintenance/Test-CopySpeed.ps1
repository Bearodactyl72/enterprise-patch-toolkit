# DOTS formatting comment

function Test-CopySpeed {
    <#
        .SYNOPSIS
            Benchmarks network copy speed between two source locations.
        .DESCRIPTION
            Copies a patch folder from two different source paths to a test
            machine and measures the elapsed time for each. Reports throughput
            in MB/s and total duration. Useful for comparing regional patch
            servers or DFS vs direct share performance.

            The test performs an SMB warm-up before timing begins so the first
            source is not penalized by connection setup overhead. Destination
            folders are cleaned up automatically after measurement.
        .PARAMETER ComputerName
            The remote machine to copy files to.
        .PARAMETER Source1
            First source UNC or local path (e.g. "\\Server1\Patches").
        .PARAMETER Source2
            Second source UNC or local path (e.g. "\\Server2\Patches").
        .PARAMETER PatchFolder
            Name of the patch subfolder to copy from each source.
        .PARAMETER Region
            Label for the test machine's region. Used in log file name.
            Default: "FSTR"
        .EXAMPLE
            Test-CopySpeed -ComputerName "PC01" -Source1 "\\SRV1\Patches" -Source2 "\\SRV2\Patches" -PatchFolder "Google_Chrome_130.0.6723.91"
        .EXAMPLE
            Test-CopySpeed -ComputerName "PC01" -Source1 "M:\Patches" -Source2 "\\DFS\Patches" -PatchFolder "Reader_25.001.21223" -Region "WEST"
        .NOTES
            Written by Skyler Werner
            Version: 2.0.0
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]
        $ComputerName,

        [Parameter(Mandatory)]
        [string]
        $Source1,

        [Parameter(Mandatory)]
        [string]
        $Source2,

        [Parameter(Mandatory)]
        [string]
        $PatchFolder,

        [Parameter()]
        [string]
        $Region = "FSTR",

        [Parameter()]
        [switch]
        $PassThru
    )

    $ComputerName = $ComputerName.Trim().ToUpper()

    # --- Helper: format bytes to human-readable size ---
    function Format-FileSize {
        param([long]$Bytes)
        if ($Bytes -ge 1GB) { return '{0:N2} GB' -f ($Bytes / 1GB) }
        if ($Bytes -ge 1MB) { return '{0:N2} MB' -f ($Bytes / 1MB) }
        if ($Bytes -ge 1KB) { return '{0:N2} KB' -f ($Bytes / 1KB) }
        return "$Bytes bytes"
    }

    # --- Helper: format timespan for display ---
    function Format-Duration {
        param([TimeSpan]$Duration)
        if ($Duration.TotalMinutes -ge 1) {
            return '{0}m {1}s' -f [int]$Duration.TotalMinutes, $Duration.Seconds
        }
        return '{0:N1}s' -f $Duration.TotalSeconds
    }

    # --- Helper: timestamped log line ---
    filter Write-Log { "$(Get-Date -Format G): $_" }

    $logDir = Join-Path $env:USERPROFILE "Desktop\Logs"
    if (-not (Test-Path $logDir)) {
        New-Item $logDir -ItemType Directory -Force | Out-Null
    }

    $dateStamp = Get-Date -Format "yyyy-MM-dd-HHmm"
    $logPath = Join-Path $logDir "SpeedTest_${ComputerName}_${Region}_${dateStamp}.log"

    $destBase = "\\$ComputerName\C`$\Temp"
    $dest1 = Join-Path $destBase "SpeedTest_1"
    $dest2 = Join-Path $destBase "SpeedTest_2"
    $source1Full = Join-Path $Source1 $PatchFolder
    $source2Full = Join-Path $Source2 $PatchFolder

    # --- Validate sources exist ---
    foreach ($src in @(@{Name = 'Source1'; Path = $source1Full}, @{Name = 'Source2'; Path = $source2Full})) {
        if (-not (Test-Path $src.Path)) {
            Write-Error "$($src.Name) path not found: $($src.Path)"
            return
        }
    }

    # --- Measure source size ---
    $sourceMeasure = Get-ChildItem $source1Full -Recurse -Force -ErrorAction SilentlyContinue |
        Where-Object { -not $_.PSIsContainer } |
        Measure-Object -Property Length -Sum
    $sourceSize = $sourceMeasure.Sum
    $sourceFileCount = $sourceMeasure.Count
    $sourceSizeDisplay = Format-FileSize $sourceSize

    # --- Connectivity check ---
    if (-not (Test-Connection $ComputerName -Count 1 -Quiet)) {
        Write-Error "Test machine $ComputerName is offline."
        return
    }

    "Test-CopySpeed starting on $ComputerName ($Region)" | Write-Log | Out-File $logPath
    "Patch folder: $PatchFolder ($sourceFileCount files, $sourceSizeDisplay)" | Write-Log | Out-File $logPath -Append

    # --- Clean destination folders from any prior run ---
    foreach ($dest in @($dest1, $dest2)) {
        if (Test-Path $dest) {
            Remove-Item $dest -Recurse -Force -ErrorAction SilentlyContinue
        }
        New-Item $dest -ItemType Directory -Force | Out-Null
    }

    # --- SMB warm-up ---
    # Write and read a small file over the UNC path to prime the SMB session
    # (authentication, share negotiation, caching). This prevents the first
    # timed copy from being penalized by connection setup overhead.
    "Warming up SMB connection..." | Write-Log | Out-File $logPath -Append
    $warmupFile = Join-Path $destBase "SpeedTest_warmup.tmp"
    try {
        [System.IO.File]::WriteAllText($warmupFile, "warmup")
        $null = [System.IO.File]::ReadAllText($warmupFile)
        Remove-Item $warmupFile -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Warning "SMB warm-up failed: $($_.Exception.Message)"
    }

    # Also warm up each source path with a small read
    foreach ($src in @($source1Full, $source2Full)) {
        $firstFile = Get-ChildItem $src -File -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $firstFile) {
            $null = [System.IO.File]::ReadAllBytes($firstFile.FullName) | Select-Object -First 1
        }
    }
    "Warm-up complete." | Write-Log | Out-File $logPath -Append

    # --- Extract friendly source labels ---
    $source1Label = ($Source1.Split('\') | Where-Object { $_ -ne '' }) | Select-Object -First 1
    $source2Label = ($Source2.Split('\') | Where-Object { $_ -ne '' }) | Select-Object -First 1

    # --- Test 1: Copy from Source1 ---
    "Copying from $source1Label ($source1Full) to $dest1" | Write-Log | Out-File $logPath -Append
    Write-Host "Copying from $source1Label..." -NoNewline

    $source1Time = Measure-Command {
        Copy-Item $source1Full -Destination $dest1 -Force -Recurse -ErrorAction Stop
    }

    Write-Host " done. $(Format-Duration $source1Time)"
    "Copy from $source1Label complete in $(Format-Duration $source1Time)" | Write-Log | Out-File $logPath -Append

    # --- Test 2: Copy from Source2 ---
    "Copying from $source2Label ($source2Full) to $dest2" | Write-Log | Out-File $logPath -Append
    Write-Host "Copying from $source2Label..." -NoNewline

    $source2Time = Measure-Command {
        Copy-Item $source2Full -Destination $dest2 -Force -Recurse -ErrorAction Stop
    }

    Write-Host " done. $(Format-Duration $source2Time)"
    "Copy from $source2Label complete in $(Format-Duration $source2Time)" | Write-Log | Out-File $logPath -Append

    # --- Calculate throughput ---
    $source1MBps = 0
    $source2MBps = 0
    if ($source1Time.TotalSeconds -gt 0) {
        $source1MBps = ($sourceSize / 1MB) / $source1Time.TotalSeconds
    }
    if ($source2Time.TotalSeconds -gt 0) {
        $source2MBps = ($sourceSize / 1MB) / $source2Time.TotalSeconds
    }

    # --- Build results ---
    $results = [PSCustomObject]@{
        ComputerName = $ComputerName
        Region       = $Region
        PatchFolder  = $PatchFolder
        PatchSize    = $sourceSizeDisplay
        FileCount    = $sourceFileCount
        Source1      = $source1Label
        Source1Time  = Format-Duration $source1Time
        Source1MBps  = '{0:N2}' -f $source1MBps
        Source2      = $source2Label
        Source2Time  = Format-Duration $source2Time
        Source2MBps  = '{0:N2}' -f $source2MBps
        Winner       = if ($source1Time -lt $source2Time) { $source1Label } else { $source2Label }
    }

    # --- Log results ---
    "" | Out-File $logPath -Append
    "--- RESULTS ---" | Write-Log | Out-File $logPath -Append
    "$source1Label : $(Format-Duration $source1Time) ($($results.Source1MBps) MB/s)" | Write-Log | Out-File $logPath -Append
    "$source2Label : $(Format-Duration $source2Time) ($($results.Source2MBps) MB/s)" | Write-Log | Out-File $logPath -Append
    "Winner: $($results.Winner)" | Write-Log | Out-File $logPath -Append
    "Speed Test complete." | Write-Log | Out-File $logPath -Append

    # --- Display results ---
    Write-Host ""
    $results | Format-List | Out-Host

    # --- Clean up destination folders ---
    "Cleaning up test folders on $ComputerName..." | Write-Log | Out-File $logPath -Append
    foreach ($dest in @($dest1, $dest2)) {
        if (Test-Path $dest) {
            Remove-Item $dest -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    "Cleanup complete." | Write-Log | Out-File $logPath -Append

    # --- Copy log to shared location ---
    $sharedLogDir = "M:\Share\VMT\Logs\$env:COMPUTERNAME"
    if (Test-Path (Split-Path $sharedLogDir)) {
        if (-not (Test-Path $sharedLogDir)) {
            New-Item $sharedLogDir -ItemType Directory -Force | Out-Null
        }
        Copy-Item $logPath -Destination $sharedLogDir -ErrorAction SilentlyContinue
    }

    Write-Host "Log saved to $logPath"
    if ($PassThru) { return $results }
}
