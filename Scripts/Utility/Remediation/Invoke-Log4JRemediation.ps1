# DOTS formatting comment

function Invoke-Log4JRemediation {
    <#
        .SYNOPSIS
            Finds and removes vulnerable Log4J JAR files from remote machines.
        .DESCRIPTION
            Recursively searches specified directories on remote machines for *log4j*.jar
            files, reads the version from the JAR manifest (without extracting), and removes
            any file that is vulnerable (v2.0.0 - v2.16.x), unsupported (< v2.0.0), or
            unversioned.

            Writes a timestamped log on each remote machine and returns structured results.
            Uses Invoke-RunspacePool for parallel execution. Supports -WhatIf for dry-run.
        .PARAMETER ComputerName
            One or more target machine names. Accepts pipeline input.
        .PARAMETER CompliantVersion
            Minimum safe Log4J version. JARs with versions below this (in the 2.x range)
            are considered vulnerable. Default: "2.17.0"
        .PARAMETER SearchPath
            Directories to search recursively on the remote machine.
            Default: C:\Users, C:\Program Files*\ForeScout\
        .PARAMETER ThrottleLimit
            Maximum concurrent runspaces. Default 50.
        .PARAMETER Timeout
            Timeout in minutes before a runspace is stopped. Default 15.
        .EXAMPLE
            $list = Get-Content "$env:USERPROFILE\Desktop\Lists\Target_Machines.txt"
            Invoke-Log4JRemediation -ComputerName $list -WhatIf
        .EXAMPLE
            Invoke-Log4JRemediation -ComputerName "PC001","PC002" -CompliantVersion "2.17.1"
        .NOTES
            Written by Skyler Werner
            Version: 2.0.0
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [string[]]
        $ComputerName,

        [Parameter()]
        [string]
        $CompliantVersion = "2.17.0",

        [Parameter()]
        [string[]]
        $SearchPath = @(
            "C:\Users",
            "C:\Program Files*\ForeScout\"
        ),

        [Parameter()]
        [ValidateRange(1, 300)]
        [int]
        $ThrottleLimit = 50,

        [Parameter()]
        [Alias("TimeoutMinutes")]
        [ValidateRange(1, 120)]
        [int]
        $Timeout = 15,

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


# --- Connectivity check ---
Write-Host ""
Write-Host "Checking for online machines..."

$pingResults = Test-ConnectionAsJob -ComputerName $targets

$onlineList  = @()
$allResults  = @()

foreach ($ping in $pingResults) {
    if ($ping.Reachable -eq $true) {
        $onlineList += $ping.ComputerName
    }
    else {
        $allResults += [PSCustomObject][ordered]@{
            ComputerName = $ping.ComputerName
            Status       = "Offline"
            FilesFound   = $null
            Vulnerable   = $null
            Removed      = $null
            Remediated   = $null
            Comment      = $null
        }
    }
}

if ($onlineList.Count -eq 0) {
    Write-Host "No online machines found." -ForegroundColor Yellow
    $allResults | Sort-Object ComputerName | Format-Table -AutoSize | Out-Host
    return
}


# --- Capture WhatIf state for passing into runspaces ---
$isWhatIf = $WhatIfPreference


# --- Thread-safe dictionary for partial results ---
$partialResults = [System.Collections.Concurrent.ConcurrentDictionary[string, hashtable]]::new()


# --- Scriptblock executed in each runspace (one per machine) ---
$remediationScriptBlock = {

    $computer         = $args[0]
    $searchPaths      = $args[1]
    $compliantVer     = $args[2]
    $dryRun           = $args[3]
    $partialData      = $args[4]

    # Save computer name immediately so timeout still yields a result
    $partialData[$computer] = @{ ComputerName = $computer }

    $PhaseTracker[$computer] = "Log4J Search"

    $remoteArgs = @($searchPaths, $compliantVer, $dryRun)

    try {
        $remoteResult = Invoke-Command -ComputerName $computer -ArgumentList $remoteArgs -ErrorAction Stop -ScriptBlock {
            param($Paths, $SafeVersion, $IsDryRun)

            # --- Setup logging (suppressed under -WhatIf so dry runs leave no footprint) ---
            $logPath = "$env:SystemRoot\Logs\Log4j-Remediation.log"
            if ($IsDryRun) {
                filter Write-Log { }
            }
            else {
                filter Write-Log { "$(Get-Date -Format G): $_" | Out-File -FilePath $logPath -Append }
                "Running Log4J Remediation Script." | Write-Log
            }

            # Load compression assembly
            [System.Reflection.Assembly]::LoadWithPartialName('System.IO.Compression') | Out-Null

            $safeVer = [version]$SafeVersion

            # --- Search for JAR files ---
            $jarFiles = @()
            foreach ($searchPath in $Paths) {
                $found = Get-ChildItem $searchPath -Include "*log4j*.jar" -Recurse -ErrorAction SilentlyContinue -Force
                if ($null -ne $found) {
                    $jarFiles += $found
                }
            }

            if ($jarFiles.Count -eq 0) {
                "No log4j JAR files found." | Write-Log
                return [PSCustomObject]@{
                    FilesFound = 0
                    Vulnerable = 0
                    Removed    = $null
                    Remediated = $true
                    Comment    = "No log4j files found"
                }
            }

            "Found $($jarFiles.Count) log4j JAR file(s)." | Write-Log

            # --- Process each JAR ---
            $vulnerableCount = 0
            $removedCount    = $null
            $failedPaths     = @()

            if (-not $IsDryRun) { $removedCount = 0 }

            foreach ($jar in $jarFiles) {
                $jarPath    = $jar.FullName
                $jarVersion = $null

                # Read version from manifest without extracting
                try {
                    $zip = [System.IO.Compression.ZipFile]::OpenRead($jarPath)
                    try {
                        $manifest = $zip.GetEntry("META-INF/MANIFEST.MF")
                        if ($null -ne $manifest) {
                            $reader  = [System.IO.StreamReader]::new($manifest.Open())
                            $content = $reader.ReadToEnd()
                            $reader.Close()

                            foreach ($line in ($content -split "`n")) {
                                if ($line -match 'Implementation-Version\s*:\s*(.+)') {
                                    $jarVersion = ($Matches[1]).Trim()
                                    break
                                }
                            }
                        }
                    }
                    finally {
                        $zip.Dispose()
                    }
                }
                catch {
                    "Failed to read manifest from: $jarPath - $_" | Write-Log
                }

                # --- Classify the JAR ---
                $classification = $null

                if ($null -ne $jarVersion -and $jarVersion.Length -gt 0) {
                    try {
                        $ver = [version]$jarVersion

                        if ($ver -ge $safeVer) {
                            $classification = "Safe"
                            "Safe log4j JAR (v$jarVersion): $jarPath" | Write-Log
                        }
                        elseif ($ver -ge [version]"2.0.0") {
                            $classification = "Vulnerable"
                            "Vulnerable log4j JAR (v$jarVersion): $jarPath" | Write-Log
                        }
                        else {
                            $classification = "Unsupported"
                            "Unsupported log4j JAR (v$jarVersion): $jarPath" | Write-Log
                        }
                    }
                    catch {
                        $classification = "Unknown"
                        "Could not parse version '$jarVersion' for: $jarPath" | Write-Log
                    }
                }
                else {
                    $classification = "Unknown"
                    "Could not determine version for: $jarPath" | Write-Log
                }

                # --- Remove if not safe ---
                if ($classification -ne "Safe") {
                    $vulnerableCount++

                    if ($IsDryRun) {
                        "WhatIf: Would remove: $jarPath" | Write-Log
                    }
                    else {
                        Remove-Item $jarPath -Force -ErrorAction SilentlyContinue
                        if (Test-Path $jarPath) {
                            "Failed to delete: $jarPath" | Write-Log
                            $failedPaths += $jarPath
                        }
                        else {
                            "Successfully deleted: $jarPath" | Write-Log
                            $removedCount++
                        }
                    }
                }
            }

            # --- Build result ---
            $remediated = $true
            $comment    = $null

            if (-not $IsDryRun -and $failedPaths.Count -gt 0) {
                $remediated = $false
                $comment = "Failed to remove: " + ($failedPaths -join ", ")
            }

            if ($IsDryRun) {
                $remediated = $null
            }

            "Remediation complete. Files: $($jarFiles.Count), Vulnerable: $vulnerableCount" | Write-Log

            [PSCustomObject]@{
                FilesFound = $jarFiles.Count
                Vulnerable = $vulnerableCount
                Removed    = $removedCount
                Remediated = $remediated
                Comment    = $comment
            }
        }

        # Build output object
        [PSCustomObject][ordered]@{
            ComputerName = $computer
            Status       = "Online"
            FilesFound   = $remoteResult.FilesFound
            Vulnerable   = $remoteResult.Vulnerable
            Removed      = $remoteResult.Removed
            Remediated   = $remoteResult.Remediated
            Comment      = $remoteResult.Comment
        }
    }
    catch {
        [PSCustomObject][ordered]@{
            ComputerName = $computer
            Status       = "Online"
            FilesFound   = $null
            Vulnerable   = $null
            Removed      = $null
            Remediated   = $null
            Comment      = "Failed: $_"
        }
    }
}


# --- Build argument sets (one per machine) ---
$argumentSets = @(
    foreach ($machine in $onlineList) {
        , @(
            $machine,
            $SearchPath,
            $CompliantVersion,
            $isWhatIf,
            $partialResults
        )
    }
)


# --- Execute via Invoke-RunspacePool ---
$runspaceParams = @{
    ScriptBlock    = $remediationScriptBlock
    ArgumentList   = $argumentSets
    ThrottleLimit  = $ThrottleLimit
    TimeoutMinutes = $Timeout
    ActivityName   = "Log4J Remediation"
}

$poolResults = Invoke-RunspacePool @runspaceParams


# --- Post-processing ---
foreach ($result in $poolResults) {
    if ($result -isnot [PSCustomObject]) { continue }

    # Results with a Status property came from our scriptblock
    if ($null -ne $result.PSObject.Properties['Status']) {
        $allResults += $result
        continue
    }

    # Incomplete results from Invoke-RunspacePool (timed out / failed)
    $allResults += [PSCustomObject][ordered]@{
        ComputerName = $result.ComputerName
        Status       = "Online"
        FilesFound   = $null
        Vulnerable   = $null
        Removed      = $null
        Remediated   = $null
        Comment      = if ($null -ne $result.Comment) { $result.Comment } else { "Task Stopped" }
    }
}


# --- Output results ---
$selectProps = @(
    "ComputerName",
    "Status",
    "FilesFound",
    "Vulnerable",
    "Removed",
    "Remediated",
    "Comment"
)

    $sorted = $allResults | Select-Object $selectProps | Sort-Object -Property @(
        @{ Expression = "Status";       Descending = $true  }
        @{ Expression = "Remediated";   Descending = $false }
        @{ Expression = "Vulnerable";   Descending = $true  }
        @{ Expression = "ComputerName"; Descending = $false }
    )

    $sorted | Format-Table -AutoSize | Out-Host
    if ($PassThru) { return $sorted }

    } # end
}
