# DOTS formatting comment

function Repair-MachineHealth {
    <#
        .SYNOPSIS
            Runs a suite of system health repair tasks on remote machines.
        .DESCRIPTION
            Executes a full health repair sequence on each target machine in parallel:
            pending reboot check, service state verification, clock sync, SFC, DISM,
            temp file cleanup, SCCM client repair, and group policy refresh. Results
            are collected into a summary table and detailed logs are copied back to
            the local machine.

            Uses Invoke-RunspacePool for parallel execution across many machines.
        .PARAMETER ComputerName
            One or more target machine names. Accepts pipeline input.
        .PARAMETER ThrottleLimit
            Maximum concurrent runspaces. Default 25.
        .PARAMETER Timeout
            Timeout in minutes before a runspace is stopped. Default 60.
            This is intentionally high because SFC + DISM can take a while.
        .EXAMPLE
            $list = Get-Content "$env:USERPROFILE\Desktop\Lists\Target_Machines.txt"
            Repair-MachineHealth -ComputerName $list
        .EXAMPLE
            Repair-MachineHealth -ComputerName "PC001" -Timeout 90
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
        [ValidateRange(1, 300)]
        [int]
        $ThrottleLimit = 25,

        [Parameter()]
        [Alias("TimeoutMinutes")]
        [ValidateRange(1, 180)]
        [int]
        $Timeout = 60,

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


# --- Remote log path (used by remote block and Copy-Log) ---
$remoteLogFolder = "C:\Windows\Logs\Repair"
$scriptLabel     = "Repair-MachineHealth"


# --- Connectivity check ---
Write-Host ""
Write-Host "Checking for online machines..."

    $pingResults = Test-ConnectionAsJob -ComputerName $targets

    $onlineList = [System.Collections.Generic.List[string]]::new()
    $allResults = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($ping in $pingResults) {
    if ($ping.Reachable -eq $true) {
        $onlineList.Add($ping.ComputerName)
    }
    else {
        $allResults.Add([PSCustomObject][ordered]@{
            ComputerName  = $ping.ComputerName
            Status        = "Offline"
            PendingReboot = $null
            DiskFreeGB    = $null
            SfcScan       = $null
            DismCommands  = $null
            TempCleanup   = $null
            SccmRepair    = $null
            Comment       = $null
        })
    }
}

if ($onlineList.Count -eq 0) {
    Write-Host "No online machines found." -ForegroundColor Yellow
    $allResults | Sort-Object -Property @(
        @{ Expression = "Status";       Descending = $true  }
        @{ Expression = "ComputerName"; Descending = $false }
    ) | Format-Table -AutoSize | Out-Host
    return
}


# --- Capture WhatIf state for passing into runspaces ---
$isWhatIf = $WhatIfPreference


# --- Scriptblock executed in each runspace (one per machine) ---
$repairHealthBlock = {

    $computer = $args[0]
    $dryRun   = $args[1]

    $PhaseTracker[$computer] = "Connecting"

    if ($dryRun) {
        $PhaseTracker[$computer] = "WhatIf"
        return [PSCustomObject][ordered]@{
            ComputerName  = $computer
            Status        = "Online"
            PendingReboot = 'WhatIf'
            DiskFreeGB    = $null
            SfcScan       = 'WhatIf'
            DismCommands  = 'WhatIf'
            TempCleanup   = 'WhatIf'
            SccmRepair    = 'WhatIf'
            Comment       = '8 repair steps (see above)'
        }
    }

    try {
        $PhaseTracker[$computer] = "Running Repairs (SFC, DISM, Cleanup, SCCM)"
        $remoteResult = Invoke-Command -ComputerName $computer -ErrorAction Stop -ScriptBlock {

            $date     = Get-Date -Format "yyyy-MM-dd-HHmm"
            $logDir   = "C:\Windows\Logs\Repair"
            if (-not (Test-Path $logDir)) { New-Item $logDir -ItemType Directory -Force | Out-Null }
            $logFile  = "$logDir\Repair-MachineHealth-$env:COMPUTERNAME-$date.txt"

            function Log {
                param ([string]$Message)
                $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                "$timestamp - $Message" | Tee-Object -FilePath $logFile -Append
            }

            $cn = $env:COMPUTERNAME
            Log "===== System Repair Script Started on $cn ====="

            # --- Pending Reboot Check ---
            Log "Checking for pending reboot..."
            $pendingReboot = "No"
            $rebootKeys = @(
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending",
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired",
                "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager"
            )
            if ((Test-Path $rebootKeys[0]) -or (Test-Path $rebootKeys[1])) {
                $pendingReboot = "Yes"
            }
            else {
                $pfro = Get-ItemProperty $rebootKeys[2] -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
                if ($null -ne $pfro.PendingFileRenameOperations) {
                    $pendingReboot = "Yes"
                }
            }
            Log "Pending reboot: $pendingReboot"

            # --- Service State Check ---
            Log "Checking critical service states..."
            $servicesToCheck = @("wuauserv", "BITS", "CryptSvc")
            foreach ($svcName in $servicesToCheck) {
                $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
                if ($null -eq $svc) {
                    Log "  $svcName - not found"
                    continue
                }
                Log "  $svcName - Status: $($svc.Status), StartType: $($svc.StartType)"
                if ($svc.StartType -eq 'Disabled') {
                    try {
                        Set-Service -Name $svcName -StartupType Manual
                        Log "  $svcName - changed StartType from Disabled to Manual"
                    }
                    catch { Log "  $svcName - failed to change StartType: $_" }
                }
                if ($svc.Status -ne 'Running') {
                    try {
                        Start-Service -Name $svcName -ErrorAction Stop
                        Log "  $svcName - started"
                    }
                    catch { Log "  $svcName - failed to start: $_" }
                }
            }

            # --- Clock Sync ---
            Log "Syncing system clock..."
            try {
                & w32tm /resync /force 2>&1 | Tee-Object -FilePath $logFile -Append
                Log "Clock sync completed."
            }
            catch {
                Log "Clock sync failed: $_"
            }

            # --- SFC Scan ---
            Log "Running SFC Scan..."
            $sfcResult = "Completed"
            try {
                & sfc /scannow | Tee-Object -FilePath $logFile -Append
                Log "SFC Scan completed."
            }
            catch {
                Log "SFC Scan failed: $_"
                $sfcResult = "Failed"
            }

            # --- DISM Commands ---
            Log "Running DISM commands..."
            $dismResult  = "Completed"
            $dismCommands = @(
                @("DISM", "/Online", "/Cleanup-Image", "/ScanHealth"),
                @("DISM", "/Online", "/Cleanup-Image", "/CheckHealth"),
                @("DISM", "/Online", "/Cleanup-Image", "/RestoreHealth"),
                @("DISM", "/Online", "/Cleanup-Image", "/StartComponentCleanup")
            )

            foreach ($cmd in $dismCommands) {
                $label = $cmd -join " "
                Log "Running $label..."
                try {
                    & $cmd[0] $cmd[1] $cmd[2] $cmd[3] 2>&1 | Tee-Object -FilePath $logFile -Append
                    Log "$label completed."
                }
                catch {
                    Log "$label failed: $_"
                    $dismResult = "Failed"
                }
            }

            # --- Temp File Cleanup ---
            Log "Running temp file cleanup..."
            $tempResult = "Completed"
            $cleanupPaths = @(
                "$env:SystemRoot\Temp\*",
                "$env:SystemRoot\Logs\CBS\*.log",
                "C:\Users\*\AppData\Local\Temp\*",
                "C:\ProgramData\Microsoft\Windows\WER\*"
            )

            $totalRemoved = 0
            foreach ($path in $cleanupPaths) {
                try {
                    $items = @(Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue)
                    if ($items.Count -gt 0) {
                        $items | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                        $totalRemoved += $items.Count
                    }
                }
                catch {
                    Log "Cleanup warning for ${path}: $_"
                }
            }

            # Windows Update download cache
            try {
                Log "Clearing Windows Update download cache..."
                Stop-Service wuauserv -Force -ErrorAction SilentlyContinue
                $wuItems = @(Get-ChildItem "C:\Windows\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue)
                if ($wuItems.Count -gt 0) {
                    $wuItems | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                    $totalRemoved += $wuItems.Count
                }
                Start-Service wuauserv -ErrorAction SilentlyContinue
            }
            catch {
                Log "Windows Update cache cleanup warning: $_"
                Start-Service wuauserv -ErrorAction SilentlyContinue
            }

            # Recycle Bin
            try {
                Clear-RecycleBin -DriveLetter C -Force -ErrorAction SilentlyContinue
                Log "Recycle Bin cleared."
            }
            catch {
                Log "Recycle Bin cleanup warning: $_"
            }

            Log "Temp cleanup finished. $totalRemoved items processed."

            # --- CCMRepair ---
            $ccmRepairPath = "C:\Windows\CCM\ccmrepair.exe"
            $ccmResult = "Not Found"
            if (Test-Path $ccmRepairPath) {
                Log "Running CCMRepair..."
                try {
                    Start-Process -FilePath $ccmRepairPath -Wait -NoNewWindow
                    Log "CCMRepair completed."
                    $ccmResult = "Completed"
                }
                catch {
                    Log "CCMRepair failed: $_"
                    $ccmResult = "Failed"
                }
            }
            else {
                Log "CCMRepair not found at $ccmRepairPath"
            }

            # --- Group Policy Refresh ---
            Log "Running gpupdate /force..."
            try {
                & gpupdate /force 2>&1 | Tee-Object -FilePath $logFile -Append
                Log "Group Policy refresh completed."
            }
            catch {
                Log "Group Policy refresh failed: $_"
            }

            # --- Disk Free Space ---
            $diskFreeGB = $null
            try {
                $disk = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'"
                $diskFreeGB = [math]::Round($disk.FreeSpace / 1GB, 1)
                Log "Disk free space on C: $diskFreeGB GB"
            }
            catch {
                Log "Failed to read disk space: $_"
            }

            Log "===== System Repair Script Completed on $cn ====="

            [PSCustomObject]@{
                PendingReboot = $pendingReboot
                DiskFreeGB    = $diskFreeGB
                SfcScan       = $sfcResult
                DismCommands  = $dismResult
                TempCleanup   = $tempResult
                SccmRepair    = $ccmResult
            }
        }

        $PhaseTracker[$computer] = "Complete"

        [PSCustomObject][ordered]@{
            ComputerName  = $computer
            Status        = "Online"
            PendingReboot = $remoteResult.PendingReboot
            DiskFreeGB    = $remoteResult.DiskFreeGB
            SfcScan       = $remoteResult.SfcScan
            DismCommands  = $remoteResult.DismCommands
            TempCleanup   = $remoteResult.TempCleanup
            SccmRepair    = $remoteResult.SccmRepair
            Comment       = $null
        }
    }
    catch {
        $PhaseTracker[$computer] = "Error"

        [PSCustomObject][ordered]@{
            ComputerName  = $computer
            Status        = "Online"
            PendingReboot = $null
            DiskFreeGB    = $null
            SfcScan       = $null
            DismCommands  = $null
            TempCleanup   = $null
            SccmRepair    = $null
            Comment       = "Failed: $_"
        }
    }
}


# --- Build argument sets (one per machine) ---
$argumentSets = @(
    foreach ($machine in $onlineList) {
        , @($machine, $isWhatIf)
    }
)


# --- Execute via Invoke-RunspacePool ---
if ($isWhatIf) {
    Write-Host ""
    Write-Host "WhatIf: The following repair steps would run on each machine:" -ForegroundColor Yellow
    Write-Host "  1. Check pending reboot (registry keys)"
    Write-Host "  2. Verify/start services: wuauserv, BITS, CryptSvc"
    Write-Host "  3. w32tm /resync /force (clock sync)"
    Write-Host "  4. sfc /scannow"
    Write-Host "  5. DISM /ScanHealth, /CheckHealth, /RestoreHealth, /StartComponentCleanup"
    Write-Host "  6. Temp cleanup (Windows\Temp, CBS logs, user Temp, WER, WU cache, Recycle Bin)"
    Write-Host "  7. ccmrepair.exe (if SCCM client present)"
    Write-Host "  8. gpupdate /force"
    Write-Host ""
    Write-Host "Targets: $($onlineList.Count) machine(s)"
    Write-Host ""
}

$runspaceParams = @{
    ScriptBlock    = $repairHealthBlock
    ArgumentList   = $argumentSets
    ThrottleLimit  = $ThrottleLimit
    TimeoutMinutes = $Timeout
    ActivityName   = "Repair Machine Health"
}

$poolResults = Invoke-RunspacePool @runspaceParams


# --- Post-processing ---
foreach ($result in $poolResults) {
    if ($result -isnot [PSCustomObject]) { continue }

    if ($null -ne $result.PSObject.Properties['Status']) {
        $allResults.Add($result)
        continue
    }

    $allResults.Add([PSCustomObject][ordered]@{
        ComputerName  = $result.ComputerName
        Status        = "Online"
        PendingReboot = $null
        DiskFreeGB    = $null
        SfcScan       = $null
        DismCommands  = $null
        TempCleanup   = $null
        SccmRepair    = $null
        Comment       = if ($null -ne $result.Comment) { $result.Comment } else { "Task Stopped" }
    })
}


# --- Output results ---
$allResults | Select-Object ComputerName, Status, PendingReboot, DiskFreeGB, SfcScan, DismCommands, TempCleanup, SccmRepair, Comment |
    Sort-Object -Property @(
        @{ Expression = "Status";       Descending = $true  }
        @{ Expression = "ComputerName"; Descending = $false }
    ) | Format-Table -AutoSize | Out-Host


    # --- Retrieve logs from remote machines ---
    $onlineNames = @($allResults | Where-Object { $_.Status -eq "Online" } | ForEach-Object { $_.ComputerName })
    if ($onlineNames.Count -gt 0) {
        Copy-Log -ComputerName $onlineNames -ScriptName $scriptLabel -RemoteLogPath $remoteLogFolder
    }

    if ($PassThru) { return $allResults }

    } # end
}
