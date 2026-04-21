# DOTS formatting comment

function Repair-WindowsUpdate {
    <#
        .SYNOPSIS
            Performs aggressive Windows Update agent repairs on remote machines.
        .DESCRIPTION
            Resets the Windows Update agent on each target machine in parallel. This
            is a more aggressive repair than Repair-MachineHealth and should be used
            on machines that are confirmed stuck and not accepting patches.

            Repair sequence per machine:
              1. Secure channel repair (domain trust)
              2. Stop WU-related services (wuauserv, BITS, CryptSvc, msiserver)
              3. Rename SoftwareDistribution and catroot2 (forces rebuild)
              4. Re-register Windows Update DLLs
              5. Restart WU-related services
              6. Reset WSUS client identity (forces re-registration with WSUS)
              7. Trigger update detection cycle

            Uses Invoke-RunspacePool for parallel execution across many machines.
        .PARAMETER ComputerName
            One or more target machine names. Accepts pipeline input.
        .PARAMETER ThrottleLimit
            Maximum concurrent runspaces. Default 25.
        .PARAMETER Timeout
            Timeout in minutes before a runspace is stopped. Default 30.
        .EXAMPLE
            $list = Get-Content "$env:USERPROFILE\Desktop\Lists\Target_Machines.txt"
            Repair-WindowsUpdate -ComputerName $list
        .EXAMPLE
            Repair-WindowsUpdate -ComputerName "PC001","PC002"
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
        [ValidateRange(1, 120)]
        [int]
        $Timeout = 30,

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
$scriptLabel     = "Repair-WindowsUpdate"


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
            SecureChannel = $null
            WUAgentReset  = $null
            WsusReset     = $null
            DetectCycle   = $null
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
$repairWUBlock = {

    $computer = $args[0]
    $dryRun   = $args[1]

    $PhaseTracker[$computer] = "Connecting"

    if ($dryRun) {
        $PhaseTracker[$computer] = "WhatIf"
        return [PSCustomObject][ordered]@{
            ComputerName  = $computer
            Status        = "Online"
            SecureChannel = 'WhatIf'
            WUAgentReset  = 'WhatIf'
            WsusReset     = 'WhatIf'
            DetectCycle   = 'WhatIf'
            Comment       = '7 repair steps (see above)'
        }
    }

    try {
        $PhaseTracker[$computer] = "Repairing WU Agent"
        $remoteResult = Invoke-Command -ComputerName $computer -ErrorAction Stop -ScriptBlock {

            $date    = Get-Date -Format "yyyy-MM-dd-HHmm"
            $logDir  = "C:\Windows\Logs\Repair"
            if (-not (Test-Path $logDir)) { New-Item $logDir -ItemType Directory -Force | Out-Null }
            $logFile = "$logDir\Repair-WindowsUpdate-$env:COMPUTERNAME-$date.txt"

            function Log {
                param ([string]$Message)
                $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                "$timestamp - $Message" | Tee-Object -FilePath $logFile -Append
            }

            $cn = $env:COMPUTERNAME
            Log "===== Windows Update Repair Started on $cn ====="

            # --- Secure Channel Repair ---
            Log "Testing secure channel..."
            $scResult = "OK"
            try {
                $scTest = Test-ComputerSecureChannel
                if ($scTest) {
                    Log "Secure channel is healthy."
                }
                else {
                    Log "Secure channel is broken. Attempting repair..."
                    $repaired = Test-ComputerSecureChannel -Repair
                    if ($repaired) {
                        Log "Secure channel repaired."
                        $scResult = "Repaired"
                    }
                    else {
                        Log "Secure channel repair failed."
                        $scResult = "Failed"
                    }
                }
            }
            catch {
                Log "Secure channel test failed: $_"
                $scResult = "Failed"
            }

            # --- Stop WU Services ---
            Log "Stopping Windows Update services..."
            $wuServices = @("wuauserv", "BITS", "CryptSvc", "msiserver")
            foreach ($svcName in $wuServices) {
                try {
                    Stop-Service -Name $svcName -Force -ErrorAction SilentlyContinue
                    Log "  Stopped $svcName"
                }
                catch {
                    Log "  Warning: could not stop ${svcName}: $_"
                }
            }

            # --- Rename SoftwareDistribution and catroot2 ---
            Log "Renaming SoftwareDistribution and catroot2..."
            $wuResetResult = "Completed"
            $timestamp = Get-Date -Format "yyyyMMdd-HHmm"

            $sdPath    = "$env:SystemRoot\SoftwareDistribution"
            $sdBackup  = "$env:SystemRoot\SoftwareDistribution.bak.$timestamp"
            $crPath    = "$env:SystemRoot\System32\catroot2"
            $crBackup  = "$env:SystemRoot\System32\catroot2.bak.$timestamp"

            try {
                if (Test-Path $sdPath) {
                    Rename-Item -Path $sdPath -NewName "SoftwareDistribution.bak.$timestamp" -Force
                    Log "  Renamed SoftwareDistribution to SoftwareDistribution.bak.$timestamp"
                }
            }
            catch {
                Log "  Failed to rename SoftwareDistribution: $_"
                $wuResetResult = "Failed"
            }

            try {
                if (Test-Path $crPath) {
                    Rename-Item -Path $crPath -NewName "catroot2.bak.$timestamp" -Force
                    Log "  Renamed catroot2 to catroot2.bak.$timestamp"
                }
            }
            catch {
                Log "  Failed to rename catroot2: $_"
                $wuResetResult = "Failed"
            }

            # --- Re-register WU DLLs ---
            Log "Re-registering Windows Update DLLs..."
            $dlls = @(
                "atl.dll", "urlmon.dll", "mshtml.dll", "shdocvw.dll",
                "browseui.dll", "jscript.dll", "vbscript.dll", "scrrun.dll",
                "msxml.dll", "msxml3.dll", "msxml6.dll", "actxprxy.dll",
                "softpub.dll", "wintrust.dll", "dssenh.dll", "rsaenh.dll",
                "gpkcsp.dll", "sccbase.dll", "slbcsp.dll", "cryptdlg.dll",
                "oleaut32.dll", "ole32.dll", "shell32.dll", "initpki.dll",
                "wuapi.dll", "wuaueng.dll", "wuaueng1.dll", "wucltui.dll",
                "wups.dll", "wups2.dll", "wuweb.dll", "qmgr.dll",
                "qmgrprxy.dll", "wucltux.dll", "muweb.dll", "wuwebv.dll"
            )

            foreach ($dll in $dlls) {
                $dllPath = "$env:SystemRoot\System32\$dll"
                if (Test-Path $dllPath) {
                    & regsvr32.exe /s $dllPath 2>&1 | Out-Null
                }
            }
            Log "DLL re-registration completed."

            # --- Restart WU Services ---
            Log "Restarting Windows Update services..."
            foreach ($svcName in $wuServices) {
                try {
                    Start-Service -Name $svcName -ErrorAction SilentlyContinue
                    Log "  Started $svcName"
                }
                catch {
                    Log "  Warning: could not start ${svcName}: $_"
                }
            }

            # --- WSUS Client Reset ---
            Log "Resetting WSUS client identity..."
            $wsusResult = "Completed"
            $wsusKeyPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate"
            try {
                $susId = Get-ItemProperty -Path $wsusKeyPath -Name SusClientId -ErrorAction SilentlyContinue
                if ($null -ne $susId.SusClientId) {
                    Remove-ItemProperty -Path $wsusKeyPath -Name SusClientId -Force
                    Remove-ItemProperty -Path $wsusKeyPath -Name SusClientIdValidation -Force -ErrorAction SilentlyContinue
                    Log "  Removed SusClientId and SusClientIdValidation"
                }
                else {
                    Log "  No SusClientId found (not WSUS-managed or already clean)"
                    $wsusResult = "N/A"
                }
            }
            catch {
                Log "  Failed to reset WSUS identity: $_"
                $wsusResult = "Failed"
            }

            # --- Trigger Detection Cycle ---
            Log "Triggering update detection cycle..."
            $detectResult = "Completed"
            try {
                & wuauclt /resetauthorization /detectnow 2>&1 | Out-Null
                Log "  wuauclt /resetauthorization /detectnow completed."

                # Also try the newer UsoClient if available
                $usoPath = "$env:SystemRoot\System32\UsoClient.exe"
                if (Test-Path $usoPath) {
                    & UsoClient StartScan 2>&1 | Out-Null
                    Log "  UsoClient StartScan completed."
                }
            }
            catch {
                Log "  Detection cycle failed: $_"
                $detectResult = "Failed"
            }

            Log "===== Windows Update Repair Completed on $cn ====="

            [PSCustomObject]@{
                SecureChannel = $scResult
                WUAgentReset  = $wuResetResult
                WsusReset     = $wsusResult
                DetectCycle   = $detectResult
            }
        }

        $PhaseTracker[$computer] = "Complete"

        [PSCustomObject][ordered]@{
            ComputerName  = $computer
            Status        = "Online"
            SecureChannel = $remoteResult.SecureChannel
            WUAgentReset  = $remoteResult.WUAgentReset
            WsusReset     = $remoteResult.WsusReset
            DetectCycle   = $remoteResult.DetectCycle
            Comment       = $null
        }
    }
    catch {
        $PhaseTracker[$computer] = "Error"

        [PSCustomObject][ordered]@{
            ComputerName  = $computer
            Status        = "Online"
            SecureChannel = $null
            WUAgentReset  = $null
            WsusReset     = $null
            DetectCycle   = $null
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
    Write-Host "  1. Test-ComputerSecureChannel (repair if broken)"
    Write-Host "  2. Stop services: wuauserv, BITS, CryptSvc, msiserver"
    Write-Host "  3. Rename SoftwareDistribution and catroot2"
    Write-Host "  4. Re-register 35 Windows Update DLLs via regsvr32"
    Write-Host "  5. Restart WU services"
    Write-Host "  6. Remove WSUS SusClientId registry values"
    Write-Host "  7. wuauclt /resetauthorization /detectnow + UsoClient StartScan"
    Write-Host ""
    Write-Host "Targets: $($onlineList.Count) machine(s)"
    Write-Host ""
}

$runspaceParams = @{
    ScriptBlock    = $repairWUBlock
    ArgumentList   = $argumentSets
    ThrottleLimit  = $ThrottleLimit
    TimeoutMinutes = $Timeout
    ActivityName   = "Repair Windows Update"
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
        SecureChannel = $null
        WUAgentReset  = $null
        WsusReset     = $null
        DetectCycle   = $null
        Comment       = if ($null -ne $result.Comment) { $result.Comment } else { "Task Stopped" }
    })
}


# --- Output results ---
$allResults | Select-Object ComputerName, Status, SecureChannel, WUAgentReset, WsusReset, DetectCycle, Comment |
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
