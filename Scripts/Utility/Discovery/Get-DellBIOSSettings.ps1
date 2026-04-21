# DOTS formatting comment

function Get-DellBIOSSettings {
    <#
        .SYNOPSIS
            Queries Dell BIOS configuration from remote machines.
        .DESCRIPTION
            Copies the DellBIOSProvider module to remote machines, imports it,
            and reads all BIOS settings via the DellSMBios PSDrive. Results are
            aggregated across machines and exported to CSV.

            Uses Invoke-RunspacePool for concurrent execution and
            Test-ConnectionAsJob for pre-filtering offline machines.
        .PARAMETER ComputerName
            One or more computer names to query. Accepts pipeline input.
        .PARAMETER ModuleSourcePath
            Local or UNC path to the DellBIOSProvider module folder to copy
            to each remote machine. Defaults to the Patching share location.
        .PARAMETER Mode
            How to aggregate BIOS values across machines.
            'Append'  - collect all unique values per setting (default)
            'Remove'  - mark settings that differ across machines with '-'
        .PARAMETER ThrottleLimit
            Maximum concurrent machines to query. Default: 25
        .PARAMETER TimeoutMinutes
            Minutes before a machine's task is auto-stopped. Default: 15
        .EXAMPLE
            $list = Get-Content "$env:USERPROFILE\Desktop\Lists\Dell_BIOS.txt"
            Get-DellBIOSSettings -ComputerName $list
        .EXAMPLE
            Get-DellBIOSSettings -ComputerName "PC01","PC02" -Mode Remove
        .NOTES
            Written by Skyler Werner
            Date: 2026/03/27
            Version: 2.0.0
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [string[]]
        $ComputerName,

        [Parameter()]
        [string]
        $ModuleSourcePath = "M:\Share\VMT\Patches\DellBIOSProvider",

        [Parameter()]
        [ValidateSet('Append', 'Remove')]
        [string]
        $Mode = 'Append',

        [Parameter()]
        [ValidateRange(1, 300)]
        [int]
        $ThrottleLimit = 25,

        [Parameter()]
        [ValidateRange(1, 120)]
        [int]
        $TimeoutMinutes = 15,

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

        # --- Validate module source ---
        if (-not (Test-Path $ModuleSourcePath)) {
            Write-Error "DellBIOSProvider module not found at '$ModuleSourcePath'. Use -ModuleSourcePath to specify the correct location."
            return
        }

        # --- Sanitize input ---
        $targets = @(Format-ComputerList $collectedNames -ToUpper)
        if ($targets.Count -eq 0) {
            Write-Warning "No valid computer names provided."
            return
        }

        # --- Ping check ---
        Write-Host "Checking for online machines..."
        $pingResults = Test-ConnectionAsJob -ComputerName $targets
        $online      = @($pingResults | Where-Object { $_.Reachable } | Select-Object -ExpandProperty ComputerName)
        $offline     = @($pingResults | Where-Object { -not $_.Reachable } | Select-Object -ExpandProperty ComputerName)

        $offlineResults = @()
        foreach ($pc in $offline) {
            Write-Host "$pc Offline" -ForegroundColor Red
            $offlineResults += [PSCustomObject]@{
                ComputerName      = $pc
                Status            = 'Offline'
                Comment           = 'Offline'
                BIOSConfiguration = $null
            }
        }

        if ($online.Count -eq 0) {
            Write-Warning "No machines responded to ping."
            $offlineResults | Format-Table ComputerName, Status, Comment -AutoSize | Out-Host
            if ($PassThru) { return $offlineResults }
            return
        }

        # --- Build argument list for runspace pool ---
        $moduleDest = "C:\Program Files\WindowsPowerShell\Modules\DellBIOSProvider"
        $argList = $online | ForEach-Object { ,@($_, $ModuleSourcePath, $moduleDest) }

        # --- Runspace scriptblock ---
        $scriptBlock = {
            $computer         = $args[0]
            $moduleSource     = [string]$args[1]
            $moduleDestRemote = [string]$args[2]

            $result = [PSCustomObject]@{
                ComputerName      = $computer
                Status            = 'Online'
                Comment           = ''
                BIOSConfiguration = $null
            }

            # --- Copy module to remote machine via robocopy ---
            $uncDest = "\\$computer\C$\Program Files\WindowsPowerShell\Modules\DellBIOSProvider"
            try {
                $roboArgs = @(
                    "`"$moduleSource`""
                    "`"$uncDest`""
                    '/MIR'
                    '/R:2'
                    '/W:2'
                    '/NP'
                    '/NFL'
                    '/NDL'
                    '/NJH'
                    '/NJS'
                )
                $roboProcess = Start-Process 'robocopy.exe' -ArgumentList $roboArgs -Wait -PassThru -NoNewWindow
                # Robocopy exit codes 0-7 are success/info
                if ($roboProcess.ExitCode -gt 7) {
                    $result.Comment = "Module copy failed (robocopy exit $($roboProcess.ExitCode))"
                    return $result
                }
            }
            catch {
                $result.Comment = "Module copy failed: $($_.Exception.Message)"
                return $result
            }

            # --- Run BIOS query remotely ---
            try {
                $remoteResult = Invoke-Command -ComputerName $computer -ScriptBlock {
                    Set-ExecutionPolicy Bypass -Scope Process -Force
                    Import-Module -Name "DellBIOSProvider" -ErrorAction Stop

                    $biosData = @(Get-ChildItem -Path DellSMBios:\* -Recurse -ErrorAction SilentlyContinue |
                        Select-Object PSChildName, CurrentValue)

                    # Capture warnings (e.g. Supported Battery)
                    $warningLog = Get-ChildItem -Path DellSMBios:\* -Recurse 3>&1 2>&1
                    $warnings = @($warningLog | Where-Object { $_ -match "Supported Battery" })

                    Remove-Module -Name "DellBIOSProvider" -ErrorAction SilentlyContinue

                    [PSCustomObject]@{
                        BIOSData = $biosData
                        Warnings = $warnings
                    }
                } -ErrorAction Stop

                $result.BIOSConfiguration = $remoteResult.BIOSData
                if ($remoteResult.Warnings.Count -gt 0) {
                    $result.Comment = ($remoteResult.Warnings | ForEach-Object { "$_" }) -join '; '
                }
            }
            catch {
                $errMsg = ($_.Exception.Message) -replace ',', ';'
                $result.Comment = "BIOS query failed: $errMsg"
            }

            return $result
        }

        # --- Execute via RunspacePool ---
        Write-Host "Starting BIOS configuration query on $($online.Count) machines..."
        $runspaceResults = @(Invoke-RunspacePool $scriptBlock $argList `
            -ThrottleLimit $ThrottleLimit `
            -TimeoutMinutes $TimeoutMinutes)

        # --- Normalize timed-out/failed results ---
        $onlineResults = foreach ($r in $runspaceResults) {
            if ($r -isnot [PSCustomObject]) { continue }
            if ($null -ne $r.PSObject.Properties['BIOSConfiguration']) {
                $r
            }
            else {
                [PSCustomObject]@{
                    ComputerName      = $r.ComputerName
                    Status            = 'Online'
                    Comment           = if ($null -ne $r.Comment) { $r.Comment } else { 'Task Stopped' }
                    BIOSConfiguration = $null
                }
            }
        }
        $allResults = @($onlineResults) + @($offlineResults)


        # --- Display per-machine summary ---
        $allResults | Sort-Object -Property (
            @{Expression = 'BIOSConfiguration'; Descending = $false},
            @{Expression = 'Comment'; Descending = $true}
        ) | Format-Table ComputerName, Status, Comment -AutoSize | Out-Host


        # --- BIOS Comparison / Aggregation ---
        Write-Host "Compiling all BIOS settings in $Mode mode..."

        $master_BIOS_Settings = @()
        $master_PSChildName_Array = @()
        $temp_Array = @()

        foreach ($r in $allResults) {
            if ($null -ne $r.BIOSConfiguration) {
                $temp_Array += $r.BIOSConfiguration.PSChildName
            }
        }

        $master_PSChildName_Array = $temp_Array | Sort-Object -Unique

        foreach ($master_PSChildName in $master_PSChildName_Array) {
            $master_BIOS_Settings += New-Object -TypeName PSObject -Property ([Ordered]@{
                PSChildName  = $master_PSChildName
                CurrentValue = $null
            })
        }

        foreach ($master_BIOS_Setting in $master_BIOS_Settings) {
            foreach ($r in $allResults) {
                if ($null -eq $r.BIOSConfiguration) { continue }

                foreach ($biosSetting in $r.BIOSConfiguration) {
                    if ($biosSetting.PSChildName -eq $master_BIOS_Setting.PSChildName) {
                        if ($null -eq $biosSetting.CurrentValue) {
                            continue
                        }
                        if ($null -eq $master_BIOS_Setting.CurrentValue) {
                            $master_BIOS_Setting.CurrentValue = $biosSetting.CurrentValue
                        }
                        elseif ($Mode -eq "Append") {
                            if ($master_BIOS_Setting.CurrentValue -notcontains $biosSetting.CurrentValue) {
                                if ($master_BIOS_Setting.CurrentValue.Length -gt 1) {
                                    [array]$master_BIOS_Setting.CurrentValue += $($biosSetting.CurrentValue)
                                }
                                else {
                                    $master_BIOS_Setting.CurrentValue += $biosSetting.CurrentValue
                                }
                            }
                        }
                        elseif ($Mode -eq "Remove") {
                            if ($master_BIOS_Setting.CurrentValue -ne $biosSetting.CurrentValue) {
                                $master_BIOS_Setting.CurrentValue = "-"
                            }
                        }
                    }
                }
            }
        }

        # Flatten arrays to comma-separated strings for display
        foreach ($master_BIOS_Setting in $master_BIOS_Settings) {
            if ($null -eq $master_BIOS_Setting.CurrentValue) {
                continue
            }
            $ofs = ","
            [string]$master_BIOS_Setting.CurrentValue = [array]$master_BIOS_Setting.CurrentValue | Sort-Object -Unique
            [string]$master_BIOS_Setting.CurrentValue = [string]$master_BIOS_Setting.CurrentValue.TrimStart(",")
        }
        $ofs = " "

        $master_BIOS_Settings | Format-Table -AutoSize | Out-Host


        # --- Export results ---
        $resultsRoot = "$env:USERPROFILE\Desktop\BIOS_Results"
        if (-not (Test-Path $resultsRoot -PathType Container)) {
            mkdir $resultsRoot -Force > $null
        }

        $dateOutput = Get-Date -Format "yyyy-MM-dd-HHmm"
        $fileName = "BIOS_Configuration_${Mode}_$dateOutput"

        $master_BIOS_Settings | Export-Csv "$resultsRoot\$fileName.csv" -Append -Force -NoTypeInformation

        Write-Host "Report '$fileName.csv' saved to $resultsRoot."

        # Return the aggregated settings for pipeline use
        if ($PassThru) { return $master_BIOS_Settings }
    }
}
