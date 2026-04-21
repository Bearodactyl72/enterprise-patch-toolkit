# DOTS formatting comment

function New-RegistryKey {
    <#
        .SYNOPSIS
            Creates or updates registry keys and values on remote machines.
        .DESCRIPTION
            Sets one or more registry values on remote machines via
            Invoke-RunspacePool. Creates parent keys if they do not exist.
            Reports whether each value was created, updated, or already correct.
        .PARAMETER ComputerName
            One or more computer names to target. Accepts pipeline input.
        .PARAMETER RegistryPath
            One or more registry paths where the value should be set.
            Example: "HKLM:\Software\Microsoft\Cryptography\Wintrust\Config\"
        .PARAMETER Name
            The registry value name to create or update.
        .PARAMETER Value
            The data to set for the registry value.
        .PARAMETER PropertyType
            The registry value type. Default: DWord.
            Valid values: String, ExpandString, Binary, DWord, MultiString, QWord.
        .PARAMETER ThrottleLimit
            Maximum concurrent machines. Default: 50
        .PARAMETER TimeoutMinutes
            Minutes before a machine's task is auto-stopped. Default: 5
        .EXAMPLE
            $list = Get-Content "$env:USERPROFILE\Desktop\Lists\Target_Machines.txt"
            New-RegistryKey -ComputerName $list -RegistryPath @(
                "HKLM:\Software\Wow6432Node\Microsoft\Cryptography\Wintrust\Config\"
                "HKLM:\Software\Microsoft\Cryptography\Wintrust\Config\"
            ) -Name "EnableCertPaddingCheck" -Value 1 -PropertyType DWord
        .NOTES
            Written by Skyler Werner
            Version: 2.0.0
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [string[]]
        $ComputerName,

        [Parameter(Mandatory)]
        [string[]]
        $RegistryPath,

        [Parameter(Mandatory)]
        [string]
        $Name,

        [Parameter(Mandatory)]
        $Value,

        [Parameter()]
        [ValidateSet('String', 'ExpandString', 'Binary', 'DWord', 'MultiString', 'QWord')]
        [string]
        $PropertyType = 'DWord',

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
        foreach ($n in $ComputerName) {
            if ($n.Length -gt 0) {
                $collectedNames += $n
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
        $pingResults = Test-ConnectionAsJob -ComputerName $targets
        $online      = @($pingResults | Where-Object { $_.Reachable } | Select-Object -ExpandProperty ComputerName)
        $offline     = @($pingResults | Where-Object { -not $_.Reachable } | Select-Object -ExpandProperty ComputerName)

        $offlineResults = @()
        foreach ($pc in $offline) {
            Write-Host "$pc Offline" -ForegroundColor Red
            $offlineResults += [PSCustomObject]@{
                ComputerName = $pc
                Status       = 'Offline'
                Action       = $null
                Comment      = 'Offline'
            }
        }

        if ($online.Count -eq 0) {
            Write-Warning "No machines responded to ping."
            if ($PassThru) { return $offlineResults }
            return
        }

        # --- Build argument list ---
        $argList = $online | ForEach-Object { ,@($_, $RegistryPath, $Name, $Value, $PropertyType) }

        # --- Remote scriptblock ---
        $scriptBlock = {
            $computer     = $args[0]
            $regPaths     = [string[]]($args[1] | ForEach-Object { "$_" })
            $regName      = [string]$args[2]
            $regValue     = $args[3]
            $regType      = [string]$args[4]

            $result = [PSCustomObject]@{
                ComputerName = $computer
                Status       = 'Online'
                Action       = $null
                Comment      = ''
            }

            try {
                $remoteOutput = Invoke-Command -ComputerName $computer -ScriptBlock {
                    param($paths, $name, $value, $type)

                    $actions = @()
                    foreach ($regKey in $paths) {
                        # Create key if it does not exist
                        if (-not (Test-Path $regKey)) {
                            New-Item -Path $regKey -Force | Out-Null
                        }

                        $current = Get-ItemProperty -Path $regKey -Name $name -ErrorAction SilentlyContinue
                        if ($null -eq $current -or $current.$name -ne $value) {
                            New-ItemProperty -Path $regKey -Name $name -Value $value -PropertyType $type -Force | Out-Null
                            $actions += "Set $regKey$name = $value ($type)"
                        }
                        else {
                            $actions += "Already correct: $regKey$name"
                        }
                    }
                    return $actions
                } -ArgumentList $regPaths, $regName, $regValue, $regType -ErrorAction Stop

                $actionList = @($remoteOutput | ForEach-Object { "$_" })
                $changed = @($actionList | Where-Object { $_ -match '^Set ' })

                if ($changed.Count -gt 0) {
                    $result.Action = 'Updated'
                }
                else {
                    $result.Action = 'No Change'
                }
                $result.Comment = $actionList -join '; '
            }
            catch {
                $errMsg = ($_.Exception.Message) -replace ',', ';'
                $result.Comment = "Failed: $errMsg"
            }

            return $result
        }

        # --- Execute via RunspacePool ---
        Write-Host "Setting registry values on $($online.Count) machines..."
        $runspaceResults = @(Invoke-RunspacePool $scriptBlock $argList `
            -ThrottleLimit $ThrottleLimit `
            -TimeoutMinutes $TimeoutMinutes)

        # --- Normalize timed-out/failed results ---
        $onlineResults = foreach ($r in $runspaceResults) {
            if ($r -isnot [PSCustomObject]) { continue }
            if ($null -ne $r.PSObject.Properties['Action']) {
                $r
            }
            else {
                [PSCustomObject]@{
                    ComputerName = $r.ComputerName
                    Status       = 'Online'
                    Action       = $null
                    Comment      = if ($null -ne $r.Comment) { $r.Comment } else { 'Task Stopped' }
                }
            }
        }
        $allResults = @($onlineResults) + @($offlineResults)

        $sorted = $allResults | Sort-Object -Property (
            @{Expression = 'Status'; Descending = $true},
            @{Expression = 'Action'; Descending = $false},
            @{Expression = 'ComputerName'; Descending = $false}
        )

        $sorted | Format-Table -AutoSize | Out-Host
        if ($PassThru) { return $sorted }
    }
}
