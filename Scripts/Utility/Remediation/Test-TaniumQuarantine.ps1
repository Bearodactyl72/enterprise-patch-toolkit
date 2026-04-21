# DOTS formatting comment

function Test-TaniumQuarantine {
    <#
        .SYNOPSIS
            Checks Tanium quarantine/isolation policy status on a remote machine.
        .DESCRIPTION
            Connects to a remote machine's registry and reads computer name, IP,
            Tanium quarantine policy, Tanium tags, and MCEDS install date.

            If a Tanium Quarantine isolation policy is found, optionally removes
            it and restarts the machine.
        .PARAMETER ComputerName
            The computer to check. Accepts a single machine name.
        .PARAMETER IPSubnet
            IP subnet filter for DHCP addresses. Default: "138.168*"
        .PARAMETER RemovePolicy
            If specified, removes the isolation policy without prompting.
            The machine will be restarted after removal.
        .EXAMPLE
            Test-TaniumQuarantine -ComputerName "PC01"
        .EXAMPLE
            Test-TaniumQuarantine -ComputerName "PC01" -RemovePolicy
        .NOTES
            Written by Nickolas Berret
            Modified by Skyler Werner
            Version: 5.1
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]
        $ComputerName,

        [Parameter()]
        [string]
        $IPSubnet = "138.168*",

        [Parameter()]
        [switch]
        $RemovePolicy,

        [Parameter()]
        [switch]
        $PassThru
    )

    $ComputerName = ($ComputerName.Trim()).ToUpper()

    # --- Connect to remote registry ---
    try {
        $key = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
    }
    catch {
        Write-Error "Cannot connect to remote registry on '$ComputerName': $($_.Exception.Message)"
        return
    }

    # --- Read registry values ---
    $regErrors = @()
    $computer = $null
    $ip = $null
    $isolationPolicy = $false
    $taniumTags = $null
    $installDate = $null

    # Registry checks to perform (used as switch labels below).
    # "Tanium Quarantine" is a logical label -- its actual path is
    # Software\Policies\Microsoft\Windows\IPsec\Policy\Local.
    $regChecks = @(
        "SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName"
        "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces"
        "Tanium Quarantine"
        "SOFTWARE\WOW6432Node\Tanium\Tanium Client\Sensor Data\Tags"
        "SOFTWARE\MCEDS"
    )

    foreach ($regCheck in $regChecks) {
        try {
            switch ($regCheck) {
                "SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName" {
                    $subKey = $key.OpenSubKey($regCheck)
                    if ($null -ne $subKey) {
                        $computer = $subKey.GetValue("ComputerName")
                        $subKey.Close()
                    }
                }

                "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces" {
                    $subKey = $key.OpenSubKey($regCheck)
                    if ($null -ne $subKey) {
                        $ip = $subKey.GetSubKeyNames() | ForEach-Object {
                            $ifKey = $key.OpenSubKey("$regCheck\$_")
                            $addr = $ifKey.GetValue("DHCPIPAddress")
                            $ifKey.Close()
                            $addr
                        } | Where-Object { $_ -like $IPSubnet }
                        $subKey.Close()
                    }
                }

                "Tanium Quarantine" {
                    $policyPath = "Software\Policies\Microsoft\Windows\IPsec\Policy\Local"
                    $subKey = $key.OpenSubKey($policyPath)
                    if ($null -ne $subKey) {
                        $policy = $subKey.GetSubKeyNames() | ForEach-Object {
                            $pKey = $key.OpenSubKey("$policyPath\$_")
                            $pName = $pKey.GetValue("ipsecName")
                            $pKey.Close()
                            $pName
                        } | Where-Object { $_ -eq "Tanium Quarantine" }
                        $isolationPolicy = ($null -ne $policy)
                        $subKey.Close()
                    }
                }

                "SOFTWARE\WOW6432Node\Tanium\Tanium Client\Sensor Data\Tags" {
                    $subKey = $key.OpenSubKey($regCheck)
                    if ($null -ne $subKey) {
                        # Filter out the expected org tag so only anomalous
                        # extra tags surface in the output.
                        $orgTag = if (Get-Command Import-RSLEnvironment -ErrorAction SilentlyContinue) {
                            (Import-RSLEnvironment).OrgComponentTag
                        } else { $null }
                        $taniumTags = $subKey.GetValueNames() |
                            Where-Object { $orgTag -and ($_ -ne $orgTag) }
                        $subKey.Close()
                    }
                }

                "SOFTWARE\MCEDS" {
                    $subKey = $key.OpenSubKey($regCheck)
                    if ($null -ne $subKey) {
                        $installDate = $subKey.GetValueNames() |
                            Where-Object { $_ -like "*InstallDate*" } |
                            ForEach-Object { "${_}:$($subKey.GetValue($_))" }
                        $subKey.Close()
                    }
                }
            }
        }
        catch {
            $regErrors += $regCheck
        }
    }

    # --- Build and display result ---
    $result = [PSCustomObject]@{
        Computer        = $computer
        IP              = $ip
        IsolationPolicy = $isolationPolicy
        TaniumTags      = $taniumTags
        InstallDate     = $installDate
        PathErrors      = $regErrors
    }

    $result | Format-List | Out-Host

    # --- Policy removal ---
    if ($result.IsolationPolicy -eq $true) {

        $shouldRemove = $false

        if ($RemovePolicy) {
            $shouldRemove = $true
        }
        else {
            $response = Read-Host "Remove Isolation Policy & RESTART $ComputerName? (Y/N)"
            if ($response -eq 'Y') {
                $shouldRemove = $true
            }
            else {
                Write-Host "$ComputerName will not be restarted."
            }
        }

        if ($shouldRemove) {
            if ($PSCmdlet.ShouldProcess($ComputerName, "Remove Tanium Quarantine policy and restart")) {
                $policyPath = "Software\Policies\Microsoft\Windows\IPsec\Policy\Local"
                $policyKey = $key.OpenSubKey($policyPath, $true)
                if ($null -ne $policyKey) {
                    # Only delete subkeys whose ipsecName is "Tanium Quarantine"
                    $deleted = 0
                    foreach ($subName in $policyKey.GetSubKeyNames()) {
                        $subKey = $policyKey.OpenSubKey($subName)
                        $ipsecName = $null
                        if ($null -ne $subKey) {
                            $ipsecName = $subKey.GetValue("ipsecName")
                            $subKey.Close()
                        }
                        if ($ipsecName -eq "Tanium Quarantine") {
                            $policyKey.DeleteSubKeyTree($subName)
                            $deleted++
                        }
                    }
                    $policyKey.Close()

                    if ($deleted -eq 0) {
                        Write-Warning "Tanium Quarantine policy detected earlier but could not be matched for deletion."
                        $key.Close()
                        if ($PassThru) { return $result }
                    }

                    Write-Host "Isolation policy removed ($deleted key(s) deleted)." -ForegroundColor Green

                    Write-Host "$ComputerName is restarting..."
                    Restart-Computer -ComputerName $ComputerName -Force

                    Write-Host "Waiting for $ComputerName to come back online..."
                    $timeout = 300
                    Start-Sleep -Seconds 15
                    $elapsed = 15
                    while ($elapsed -lt $timeout) {
                        if (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
                            Write-Host "$ComputerName is back online." -ForegroundColor Green
                            break
                        }
                        Start-Sleep -Seconds 5
                        $elapsed += 5
                    }
                    if ($elapsed -ge $timeout) {
                        Write-Warning "$ComputerName did not respond within $($timeout / 60) minutes."
                    }
                }
                else {
                    Write-Warning "Could not open policy key for deletion."
                }
            }
        }
    }

    $key.Close()
    if ($PassThru) { return $result }
}
