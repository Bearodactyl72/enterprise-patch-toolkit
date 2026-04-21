# DOTS formatting comment

function Set-MeteredConnection {
    <#
        .SYNOPSIS
            Sets network connections as not metered on remote machines.
        .DESCRIPTION
            Connects to remote machines and disables metered connection status
            for all network adapter profiles (DusmSvc UserCost) and the
            DefaultMediaCost registry keys (Ethernet, Default, WiFi).

            Requires taking ownership of the DefaultMediaCost registry key
            via RtlAdjustPrivilege since it is owned by TrustedInstaller.

            Uses Invoke-RunspacePool for concurrent execution and
            Test-ConnectionAsJob for pre-filtering offline machines.
        .PARAMETER ComputerName
            One or more computer names to target. Accepts pipeline input.
        .PARAMETER ThrottleLimit
            Maximum concurrent machines. Default: 50
        .PARAMETER TimeoutMinutes
            Minutes before a machine's task is auto-stopped. Default: 5
        .EXAMPLE
            $list = Get-Content "$env:USERPROFILE\Desktop\Lists\Target_Machines.txt"
            Set-MeteredConnection -ComputerName $list
        .EXAMPLE
            Set-MeteredConnection -ComputerName "PC01","PC02" | Format-Table -AutoSize
        .NOTES
            Original author: Michael Pietroforte (4sysops.com)
            Rewritten by Skyler Werner
            Date: 2026/03/27
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
        $pingResults = Test-ConnectionAsJob -ComputerName $targets
        $online      = @($pingResults | Where-Object { $_.Reachable } | Select-Object -ExpandProperty ComputerName)
        $offline     = @($pingResults | Where-Object { -not $_.Reachable } | Select-Object -ExpandProperty ComputerName)

        $offlineResults = @()
        foreach ($pc in $offline) {
            Write-Host "$pc Offline" -ForegroundColor Red
            $offlineResults += [PSCustomObject]@{
                ComputerName = $pc
                Status       = 'Offline'
                Changed      = $false
                Comment      = 'Offline'
            }
        }

        if ($online.Count -eq 0) {
            Write-Warning "No machines responded to ping."
            if ($PassThru) { return $offlineResults }
            return
        }

        # --- Build argument list ---
        $argList = $online | ForEach-Object { ,@($_) }

        # --- Remote scriptblock ---
        $scriptBlock = {
            $computer = $args[0]

            $result = [PSCustomObject]@{
                ComputerName = $computer
                Status       = 'Online'
                Changed      = $false
                Comment      = ''
            }

            try {
                $remoteOutput = Invoke-Command -ComputerName $computer -ScriptBlock {

                    $changes  = @()
                    $warnings = @()

                    # --- Part 1: DusmSvc profile UserCost ---
                    $dusmPath = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\DusmSvc\Profiles'
                    if (Test-Path $dusmPath) {
                        $regkeys = @(Get-ChildItem "$dusmPath\*" -ErrorAction SilentlyContinue)
                        foreach ($regkey in $regkeys) {
                            $profilePath = "Registry::$($regkey.Name)\*"
                            $children = Get-ChildItem $profilePath -ErrorAction SilentlyContinue
                            foreach ($child in $children) {
                                $props = Get-ItemProperty $child.PSPath -ErrorAction SilentlyContinue
                                if ($null -ne $props -and $props.UserCost -ne 0) {
                                    Set-ItemProperty $child.PSPath -Name UserCost -Value 0 -Force
                                    $changes += "DusmSvc profile set to not metered"
                                }
                            }
                        }

                        if ($changes.Count -gt 0) {
                            Restart-Service DusmSvc -ErrorAction SilentlyContinue
                        }
                    }

                    # --- Part 2: DefaultMediaCost registry key ---
                    # This key is owned by TrustedInstaller. We need SeTakeOwnership
                    # privilege to change the owner to Administrators.

                    $csDef = @"
using System;
using System.Runtime.InteropServices;
namespace Win32Api {
    public class NtDll {
        [DllImport("ntdll.dll", EntryPoint="RtlAdjustPrivilege")]
        public static extern int RtlAdjustPrivilege(ulong Privilege, bool Enable, bool CurrentThread, ref bool Enabled);
    }
}
"@
                    try {
                        Add-Type -TypeDefinition $csDef -ErrorAction SilentlyContinue
                    }
                    catch {
                        # Type may already be loaded from a prior run
                    }
                    [Win32Api.NtDll]::RtlAdjustPrivilege(9, $true, $false, [ref]$false) | Out-Null

                    # Take ownership
                    $regPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\DefaultMediaCost"
                    $key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey(
                        $regPath,
                        [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree,
                        [System.Security.AccessControl.RegistryRights]::TakeOwnership)

                    if ($null -ne $key) {
                        $acl = $key.GetAccessControl()
                        $acl.SetOwner([System.Security.Principal.NTAccount]"Administrators")
                        $key.SetAccessControl($acl)

                        # Grant FullControl
                        $rule = New-Object System.Security.AccessControl.RegistryAccessRule(
                            [System.Security.Principal.NTAccount]"Administrators",
                            "FullControl",
                            "Allow")
                        $acl.SetAccessRule($rule)
                        $key.SetAccessControl($acl)
                        $key.Close()

                        # Set connections to not metered
                        $costPath = "HKLM:\$regPath"
                        $names = @("Ethernet", "Default", "WiFi")
                        $netChanged = $false

                        foreach ($name in $names) {
                            $metered = Get-ItemProperty -Path $costPath -ErrorAction SilentlyContinue |
                                Select-Object -ExpandProperty $name -ErrorAction SilentlyContinue
                            if ($metered -eq 2) {
                                New-ItemProperty -Path $costPath -Name $name -Value 1 -PropertyType DWORD -Force | Out-Null
                                $changes += "$name set to not metered"
                                $netChanged = $true
                            }
                        }

                        if ($netChanged) {
                            # Restart Network Connections service to apply
                            & net stop Netman 2>$null
                            & net start Netman 2>$null
                        }
                    }
                    else {
                        $warnings += "Could not open DefaultMediaCost key"
                    }

                    return [PSCustomObject]@{
                        Changes  = $changes
                        Warnings = $warnings
                    }
                } -ErrorAction Stop

                $commentParts = @()
                if ($remoteOutput.Changes.Count -gt 0) {
                    $result.Changed = $true
                    $commentParts += @($remoteOutput.Changes | ForEach-Object { "$_" })
                }
                if ($remoteOutput.Warnings.Count -gt 0) {
                    $commentParts += @($remoteOutput.Warnings | ForEach-Object { "$_" })
                }
                if ($commentParts.Count -gt 0) {
                    $result.Comment = $commentParts -join '; '
                }
                else {
                    $result.Comment = "No changes needed"
                }
            }
            catch {
                $errMsg = ($_.Exception.Message) -replace ',', ';'
                $result.Comment = "Failed: $errMsg"
            }

            return $result
        }

        # --- Execute via RunspacePool ---
        Write-Host "Setting metered connection status on $($online.Count) machines..."
        $runspaceResults = @(Invoke-RunspacePool $scriptBlock $argList `
            -ThrottleLimit $ThrottleLimit `
            -TimeoutMinutes $TimeoutMinutes)

        # --- Normalize timed-out/failed results ---
        $onlineResults = foreach ($r in $runspaceResults) {
            if ($r -isnot [PSCustomObject]) { continue }
            if ($null -ne $r.PSObject.Properties['Changed']) {
                $r
            }
            else {
                [PSCustomObject]@{
                    ComputerName = $r.ComputerName
                    Status       = 'Online'
                    Changed      = $false
                    Comment      = if ($null -ne $r.Comment) { $r.Comment } else { 'Task Stopped' }
                }
            }
        }
        $allResults = @($onlineResults) + @($offlineResults)

        $sorted = $allResults | Sort-Object -Property (
            @{Expression = 'Status'; Descending = $true},
            @{Expression = 'ComputerName'; Descending = $false}
        )

        $sorted | Format-Table -AutoSize | Out-Host
        if ($PassThru) { return $sorted }
    }
}
