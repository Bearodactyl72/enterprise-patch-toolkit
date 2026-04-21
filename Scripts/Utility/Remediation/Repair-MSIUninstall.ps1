# DOTS formatting comment

function Repair-MSIUninstall {
    <#
        .SYNOPSIS
            Fixes stuck MSI uninstalls by cleaning up corrupted Windows Installer
            registry entries on remote machines.
        .DESCRIPTION
            USE WHEN: An MSI-based software uninstall keeps failing with exit code
            1603, or the software appears uninstalled but its entry persists in
            Programs and Features / the registry, blocking reinstalls.

            HOW IT WORKS: Windows Installer tracks every MSI product in multiple
            registry locations using a "packed GUID" format (the product code with
            each segment reversed). When these entries become orphaned or corrupted
            -- due to interrupted installs, failed patches, or partial uninstalls --
            future MSI operations on that product fail. This function finds and
            removes those entries, replicating the core fix from Microsoft's
            now-retired "Program Install and Uninstall" troubleshooter (MSPIU/MATS).

            WHAT IT SEARCHES (per product code):
            - HKCR:\Installer\Products\{packedGUID}
            - HKLM:\...\Installer\UserData\{SID}\Products\{packedGUID} (all SIDs)
            - HKLM:\...\Uninstall\{productCode} (native and WOW6432Node)

            By default runs in discovery mode -- shows what it found without
            changing anything. Use -Remove to actually delete the keys.
            Use -WhatIf with -Remove for a dry-run preview.

            Uses Invoke-RunspacePool for concurrent execution.
        .PARAMETER ComputerName
            One or more computer names to target. Accepts pipeline input.
        .PARAMETER SoftwareName
            Display name pattern to search for. Uses -match (regex) by default.
        .PARAMETER ExactMatch
            If specified, uses -eq instead of -match for name comparison.
        .PARAMETER Remove
            If specified, removes the discovered orphaned registry keys.
            Without this switch, only reports what was found.
        .PARAMETER ThrottleLimit
            Maximum concurrent machines. Default: 32
        .PARAMETER TimeoutMinutes
            Minutes before a machine's task is auto-stopped. Default: 10
        .EXAMPLE
            # Step 1: Discover -- see what product codes and orphaned keys exist
            Repair-MSIUninstall -ComputerName "PC01" -SoftwareName "Microsoft Edge"

            # Step 2: Preview -- confirm which keys would be removed
            Repair-MSIUninstall -ComputerName "PC01" -SoftwareName "Microsoft Edge" -Remove -WhatIf

            # Step 3: Fix -- remove the orphaned keys so MSI operations work again
            Repair-MSIUninstall -ComputerName "PC01" -SoftwareName "Microsoft Edge" -Remove
        .EXAMPLE
            $list = Get-Content "$env:USERPROFILE\Desktop\Lists\Target_Machines.txt"
            Repair-MSIUninstall -ComputerName $list -SoftwareName "^Java\s" -Remove
        .NOTES
            Written by Skyler Werner
            Based on Microsoft Program Install and Uninstall troubleshooter (MSPIU).
            C# CompressGUID function derived from MSIMATSFN.ps1 (Microsoft Corp.).
            Version: 1.0.0
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [string[]]
        $ComputerName,

        [Parameter(Mandatory, Position = 1)]
        [string]
        $SoftwareName,

        [Parameter()]
        [switch]
        $ExactMatch,

        [Parameter()]
        [switch]
        $Remove,

        [Parameter()]
        [ValidateRange(1, 300)]
        [int]
        $ThrottleLimit = 32,

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

        # --- Sanitize input ---
        $targets = @(Format-ComputerList $collectedNames -ToUpper)
        if ($targets.Count -eq 0) {
            Write-Warning "No valid computer names provided."
            return
        }

        # --- Ping check ---
        Write-Host "Checking for online machines..."
        $pingResults = Test-ConnectionAsJob -ComputerName $targets
        $online  = @($pingResults | Where-Object { $_.Reachable } | Select-Object -ExpandProperty ComputerName)
        $offline = @($pingResults | Where-Object { -not $_.Reachable } | Select-Object -ExpandProperty ComputerName)

        $offlineResults = @()
        foreach ($pc in $offline) {
            Write-Host "$pc Offline" -ForegroundColor Red
            $offlineResults += [PSCustomObject]@{
                ComputerName  = $pc
                Status        = 'Offline'
                ProductCode   = $null
                Version       = $null
                PackedGUID    = $null
                RegPaths      = $null
                Removed       = $null
                OrphanedPaths = @()
                Comment       = ''
            }
        }

        if ($online.Count -eq 0) {
            Write-Warning "No machines responded to ping."
            if ($PassThru) { return $offlineResults }
        }

        # --- Capture WhatIf state for passing into runspaces ---
        $isWhatIf = $WhatIfPreference

        # --- Build argument list ---
        $argList = $online | ForEach-Object {
            ,@($_, $SoftwareName, [bool]$ExactMatch, [bool]$Remove, $isWhatIf)
        }

        # --- Remote scriptblock ---
        $scriptBlock = {
            $computer    = $args[0]
            $swName      = [string]$args[1]
            $useExact    = [bool]$args[2]
            $doRemove    = [bool]$args[3]
            $dryRun      = [bool]$args[4]

            $results = @()

            try {
                $remoteOutput = Invoke-Command -ComputerName $computer -ErrorAction Stop `
                    -ArgumentList @($swName, $useExact, $doRemove, $dryRun) -ScriptBlock {
                    param($softwareName, $exactMatch, $removeKeys, $isDryRun)

                    $output = @()

                    # --- Compile C# GUID compression helper ---
                    # From Microsoft MSPIU troubleshooter (MSIMATSFN.ps1)
                    $csSource = @'
using System;
public class MSIGuidHelper {
    public static string ReverseString(string s) {
        char[] arr = s.ToCharArray();
        Array.Reverse(arr);
        return new string(arr);
    }
    public static string CompressGUID(string guid) {
        guid = guid.Trim('{', '}');
        return (ReverseString(guid.Substring(0, 8)) +
                ReverseString(guid.Substring(9, 4)) +
                ReverseString(guid.Substring(14, 4)) +
                ReverseString(guid.Substring(19, 2)) +
                ReverseString(guid.Substring(21, 2)) +
                ReverseString(guid.Substring(24, 2)) +
                ReverseString(guid.Substring(26, 2)) +
                ReverseString(guid.Substring(28, 2)) +
                ReverseString(guid.Substring(30, 2)) +
                ReverseString(guid.Substring(32, 2)) +
                ReverseString(guid.Substring(34, 2)));
    }
}
'@
                    try {
                        Add-Type -TypeDefinition $csSource -ErrorAction SilentlyContinue
                    }
                    catch {
                        # Type may already be loaded
                    }

                    # --- Discover product codes by display name ---
                    $regHives = @(
                        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
                        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
                        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
                        'HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
                    )

                    $productMap = @{}
                    foreach ($hive in $regHives) {
                        $children = @(Get-ChildItem $hive -ErrorAction SilentlyContinue -Force)
                        if ($children.Count -eq 0) { continue }

                        $entries = @(Get-ItemProperty $children.PSPath -ErrorAction SilentlyContinue)
                        foreach ($entry in $entries) {
                            if ([string]::IsNullOrEmpty($entry.DisplayName)) { continue }

                            $matched = $false
                            if ($exactMatch) {
                                $matched = ($entry.DisplayName -eq $softwareName)
                            }
                            else {
                                $matched = ($entry.DisplayName -match $softwareName)
                            }
                            if (-not $matched) { continue }

                            # Extract the key name (product code or identifier)
                            $keyName = Split-Path $entry.PSPath -Leaf
                            if (-not $productMap.ContainsKey($keyName)) {
                                $productMap[$keyName] = $entry.DisplayVersion
                            }
                        }
                    }
                    $productCodes = @($productMap.Keys)

                    if ($productCodes.Count -eq 0) {
                        $output += [PSCustomObject]@{
                            ProductCode   = $null
                            Version       = $null
                            PackedGUID    = $null
                            RegPaths      = 0
                            Removed       = 0
                            OrphanedPaths = @()
                            Comment       = 'No matching products found'
                        }
                        return $output
                    }

                    # --- Map HKCR for registry access ---
                    $hkcrMapped = $false
                    if (-not (Test-Path 'HKCR:\')) {
                        New-PSDrive -Name HKCR -PSProvider Registry `
                            -Root HKEY_CLASSES_ROOT -ErrorAction SilentlyContinue | Out-Null
                        $hkcrMapped = $true
                    }

                    # --- Process each product code ---
                    foreach ($code in $productCodes) {
                        $foundPaths  = @()
                        $removedCount = 0

                        # Try to compute packed GUID (only works for proper GUIDs)
                        $packedGuid = $null
                        if ($code -match '^\{[0-9A-Fa-f-]+\}$') {
                            try {
                                $packedGuid = [MSIGuidHelper]::CompressGUID($code)
                            }
                            catch { }
                        }

                        # --- Build list of paths to check ---
                        $pathsToCheck = @()

                        if ($null -ne $packedGuid) {
                            # HKCR packed GUID location
                            $pathsToCheck += "HKCR:\Installer\Products\$packedGuid"

                            # Per-SID UserData locations
                            $userDataRoot = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData'
                            $sids = @(Get-ChildItem $userDataRoot -ErrorAction SilentlyContinue)
                            foreach ($sid in $sids) {
                                $pathsToCheck += Join-Path $sid.PSPath "Products\$packedGuid"
                            }
                        }

                        # Standard uninstall key locations
                        $pathsToCheck += "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$code"
                        $pathsToCheck += "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$code"

                        # --- Check each path ---
                        foreach ($path in $pathsToCheck) {
                            if (Test-Path -Path $path) {
                                $foundPaths += $path

                                if ($removeKeys -and -not $isDryRun) {
                                    try {
                                        Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
                                        $removedCount++
                                    }
                                    catch {
                                        # Path found but removal failed - still count as found
                                    }
                                }
                            }
                        }

                        $comment = $null
                        if ($foundPaths.Count -eq 0) {
                            $comment = 'No orphaned keys found'
                        }
                        elseif ($isDryRun) {
                            $comment = "WhatIf - $($foundPaths.Count) key(s) would be removed"
                        }
                        elseif ($removeKeys) {
                            if ($removedCount -eq $foundPaths.Count) {
                                $comment = "Removed $removedCount key(s)"
                            }
                            else {
                                $comment = "Removed $removedCount of $($foundPaths.Count) key(s)"
                            }
                        }
                        else {
                            $comment = "Found $($foundPaths.Count) key(s) - use -Remove to delete"
                        }

                        $output += [PSCustomObject]@{
                            ProductCode   = $code
                            Version       = $productMap[$code]
                            PackedGUID    = $packedGuid
                            RegPaths      = $foundPaths.Count
                            Removed       = if ($removeKeys -and -not $isDryRun) { $removedCount } else { $null }
                            OrphanedPaths = $foundPaths
                            Comment       = $comment
                        }
                    }

                    if ($hkcrMapped) {
                        Remove-PSDrive HKCR -ErrorAction SilentlyContinue
                    }

                    return $output
                }

                foreach ($item in @($remoteOutput)) {
                    $results += [PSCustomObject]@{
                        ComputerName  = $computer
                        Status        = 'Online'
                        ProductCode   = "$($item.ProductCode)"
                        Version       = "$($item.Version)"
                        PackedGUID    = "$($item.PackedGUID)"
                        RegPaths      = $item.RegPaths
                        Removed       = $item.Removed
                        OrphanedPaths = @($item.OrphanedPaths | ForEach-Object { "$_" })
                        Comment       = "$($item.Comment)"
                    }
                }
            }
            catch {
                $errMsg = ($_.Exception.Message) -replace ',', ';'
                $results += [PSCustomObject]@{
                    ComputerName  = $computer
                    Status        = 'Online'
                    ProductCode   = $null
                    Version       = $null
                    PackedGUID    = $null
                    RegPaths      = $null
                    Removed       = $null
                    OrphanedPaths = @()
                    Comment       = "Failed: $errMsg"
                }
            }

            return $results
        }

        # --- Execute via RunspacePool ---
        $actionLabel = 'Scanning'
        if ($Remove) { $actionLabel = 'Removing' }
        Write-Host "$actionLabel MSI registrations for '$SoftwareName' on $($online.Count) machine(s)..."

        $runspaceResults = @(Invoke-RunspacePool $scriptBlock $argList `
            -ThrottleLimit $ThrottleLimit `
            -TimeoutMinutes $TimeoutMinutes)

        # --- Normalize timed-out/failed results ---
        $onlineResults = foreach ($r in $runspaceResults) {
            if ($r -isnot [PSCustomObject]) { continue }
            if ($null -ne $r.PSObject.Properties['RegPaths']) {
                $r
            }
            else {
                [PSCustomObject]@{
                    ComputerName  = $r.ComputerName
                    Status        = 'Online'
                    ProductCode   = $null
                    Version       = $null
                    PackedGUID    = $null
                    RegPaths      = $null
                    Removed       = $null
                    OrphanedPaths = @()
                    Comment       = if ($null -ne $r.Comment) { $r.Comment } else { 'Task Stopped' }
                }
            }
        }
        $allResults = @($onlineResults) + @($offlineResults)

        $sorted = $allResults | Sort-Object -Property (
            @{Expression = 'Status'; Descending = $true},
            @{Expression = 'RegPaths'; Descending = $true},
            @{Expression = 'ComputerName'; Descending = $false}
        )

        $sorted | Format-Table ComputerName, Status, ProductCode, Version,
            RegPaths, Removed, Comment -AutoSize | Out-Host

        # --- Detail section: list orphaned paths per machine ---
        $withKeys = @($sorted | Where-Object { $_.RegPaths -gt 0 })
        if ($withKeys.Count -gt 0) {
            Write-Host ""
            if ($Remove -and -not $isWhatIf) {
                Write-Host "Orphaned registry keys removed:" -ForegroundColor Yellow
            }
            elseif ($isWhatIf) {
                Write-Host "Orphaned registry keys that would be removed:" -ForegroundColor Yellow
            }
            else {
                Write-Host "Orphaned registry keys found:" -ForegroundColor Yellow
            }
            Write-Host ""

            foreach ($machine in ($withKeys | Sort-Object ComputerName)) {
                Write-Host "  $($machine.ComputerName) [$($machine.ProductCode)]" -ForegroundColor Cyan
                foreach ($path in $machine.OrphanedPaths) {
                    Write-Host "    $path"
                }
                Write-Host ""
            }
        }

        if ($PassThru) { return $sorted }
    }
}
