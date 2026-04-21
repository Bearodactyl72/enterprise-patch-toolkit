# DOTS formatting comment

function Clear-CcmCache {
    <#
        .SYNOPSIS
            Clears old SCCM/MECM ccmcache content from remote machines using the
            Configuration Manager client COM interface.
        .DESCRIPTION
            Connects to each target machine and removes stale cache elements from
            C:\Windows\ccmcache through the SCCM client's UIResource.UIResourceMgr
            COM object. This ensures the client's internal tracking database stays
            in sync with the filesystem, preventing content re-download failures
            and deployment issues.

            Safety checks:
            - Verifies the CcmExec service is running before attempting cleanup
            - Skips cache elements marked as PersistInCache (pinned by active
              deployments or task sequences)
            - Uses LastReferenced time from the SCCM WMI provider instead of
              filesystem timestamps for accurate age detection
            - Skips machines with sufficient free disk space

            Use -WhatIf to preview what would be removed without making changes.
        .PARAMETER ComputerName
            One or more computer names to clean up. Accepts pipeline input.
        .PARAMETER MinFreeGB
            Minimum free disk space in GB. Machines above this threshold are skipped.
            Default: 10
        .PARAMETER MaxAgeDays
            Cache elements not referenced within this many days are eligible for
            removal. Default: 30
        .PARAMETER ThrottleLimit
            Maximum concurrent machines to process. Default: 25
        .PARAMETER TimeoutMinutes
            Minutes before a machine's cleanup task is auto-stopped. Default: 20
        .EXAMPLE
            $list = Get-Content "$env:USERPROFILE\Desktop\Lists\Target_Machines.txt"
            Clear-CcmCache -ComputerName $list
        .EXAMPLE
            Clear-CcmCache -ComputerName $list -WhatIf
        .EXAMPLE
            Clear-CcmCache -ComputerName "PC01","PC02" -MinFreeGB 5 -MaxAgeDays 14
        .NOTES
            Written by Skyler Werner
            Date: 2026/03/23
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [string[]]
        $ComputerName,

        [Parameter()]
        [ValidateRange(1, 100)]
        [int]
        $MinFreeGB = 10,

        [Parameter()]
        [ValidateRange(1, 3650)]
        [int]
        $MaxAgeDays = 30,

        [Parameter()]
        [ValidateRange(1, 100)]
        [int]
        $ThrottleLimit = 25,

        [Parameter()]
        [ValidateRange(1, 120)]
        [int]
        $TimeoutMinutes = 20,

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

    $list = Format-ComputerList $collectedNames -ToUpper

    $dryRun = $WhatIfPreference


    $scriptblock = {
        $computer = $args[0]
        $size     = $args[1]
        $maxAge   = $args[2]
        $dryRun   = $args[3]

        $remoteBlock = {
            param($size, $maxAge, $dryRun)

            $cutoffDate = (Get-Date).AddDays(-$maxAge)

            # --- Check disk space ---
            $disk = Get-PSDrive ($env:SystemDrive).Replace(":", "")
            $diskFreeGB = [math]::Round($($disk.Free / 1gb), 2)

            # Template for early returns
            $emptyResult = [PSCustomObject]@{
                PkgsRemoved     = $null
                PkgsSkipped     = $null
                PinnedCount     = $null
                OrphansRemoved  = $null
                SizeRemovedMB   = $null
                SpaceFreed      = $null
                TotalFree       = $diskFreeGB
                RemovedNames    = @()
                SkippedNames    = @()
                PinnedNames     = @()
                OrphanNames     = @()
                Comment         = ""
            }

            if ($diskFreeGB -gt $size) {
                $emptyResult.Comment = "Cleanup not initiated"
                return $emptyResult
            }

            # --- Verify CcmExec service is running ---
            $ccmService = Get-Service CcmExec -ErrorAction SilentlyContinue
            if (-not $ccmService) {
                $emptyResult.Comment = "SCCM client not installed"
                return $emptyResult
            }
            if ($ccmService.Status -ne 'Running') {
                $emptyResult.Comment = "CcmExec service not running ($($ccmService.Status))"
                return $emptyResult
            }

            # --- Query cache elements via CIM ---
            # CacheInfoEx in ROOT\ccm\SoftMgmt provides LastReferenced and
            # PersistInCache, which the COM interface does not expose directly.
            $cimElements = @()
            try {
                $cimElements = @(Get-CimInstance -Namespace ROOT\ccm\SoftMgmt `
                    -ClassName CacheInfoEx -ErrorAction Stop)
            }
            catch {
                $emptyResult.Comment = "Failed to query SCCM cache: $_"
                return $emptyResult
            }

            # Build lookup of CacheElementId -> CIM properties for age/pinned checks
            $cimById = @{}
            foreach ($el in $cimElements) {
                $cimById[$el.CacheElementId] = $el
            }

            # --- Get COM cache manager for proper deletion ---
            $comCache = $null
            try {
                $cmObject = New-Object -ComObject UIResource.UIResourceMgr
                $comCache = $cmObject.GetCacheInfo()
            }
            catch {
                $emptyResult.Comment = "Failed to connect to SCCM COM interface: $_"
                return $emptyResult
            }

            $comElements = @($comCache.GetCacheElements())

            # --- Evaluate each tracked cache element ---
            $removedNames  = @()
            $skippedNames  = @()
            $pinnedNames   = @()
            $sizeRemovedKB = 0

            # Build set of tracked folder paths for orphan detection later
            $trackedLocations = @{}

            foreach ($comEl in $comElements) {
                $elId = $comEl.CacheElementID
                $contentId = $comEl.ContentId
                $contentVer = $comEl.ContentVersion
                $location = $comEl.Location
                $contentSizeKB = $comEl.ContentSize

                if ($location) {
                    $trackedLocations[$location] = $true
                }

                # Display name: ContentId with version, or folder name as fallback
                $displayName = $contentId
                if ($contentVer) {
                    $displayName = "$contentId v$contentVer"
                }
                if (-not $displayName) {
                    $displayName = Split-Path $location -Leaf
                }

                # Check PersistInCache from CIM data
                $cimEl = $cimById[$elId]
                if ($cimEl -and $cimEl.PersistInCache) {
                    $pinnedNames += $displayName
                    continue
                }

                # Check age using LastReferenced from CIM data
                $lastRef = $null
                if ($cimEl -and $cimEl.LastReferenced) {
                    $lastRef = $cimEl.LastReferenced
                }

                # Fall back to COM LastReferenceTime if CIM didn't have it
                if (-not $lastRef) {
                    try {
                        $lastRef = $comEl.LastReferenceTime
                    }
                    catch {}
                }

                # If still no reference time, use a very old date so it gets cleaned
                if (-not $lastRef) {
                    $lastRef = [datetime]::MinValue
                }

                if ($lastRef -gt $cutoffDate) {
                    $skippedNames += "$displayName (referenced $('{0:yyyy-MM-dd}' -f $lastRef))"
                    continue
                }

                # --- Element is eligible for removal ---
                if (-not $dryRun) {
                    try {
                        $comCache.DeleteCacheElement($elId)
                    }
                    catch {
                        $skippedNames += "$displayName (delete failed: $_)"
                        continue
                    }
                }

                $sizeRemovedKB += $contentSizeKB
                $sizeMB = [math]::Round($contentSizeKB / 1024, 1)
                $removedNames += "$displayName ($sizeMB MB)"
            }

            $sizeRemovedMB = [math]::Round($sizeRemovedKB / 1024, 1)


            # --- Orphaned folder cleanup ---
            # Folders in ccmcache that the SCCM client doesn't track (lost during
            # client repair, reinstall, or database corruption). Safe to remove
            # since the client has no record of them.
            $orphanNames = @()
            $cachePath = "C:\Windows\ccmcache"

            if (Test-Path $cachePath) {
                $diskFolders = @(Get-ChildItem $cachePath -Directory -Force -ErrorAction SilentlyContinue)

                foreach ($folder in $diskFolders) {
                    if ($trackedLocations.ContainsKey($folder.FullName)) { continue }

                    # Measure folder size
                    $folderSizeKB = 0
                    $folderFiles = @(Get-ChildItem $folder.FullName -Recurse -Force -ErrorAction SilentlyContinue |
                        Where-Object { -not $_.PSIsContainer })
                    foreach ($f in $folderFiles) {
                        $folderSizeKB += [math]::Round($f.Length / 1024, 0)
                    }

                    $folderSizeMB = [math]::Round($folderSizeKB / 1024, 1)

                    if (-not $dryRun) {
                        Remove-Item $folder.FullName -Recurse -Force -ErrorAction SilentlyContinue
                    }

                    $sizeRemovedKB += $folderSizeKB
                    $sizeRemovedMB = [math]::Round($sizeRemovedKB / 1024, 1)
                    $orphanNames += "$($folder.Name) ($folderSizeMB MB)"
                }
            }


            # --- Calculate results ---
            $totalActions = $removedNames.Count + $orphanNames.Count

            if ($dryRun) {
                $freedSizeGB = $null
                if ($totalActions -eq 0) {
                    $comment = "No packages eligible for deletion"
                }
                else {
                    $comment = "WhatIf -- no changes made"
                }
            }
            else {
                $cleanDisk = Get-PSDrive ($env:SystemDrive).Replace(":", "")
                $cleanFreeGB = [math]::Round($($cleanDisk.Free / 1gb), 2)
                $freedSizeGB = [math]::Round($cleanFreeGB - $diskFreeGB, 2)

                if ($totalActions -eq 0) {
                    $comment = "No packages eligible for deletion"
                }
                else {
                    $comment = "Cleanup complete"
                }
            }

            [PSCustomObject]@{
                PkgsRemoved     = $removedNames.Count
                PkgsSkipped     = $skippedNames.Count
                PinnedCount     = $pinnedNames.Count
                OrphansRemoved  = $orphanNames.Count
                SizeRemovedMB   = $sizeRemovedMB
                SpaceFreed      = $freedSizeGB
                TotalFree       = if ($dryRun) { $diskFreeGB } else { $cleanFreeGB }
                RemovedNames    = $removedNames
                SkippedNames    = $skippedNames
                PinnedNames     = $pinnedNames
                OrphanNames     = $orphanNames
                Comment         = $comment
            }
        }

        $result = Invoke-Command -ComputerName $computer -ScriptBlock $remoteBlock `
            -ArgumentList $size, $maxAge, $dryRun

        if ($result) {
            [PSCustomObject]@{
                ComputerName       = $computer
                "Pkgs Removed"     = $result.PkgsRemoved
                "Pkgs Skipped"     = $result.PkgsSkipped
                "Pinned"           = $result.PinnedCount
                "Orphans"          = $result.OrphansRemoved
                "Cache Freed (MB)" = $result.SizeRemovedMB
                "Space Freed"      = $result.SpaceFreed
                "Total Free"       = $result.TotalFree
                RemovedNames       = $result.RemovedNames
                SkippedNames       = $result.SkippedNames
                PinnedNames        = $result.PinnedNames
                OrphanNames        = $result.OrphanNames
                Comment            = $result.Comment
            }
        }
    }


    # --- Test connection ---
    Write-Host ""
    Write-Host "Checking for online machines..."

    $pingResults = Test-ConnectionAsJob -ComputerName $list

    $offlineResults = @()
    $onlineList = @()

    foreach ($pingResult in $pingResults) {
        if ($pingResult.Reachable -eq $true) {
            $onlineList += $pingResult.ComputerName
        }
        else {
            $offlineResults += [PSCustomObject]@{
                ComputerName       = $pingResult.ComputerName
                "Pkgs Removed"     = $null
                "Pkgs Skipped"     = $null
                "Pinned"           = $null
                "Orphans"          = $null
                "Cache Freed (MB)" = $null
                "Space Freed"      = $null
                "Total Free"       = $null
                RemovedNames       = @()
                SkippedNames       = @()
                PinnedNames        = @()
                OrphanNames        = @()
                Comment            = "Offline"
            }
        }
    }

    if ($onlineList.Count -eq 0) {
        $offlineSorted = @($offlineResults | Sort-Object ComputerName)
        $offlineSorted | Format-Table -AutoSize | Out-Host
        if ($PassThru) { return $offlineSorted }
        return
    }


    # --- Run cleanup via Invoke-RunspacePool ---
    $argList = foreach ($computer in $onlineList) {
        , @($computer, $MinFreeGB, $MaxAgeDays, $dryRun)
    }

    $activityLabel = "Clear CcmCache"
    if ($dryRun) { $activityLabel = "Clear CcmCache (WhatIf)" }

    $results = @(Invoke-RunspacePool -ScriptBlock $scriptblock -ArgumentList $argList `
        -ThrottleLimit $ThrottleLimit -TimeoutMinutes $TimeoutMinutes `
        -ActivityName $activityLabel)

    # Normalize timed-out/failed results into the same shape
    $allResults = @($results) + @($offlineResults)

    $finalResults = foreach ($r in $allResults) {
        if ($r.PSObject.Properties.Name -contains "Pkgs Removed") {
            $r
        }
        else {
            [PSCustomObject]@{
                ComputerName       = $r.ComputerName
                "Pkgs Removed"     = $null
                "Pkgs Skipped"     = $null
                "Pinned"           = $null
                "Orphans"          = $null
                "Cache Freed (MB)" = $null
                "Space Freed"      = $null
                "Total Free"       = $null
                RemovedNames       = @()
                SkippedNames       = @()
                PinnedNames        = @()
                OrphanNames        = @()
                Comment            = $r.Comment
            }
        }
    }


    # --- Output results ---

    $sortedResults = @($finalResults | Sort-Object -Property (
        @{Expression = { if ($_.Comment -eq "Offline") { 1 } else { 0 } }},
        @{Expression = "Total Free"; Descending = $false},
        @{Expression = "Space Freed"; Descending = $false},
        @{Expression = "Comment"}
    ))

    # --- Header ---
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan

    if ($dryRun) {
        Write-Host "  Clear CcmCache -- WhatIf Preview" -ForegroundColor Cyan
        Write-Host "  (No changes were made)" -ForegroundColor Yellow
    }
    else {
        Write-Host "  Clear CcmCache -- Results Summary" -ForegroundColor Cyan
    }

    Write-Host "========================================" -ForegroundColor Cyan

    if ($dryRun) {
        Write-Host ""
        Write-Host "  Parameters:" -ForegroundColor Gray
        Write-Host "    MinFreeGB:   $MinFreeGB GB" -ForegroundColor Gray
        Write-Host "    MaxAgeDays:  $MaxAgeDays days" -ForegroundColor Gray
    }

    Write-Host ""


    # --- Summary table ---
    $tableProps = @(
        "ComputerName"
        "Pkgs Removed"
        "Pkgs Skipped"
        "Pinned"
        "Orphans"
        "Cache Freed (MB)"
    )

    if (-not $dryRun) {
        $tableProps += "Space Freed"
    }

    $tableProps += "Total Free"
    $tableProps += "Comment"

    $sortedResults | Select-Object $tableProps | Format-Table -AutoSize | Out-Host


    # --- Per-machine detail ---
    $detailLabel = if ($dryRun) { "would be removed" } else { "removed" }

    $detailResults = @($sortedResults | Where-Object {
        $_.Comment -match "(Cleanup complete|WhatIf)" -or
        @($_.PinnedNames).Count -gt 0
    })

    if ($detailResults.Count -gt 0) {
        Write-Host "----------------------------------------" -ForegroundColor DarkGray
        Write-Host "  Per-Machine Detail" -ForegroundColor Cyan
        Write-Host "----------------------------------------" -ForegroundColor DarkGray

        foreach ($r in $detailResults) {
            Write-Host ""
            Write-Host "  $($r.ComputerName)" -ForegroundColor Yellow

            if (@($r.RemovedNames).Count -gt 0) {
                Write-Host "    Packages $($detailLabel):" -ForegroundColor Gray
                foreach ($pkg in $r.RemovedNames) {
                    Write-Host "      $pkg"
                }
            }
            if (@($r.OrphanNames).Count -gt 0) {
                Write-Host "    Orphaned folders $($detailLabel):" -ForegroundColor Gray
                foreach ($pkg in $r.OrphanNames) {
                    Write-Host "      $pkg"
                }
            }
            if (@($r.PinnedNames).Count -gt 0) {
                Write-Host "    Pinned (protected):" -ForegroundColor Gray
                foreach ($pkg in $r.PinnedNames) {
                    Write-Host "      $pkg" -ForegroundColor DarkYellow
                }
            }
            if (@($r.SkippedNames).Count -gt 0) {
                Write-Host "    Skipped (recent or failed):" -ForegroundColor Gray
                foreach ($pkg in $r.SkippedNames) {
                    Write-Host "      $pkg" -ForegroundColor DarkGray
                }
            }
        }
        Write-Host ""
    }


    # --- Totals ---
    $totalRemoved = ($sortedResults | ForEach-Object { $_."Pkgs Removed" }     | Measure-Object -Sum).Sum
    $totalSkipped = ($sortedResults | ForEach-Object { $_."Pkgs Skipped" }     | Measure-Object -Sum).Sum
    $totalPinned  = ($sortedResults | ForEach-Object { $_."Pinned" }           | Measure-Object -Sum).Sum
    $totalOrphans = ($sortedResults | ForEach-Object { $_."Orphans" }          | Measure-Object -Sum).Sum
    $totalCacheMB = ($sortedResults | ForEach-Object { $_."Cache Freed (MB)" } | Measure-Object -Sum).Sum
    $totalCacheMB = [math]::Round($totalCacheMB, 1)

    $machinesOnline  = @($sortedResults | Where-Object { $_.Comment -ne "Offline" }).Count
    $machinesOffline = @($sortedResults | Where-Object { $_.Comment -eq "Offline" }).Count
    $machinesCleaned = @($sortedResults | Where-Object {
        $_.Comment -eq "Cleanup complete" -or $_.Comment -eq "WhatIf -- no changes made"
    }).Count
    $machinesSkipped = @($sortedResults | Where-Object { $_.Comment -eq "Cleanup not initiated" }).Count
    $machinesNone    = @($sortedResults | Where-Object { $_.Comment -eq "No packages eligible for deletion" }).Count
    $machinesNoSccm  = @($sortedResults | Where-Object {
        $_.Comment -match "SCCM client not installed|CcmExec service not running"
    }).Count
    $machinesFailed  = @($sortedResults | Where-Object {
        $_.Comment -match "Task (Stopped|Failed|Error)" -or
        $_.Comment -match "^Failed to" -or
        $_.Comment -eq "Job Failed"
    }).Count

    $pkgVerb = if ($dryRun) { "flagged for removal" } else { "removed" }

    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Totals" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Machines targeted:    $($sortedResults.Count)" -ForegroundColor Gray
    Write-Host "    Online:             $machinesOnline" -ForegroundColor Gray
    Write-Host "    Offline:            $machinesOffline" -ForegroundColor Gray
    Write-Host ""

    if ($dryRun) {
        Write-Host "    Would be cleaned:   $machinesCleaned" -ForegroundColor Yellow
    }
    else {
        Write-Host "    Cleaned:            $machinesCleaned" -ForegroundColor Green
    }

    Write-Host "    Disk OK (>$MinFreeGB GB):   $machinesSkipped" -ForegroundColor Gray
    Write-Host "    Nothing to clean:   $machinesNone" -ForegroundColor Gray

    if ($machinesNoSccm -gt 0) {
        Write-Host "    No SCCM client:     $machinesNoSccm" -ForegroundColor DarkYellow
    }
    if ($machinesFailed -gt 0) {
        Write-Host "    Failed/Stopped:     $machinesFailed" -ForegroundColor Red
    }

    Write-Host ""
    Write-Host "  Packages $($pkgVerb): $totalRemoved" -ForegroundColor Gray
    Write-Host "  Packages skipped:     $totalSkipped (recent)" -ForegroundColor Gray
    Write-Host "  Packages pinned:      $totalPinned (persistent)" -ForegroundColor Gray
    Write-Host "  Orphaned folders $($pkgVerb): $totalOrphans" -ForegroundColor Gray
    Write-Host "  Cache content freed:  $totalCacheMB MB" -ForegroundColor Gray

    if (-not $dryRun) {
        $totalFreed   = ($sortedResults | ForEach-Object { $_."Space Freed" } | Measure-Object -Sum).Sum
        $totalFreedGB = [math]::Round($totalFreed, 2)
        Write-Host "  Total space freed:    $totalFreedGB GB" -ForegroundColor Gray
    }

    Write-Host ""

    if ($PassThru) { return $sortedResults }

    } # end 'end' block
}
