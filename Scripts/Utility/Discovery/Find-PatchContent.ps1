# DOTS formatting comment

function Find-PatchContent {
    <#
        .SYNOPSIS
            Searches remote machines' ccmcache for files matching a pattern string.
        .DESCRIPTION
            Connects to each target machine and searches the specified cache folders
            for patch content matching the search string. Two search modes:

            1. Install scripts: Finds .ps1 files with "install" in the name whose
               content matches the search string.
            2. Package files: Finds files whose path matches the search string and
               whose parent folder was modified within the MaxAgeDays window.

            Uses Invoke-RunspacePool for concurrent execution across machines.
        .PARAMETER SearchString
            One or more strings to search for in ccmcache content. Matched against
            install script content and file/folder paths.
        .PARAMETER SearchPath
            Folder paths to search on remote machines. Default: C:\Windows\ccmcache\
        .PARAMETER MaxAgeDays
            Only match package files in folders modified within this many days.
            Default: 60
        .PARAMETER ComputerName
            One or more computer names to search. Accepts pipeline input.
        .PARAMETER ThrottleLimit
            Maximum concurrent machines to process. Default: 25
        .PARAMETER TimeoutMinutes
            Minutes before a machine's search task is auto-stopped. Default: 15
        .EXAMPLE
            $list = Get-Content "$env:USERPROFILE\Desktop\Lists\Target_Machines.txt"
            Find-PatchContent -SearchString "firefox" -ComputerName $list
        .EXAMPLE
            Find-PatchContent -SearchString "rdrdc","jre" -ComputerName $list -MaxAgeDays 30
        .NOTES
            Written by Skyler Werner
            Date: 2026/03/23
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string[]]
        $SearchString,

        [Parameter()]
        [string[]]
        $SearchPath = @("C:\Windows\ccmcache\"),

        [Parameter()]
        [ValidateRange(1, 3650)]
        [int]
        $MaxAgeDays = 60,

        [Parameter(Mandatory, Position = 1, ValueFromPipeline)]
        [string[]]
        $ComputerName,

        [Parameter()]
        [ValidateRange(1, 100)]
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

    $list = Format-ComputerList $collectedNames -ToUpper

    # Capture Get-FileMetaData definition to inject into remote sessions
    $fileMetaDataDef = "function Get-FileMetaData { ${function:Get-FileMetaData} }"


    $scriptblock = {
        $computer      = $args[0]
        $searchStrings = $args[1]
        $searchPaths   = $args[2]
        $maxAge        = $args[3]
        $metaDataDef   = $args[4]

        $remoteBlock = {
            param($searchStrings, $searchPaths, $maxAge, $metaDataDef)

            Invoke-Expression $metaDataDef

            # Force-cast to string[] -- deserialized types from runspace args
            # fail implicit string conversion in Get-ChildItem -Path (PS 5.1)
            $searchStrings = [string[]]($searchStrings | ForEach-Object { "$_" })
            $searchPaths   = [string[]]($searchPaths   | ForEach-Object { "$_" })

            $cutoffDate = (Get-Date).AddDays(-$maxAge)
            $files = @(Get-ChildItem $searchPaths -Recurse -Force -ErrorAction SilentlyContinue)

            $foundItems = [System.Collections.Generic.List[object]]::new()

            foreach ($target in $searchStrings) {

                # --- Search install .ps1 files for content matches ---
                $installFiles = @($files | Where-Object {
                    $_.FullName -match "install" -and $_.Extension -eq ".ps1"
                })

                foreach ($file in $installFiles) {
                    [string]$scriptContent = Get-Content $file.FullName -ErrorAction SilentlyContinue

                    if ($scriptContent -match $target) {
                        $foundItems.Add([PSCustomObject]@{
                            "File Name"     = $file.Name
                            "Version"       = $file.VersionInfo.FileVersion
                            "Date Modified" = $file.LastWriteTime
                            "Folder Date"   = $file.Directory.LastWriteTime
                            "Comment"       = "'$target' found in script content"
                            "Location"      = $file.Directory.FullName
                            "Remote Path"   = "\\" + $env:COMPUTERNAME + "\" +
                                $file.Directory.FullName.Replace("C:", "C$")
                            "Match Type"    = "Install Script"
                        })
                    }
                }

                # --- Search package files by path match ---
                foreach ($file in $files) {
                    if (($file.FullName -match $target) -and
                        ($file.Directory.LastWriteTime -gt $cutoffDate) -and
                        (@(Get-ChildItem $file.DirectoryName -ErrorAction SilentlyContinue).Count -ge 1)) {

                        $version = $file.VersionInfo.FileVersion
                        $comment = $null

                        if (-not $version) {
                            $meta = Get-ChildItem $file.FullName -Force -ErrorAction SilentlyContinue |
                                Get-FileMetaData -ErrorAction SilentlyContinue
                            if ($meta) { $comment = $meta.Comments }
                        }

                        if (-not $comment) {
                            $comment = "'$target' matched in path"
                        }

                        $foundItems.Add([PSCustomObject]@{
                            "File Name"     = $file.Name
                            "Version"       = $version
                            "Date Modified" = $file.LastWriteTime
                            "Folder Date"   = $file.Directory.LastWriteTime
                            "Comment"       = $comment
                            "Location"      = $file.Directory.FullName
                            "Remote Path"   = "\\" + $env:COMPUTERNAME + "\" +
                                $file.Directory.FullName.Replace("C:", "C$")
                            "Match Type"    = "Package File"
                        })
                    }
                }
            }

            [PSCustomObject]@{
                MatchCount = $foundItems.Count
                Matches    = $foundItems
            }
        }

        $result = Invoke-Command -ComputerName $computer -ScriptBlock $remoteBlock `
            -ArgumentList @(, $searchStrings), @(, $searchPaths), $maxAge, $metaDataDef `
            -ErrorAction SilentlyContinue -ErrorVariable remoteError

        if ($remoteError) {
            [PSCustomObject]@{
                ComputerName = $computer
                MatchCount   = $null
                Matches      = @()
                Comment      = "$remoteError"
            }
        }
        elseif ($result) {
            [PSCustomObject]@{
                ComputerName = $computer
                MatchCount   = $result.MatchCount
                Matches      = @($result.Matches)
                Comment      = if ($result.MatchCount -gt 0) {
                    "$($result.MatchCount) match(es) found"
                } else { "No matches" }
            }
        }
        else {
            [PSCustomObject]@{
                ComputerName = $computer
                MatchCount   = $null
                Matches      = @()
                Comment      = "No response"
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
                ComputerName = $pingResult.ComputerName
                MatchCount   = $null
                Matches      = @()
                Comment      = "Offline"
            }
        }
    }

    if ($onlineList.Count -eq 0) {
        $offlineSorted = @($offlineResults | Sort-Object ComputerName)
        $offlineSorted | Format-Table -AutoSize | Out-Host
        if ($PassThru) { return $offlineSorted }
        return
    }


    # --- Run search via Invoke-RunspacePool ---
    $argList = foreach ($computer in $onlineList) {
        , @($computer, $SearchString, $SearchPath, $MaxAgeDays, $fileMetaDataDef)
    }

    $results = @(Invoke-RunspacePool -ScriptBlock $scriptblock -ArgumentList $argList `
        -ThrottleLimit $ThrottleLimit -TimeoutMinutes $TimeoutMinutes `
        -ActivityName "Find Patch Content")

    # Merge with offline results
    $allResults = @($results) + @($offlineResults)

    $finalResults = foreach ($r in $allResults) {
        if ($r.PSObject.Properties.Name -contains "MatchCount") {
            $r
        }
        else {
            [PSCustomObject]@{
                ComputerName = $r.ComputerName
                MatchCount   = $null
                Matches      = @()
                Comment      = $r.Comment
            }
        }
    }


    # --- Output results ---

    $sortedResults = @($finalResults | Sort-Object -Property (
        @{Expression = { if ($_.Comment -eq "Offline") { 1 } else { 0 } }},
        @{Expression = "MatchCount"; Descending = $true},
        @{Expression = "Comment"}
    ))

    # --- Header ---
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Find Patch Content -- Results" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Search string(s):  $($SearchString -join ', ')" -ForegroundColor Gray
    Write-Host "  Search path(s):    $($SearchPath -join ', ')" -ForegroundColor Gray
    Write-Host "  Max age:           $MaxAgeDays days" -ForegroundColor Gray
    Write-Host ""


    # --- Summary table ---
    $sortedResults |
        Select-Object ComputerName, MatchCount, Comment |
        Format-Table -AutoSize | Out-Host


    # --- Found files detail ---
    $allMatches = @()
    foreach ($r in $sortedResults) {
        foreach ($m in @($r.Matches)) {
            if ($m) {
                $m | Add-Member -NotePropertyName PSComputerName `
                    -NotePropertyValue $r.ComputerName -Force -ErrorAction SilentlyContinue
                $allMatches += $m
            }
        }
    }

    if ($allMatches.Count -gt 0) {
        Write-Host "----------------------------------------" -ForegroundColor DarkGray
        Write-Host "  Found Files Detail" -ForegroundColor Cyan
        Write-Host "----------------------------------------" -ForegroundColor DarkGray
        Write-Host ""

        $allMatches |
            Select-Object PSComputerName, "File Name", Version, "Date Modified",
                "Folder Date", Comment, "Match Type", "Remote Path" |
            Sort-Object -Property (
                @{Expression = "Version"; Descending = $true},
                @{Expression = "Date Modified"; Descending = $true},
                @{Expression = "PSComputerName"}
            ) | Format-Table -AutoSize | Out-Host
    }


    # --- Totals ---
    $totalMatches    = ($sortedResults | ForEach-Object { $_.MatchCount } | Measure-Object -Sum).Sum
    $machinesOnline  = @($sortedResults | Where-Object { $_.Comment -ne "Offline" }).Count
    $machinesOffline = @($sortedResults | Where-Object { $_.Comment -eq "Offline" }).Count
    $machinesFound   = @($sortedResults | Where-Object { $_.MatchCount -gt 0 }).Count
    $machinesClean   = @($sortedResults | Where-Object { $_.Comment -eq "No matches" }).Count
    $machinesFailed  = @($sortedResults | Where-Object {
        $_.Comment -match "Task (Stopped|Failed|Error)" -or
        $_.Comment -eq "No response"
    }).Count

    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Totals" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Machines targeted:    $($sortedResults.Count)" -ForegroundColor Gray
    Write-Host "    Online:             $machinesOnline" -ForegroundColor Gray
    Write-Host "    Offline:            $machinesOffline" -ForegroundColor Gray
    Write-Host ""
    Write-Host "    With matches:       $machinesFound" -ForegroundColor Green
    Write-Host "    No matches:         $machinesClean" -ForegroundColor Gray

    if ($machinesFailed -gt 0) {
        Write-Host "    Failed/No response: $machinesFailed" -ForegroundColor Red
    }

    Write-Host ""
    Write-Host "  Total files found:    $totalMatches" -ForegroundColor Gray
    Write-Host ""

    if ($PassThru) { return $sortedResults }

    } # end 'end' block
}
