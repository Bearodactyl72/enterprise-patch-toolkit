# DOTS formatting comment

function Get-AcrobatLicenseDiagnostics {
    <#
        .SYNOPSIS
            Collects Adobe Acrobat DC crash and licensing diagnostics from remote
            machines, or analyzes an exported Application .evtx log offline.
        .DESCRIPTION
            Built for the "Acrobat opens, flashes a blank popup, then closes 10-20
            seconds later" failure pattern. That signature is almost always the
            Adobe licensing layer (NGL - Next Generation Licensing) failing its
            entitlement check after launch, not a document or install problem:
            Acrobat starts, the background license refresh fails, a dialog is
            thrown and the process exits. Intermittent behavior usually means a
            cached license token works some days and the refresh fails on others.

            Remote mode gathers, per machine:
              - Acrobat DC install + version from the registry uninstall hive
              - Application event log crash/hang/WER entries that reference
                Acrobat or Adobe (Event ID 1000/1001/1002/1026) with faulting
                module, plus any events from Adobe-named providers
              - NGL client logs (NGLClient_*.log in user temp) with recent
                error/denial/grace-period lines extracted
              - Leftover licensing artifacts from pre-migration installs:
                SLStore/SLCache (legacy serial) and OperatingConfigs (FRL/SDL)
              - Adobe service state (Genuine Monitor, Genuine Software Integrity,
                ARM updater)
              - TCP 443 reachability from the target to the Adobe licensing
                endpoints the NGL refresh depends on

            Offline mode (-EvtxPath) parses an exported Application .evtx (e.g.
            one mailed in from another enclave and renamed from .txt) and prints
            a timeline of Acrobat-related crash and licensing events, optionally
            narrowed to a window around a reported failure time.

            Uses Invoke-RunspacePool for concurrent execution and
            Test-ConnectionAsJob for pre-filtering offline machines. Results
            cross the runspace + Invoke-Command double-serialization boundary,
            so every field on the result object is a string, int, or bool.
        .PARAMETER ComputerName
            One or more computer names to query. Accepts pipeline input.
        .PARAMETER DaysBack
            How many days of Application event log history to search on each
            target. Default: 14
        .PARAMETER SkipEndpointTest
            Skip the TCP 443 reachability test against Adobe licensing
            endpoints (useful on enclaves where outbound 443 is always proxied
            and the direct test would be misleading).
        .PARAMETER EvtxPath
            Path to an exported .evtx file to analyze offline instead of
            querying remote machines. If the export arrived with a .txt
            extension (mail filter workaround), rename it to .evtx first.
        .PARAMETER AroundTime
            Optional timestamp of interest for offline .evtx analysis. Only
            events within +/- WindowMinutes of this time are shown in the
            timeline (the summary still counts the whole file).
        .PARAMETER WindowMinutes
            Half-width of the window around -AroundTime. Default: 90
        .PARAMETER ThrottleLimit
            Maximum concurrent machines to query. Default: 50
        .PARAMETER TimeoutMinutes
            Minutes before a machine's query task is auto-stopped. Default: 10
        .EXAMPLE
            Get-AcrobatLicenseDiagnostics -ComputerName "PC01"
        .EXAMPLE
            $list = Get-Content "$env:USERPROFILE\Desktop\Lists\Target_Machines.txt"
            Get-AcrobatLicenseDiagnostics -ComputerName $list -DaysBack 30
        .EXAMPLE
            # Analyze an exported Application log around a reported 10:20 AM failure
            Get-AcrobatLicenseDiagnostics -EvtxPath "$env:USERPROFILE\Desktop\AdobePro_Issue_Logs.evtx" -AroundTime '2026-07-29 10:20' -WindowMinutes 60
        .NOTES
            Written by Skyler Werner
            Date: 2026/07/30
            Version: 1.0.0

            Field notes on the failure pattern this hunts:
              - Legacy serial artifacts (SLStore/SLCache) or OperatingConfigs
                left behind by a pre-migration deployment make Acrobat attempt
                the old activation path first. When that entitlement is dead,
                Acrobat exits shortly after launch.
              - If NGL logs show LICENSE_EXPIRED / grace-period lines, the
                license-management system is intermittently reclaiming or
                failing to renew the user's entitlement.
              - If the licensing endpoints are unreachable on failure days,
                the cached token ages out and the refresh dies at the proxy.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Remote')]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ParameterSetName = 'Remote')]
        [string[]]
        $ComputerName,

        [Parameter(ParameterSetName = 'Remote')]
        [ValidateRange(1, 365)]
        [int]
        $DaysBack = 14,

        [Parameter(ParameterSetName = 'Remote')]
        [switch]
        $SkipEndpointTest,

        [Parameter(Mandatory, ParameterSetName = 'Evtx')]
        [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
        [string]
        $EvtxPath,

        [Parameter(ParameterSetName = 'Evtx')]
        [datetime]
        $AroundTime,

        [Parameter(ParameterSetName = 'Evtx')]
        [ValidateRange(5, 1440)]
        [int]
        $WindowMinutes = 90,

        [Parameter(ParameterSetName = 'Remote')]
        [ValidateRange(1, 300)]
        [int]
        $ThrottleLimit = 50,

        [Parameter(ParameterSetName = 'Remote')]
        [ValidateRange(1, 120)]
        [int]
        $TimeoutMinutes = 10,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        $collectedNames = @()
    }

    process {
        if ($PSCmdlet.ParameterSetName -eq 'Remote') {
            foreach ($name in $ComputerName) {
                if ($name.Length -gt 0) {
                    $collectedNames += $name
                }
            }
        }
    }

    end {

        # =============================================================
        #  Offline mode -- parse an exported .evtx
        # =============================================================
        if ($PSCmdlet.ParameterSetName -eq 'Evtx') {

            $resolved = (Resolve-Path -Path $EvtxPath).Path
            if ($resolved -notmatch '\.evtx$') {
                Write-Warning "File does not have an .evtx extension. If this is a mailed export renamed to .txt, rename it back to .evtx -- Get-WinEvent refuses other extensions."
            }

            Write-Host "Reading $resolved ..."
            try {
                $events = @(Get-WinEvent -Path $resolved -Oldest -ErrorAction Stop)
            }
            catch {
                Write-Warning "Failed to read event log file: $($_.Exception.Message)"
                return
            }

            # Crash/hang/WER IDs, everything from Adobe-named providers, and
            # anything whose message mentions Acrobat/Adobe.
            $relevant = @($events | Where-Object {
                ($_.Id -in 1000, 1001, 1002, 1026 -and "$($_.Message)" -match 'Acrobat|Adobe') -or
                ($_.ProviderName -match 'Adobe') -or
                ("$($_.Message)" -match 'Acrobat\.exe|AdobeGCClient|adobe_licensing|NGLClient')
            })

            Write-Host ""
            Write-Host "========================================"
            Write-Host "  Get-AcrobatLicenseDiagnostics -- EVTX"
            Write-Host "========================================"
            Write-Host ""
            Write-Host "  Events in file:       $($events.Count)"
            if ($events.Count -gt 0) {
                Write-Host "  File covers:          $($events[0].TimeCreated)  ->  $($events[-1].TimeCreated)"
            }
            Write-Host "  Acrobat/Adobe related: $($relevant.Count)"
            Write-Host ""

            $window = $relevant
            if ($PSBoundParameters.ContainsKey('AroundTime')) {
                $from   = $AroundTime.AddMinutes(-$WindowMinutes)
                $to     = $AroundTime.AddMinutes($WindowMinutes)
                $window = @($relevant | Where-Object { $_.TimeCreated -ge $from -and $_.TimeCreated -le $to })
                Write-Host "  Window $from -> $to :  $($window.Count) event(s)"
                Write-Host ""
            }

            if ($window.Count -eq 0) {
                Write-Warning "No Acrobat/Adobe-related events in the selected range. Widen -WindowMinutes or drop -AroundTime."
                if ($PassThru) { return $relevant }
                return
            }

            # Timeline table
            $timeline = @(foreach ($ev in $window) {
                $kind = switch ($ev.Id) {
                    1000    { 'CRASH (Application Error)' }
                    1001    { 'WER report' }
                    1002    { 'HANG (Application Hang)' }
                    1026    { '.NET Runtime error' }
                    default { "$($ev.ProviderName)" }
                }

                # Event 1000 message layout: faulting app, version, ts, faulting
                # module, version, ts, exception code, fault offset, pid ...
                $faultModule   = ''
                $exceptionCode = ''
                if ($ev.Id -eq 1000 -and $ev.Properties.Count -ge 7) {
                    $faultModule   = "$($ev.Properties[3].Value)"
                    $exceptionCode = "$($ev.Properties[6].Value)"
                }

                [PSCustomObject]@{
                    TimeCreated   = $ev.TimeCreated
                    Id            = $ev.Id
                    Kind          = $kind
                    FaultModule   = $faultModule
                    ExceptionCode = $exceptionCode
                    Message       = ("$($ev.Message)" -split "`r?`n")[0]
                }
            })

            $timeline | Format-Table TimeCreated, Id, Kind, FaultModule, ExceptionCode, Message -AutoSize -Wrap | Out-Host

            # Faulting-module rollup -- the fastest tell. A licensing kill shows
            # no 1000 at all (clean exit) or faults in Adobe licensing modules;
            # a plugin/render bug faults in annots.api, AcroForm.api, etc.
            $crashes = @($window | Where-Object { $_.Id -eq 1000 })
            Write-Host "========================================"
            Write-Host "  Interpretation"
            Write-Host "========================================"
            Write-Host ""
            if ($crashes.Count -eq 0) {
                Write-Host "  No Event 1000 crashes for Acrobat in this range. An exit with no"
                Write-Host "  crash event is the licensing-shutdown signature: Acrobat is told to"
                Write-Host "  terminate by its own licensing layer (NGL) and exits cleanly."
                Write-Host "  Next stop: NGLClient_*.log in the user's %TEMP% -- run this script's"
                Write-Host "  remote mode against the machine to pull the error lines."
            }
            else {
                $modules = @($crashes | ForEach-Object {
                    if ($_.Properties.Count -ge 4) { "$($_.Properties[3].Value)" }
                } | Group-Object | Sort-Object Count -Descending)
                Write-Host "  Faulting modules across $($crashes.Count) crash(es):"
                foreach ($m in $modules) {
                    Write-Host "    $($m.Name)  x$($m.Count)"
                }
                Write-Host ""
                Write-Host "  Licensing-adjacent modules (ngl*, adobe_licensing*, AdobeGC*) point at"
                Write-Host "  the entitlement path; document-handling modules (annots.api, *.api)"
                Write-Host "  point at content/plugin problems instead."
            }
            Write-Host ""

            if ($PassThru) { return $timeline }
            return
        }

        # =============================================================
        #  Remote mode
        # =============================================================

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

        if ($online.Count -eq 0) {
            Write-Warning "No machines responded to ping."
            return
        }

        # --- Build argument list ---
        $argList = $online | ForEach-Object { ,@($_, $DaysBack, [bool]$SkipEndpointTest) }

        # --- Remote scriptblock ---
        $scriptBlock = {
            $computer         = $args[0]
            $daysBack         = $args[1]
            $skipEndpointTest = $args[2]

            try {
                $remoteResult = Invoke-Command -ComputerName $computer -ArgumentList $daysBack, $skipEndpointTest -ScriptBlock {
                    param($daysBack, $skipEndpointTest)

                    $ErrorActionPreference = 'SilentlyContinue'

                    # --- Acrobat install/version ---
                    $acrobatVersion = ''
                    $uninstallPaths = @(
                        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
                        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
                    )
                    foreach ($path in $uninstallPaths) {
                        $entry = Get-ItemProperty -Path $path |
                            Where-Object { $_.DisplayName -match 'Adobe Acrobat' } |
                            Select-Object -First 1
                        if ($entry) {
                            $acrobatVersion = "$($entry.DisplayName) $($entry.DisplayVersion)"
                            break
                        }
                    }

                    # --- Crash / hang / WER events referencing Acrobat or Adobe ---
                    $since       = (Get-Date).AddDays(-$daysBack)
                    $crashLines  = @()
                    $crashCount  = 0
                    $lastCrash   = ''
                    $events = @(Get-WinEvent -FilterHashtable @{
                        LogName   = 'Application'
                        StartTime = $since
                    } -ErrorAction SilentlyContinue | Where-Object {
                        ($_.Id -in 1000, 1001, 1002, 1026 -and "$($_.Message)" -match 'Acrobat|Adobe') -or
                        ($_.ProviderName -match 'Adobe')
                    })
                    $crashCount = $events.Count
                    if ($crashCount -gt 0) {
                        $lastCrash = "$($events[0].TimeCreated)"
                        $crashLines = @($events | Select-Object -First 10 | ForEach-Object {
                            $firstLine = ("$($_.Message)" -split "`r?`n")[0]
                            "$($_.TimeCreated) [$($_.Id)] $firstLine"
                        })
                    }

                    # --- NGL client logs (per-user temp) ---
                    $nglErrorLines = @()
                    $nglLogsFound  = 0
                    $nglLogs = @(Get-ChildItem -Path 'C:\Users\*\AppData\Local\Temp\NGLClient_*.log' -ErrorAction SilentlyContinue |
                        Where-Object { $_.LastWriteTime -ge $since } |
                        Sort-Object LastWriteTime -Descending)
                    $nglLogsFound = $nglLogs.Count
                    foreach ($log in ($nglLogs | Select-Object -First 4)) {
                        $hits = @(Select-String -Path $log.FullName -Pattern 'ERROR|LICENSE_EXPIRED|GRACE|DENIED|NOT_ENTITLED|status\s*code\s*:\s*4|status\s*code\s*:\s*5' -ErrorAction SilentlyContinue |
                            Select-Object -Last 8)
                        foreach ($hit in $hits) {
                            $trimmed = $hit.Line.Trim()
                            if ($trimmed.Length -gt 220) { $trimmed = $trimmed.Substring(0, 220) }
                            $nglErrorLines += "$($log.Name): $trimmed"
                        }
                    }

                    # --- Leftover licensing artifacts from the pre-migration install ---
                    $slStore    = Test-Path 'C:\ProgramData\Adobe\SLStore'
                    $slCache    = Test-Path 'C:\Program Files (x86)\Common Files\Adobe\SLCache'
                    if (-not $slCache) {
                        $slCache = Test-Path 'C:\Program Files\Common Files\Adobe\SLCache'
                    }
                    $opConfigs  = @(Get-ChildItem -Path 'C:\ProgramData\Adobe\OperatingConfigs\*.operatingconfig' -ErrorAction SilentlyContinue)

                    # --- Adobe service state ---
                    $serviceStates = @()
                    foreach ($svcName in 'AGMService', 'AGSService', 'AdobeARMservice', 'AdobeUpdateService') {
                        $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
                        if ($svc) {
                            $serviceStates += "$svcName=$($svc.Status)"
                        }
                    }

                    # --- Licensing endpoint reachability (TCP 443) ---
                    $endpointFailures = @()
                    if (-not $skipEndpointTest) {
                        foreach ($endpoint in 'ims-na1.adobelogin.com', 'lcs-cops.adobe.io', 'lcs-ulecs.adobe.io', 'genuine.adobe.com') {
                            $client = New-Object System.Net.Sockets.TcpClient
                            try {
                                $async = $client.BeginConnect($endpoint, 443, $null, $null)
                                $ok    = $async.AsyncWaitHandle.WaitOne(5000, $false)
                                if (-not ($ok -and $client.Connected)) {
                                    $endpointFailures += $endpoint
                                }
                            }
                            catch {
                                $endpointFailures += $endpoint
                            }
                            finally {
                                $client.Close()
                            }
                        }
                    }

                    # Flat result -- survives the double-serialization boundary
                    [PSCustomObject]@{
                        AcrobatVersion   = $acrobatVersion
                        CrashCount       = [int]$crashCount
                        LastCrash        = $lastCrash
                        CrashLines       = [string[]]$crashLines
                        NglLogsFound     = [int]$nglLogsFound
                        NglErrorLines    = [string[]]$nglErrorLines
                        LegacySLStore    = [bool]$slStore
                        LegacySLCache    = [bool]$slCache
                        OperatingConfigs = [string[]]@($opConfigs | Select-Object -ExpandProperty Name)
                        ServiceStates    = [string[]]$serviceStates
                        EndpointFailures = [string[]]$endpointFailures
                    }
                } -ErrorAction Stop

                [PSCustomObject]@{
                    ComputerName     = $computer
                    Status           = 'Online'
                    AcrobatVersion   = $remoteResult.AcrobatVersion
                    CrashCount       = $remoteResult.CrashCount
                    LastCrash        = $remoteResult.LastCrash
                    CrashLines       = $remoteResult.CrashLines
                    NglLogsFound     = $remoteResult.NglLogsFound
                    NglErrorLines    = $remoteResult.NglErrorLines
                    LegacySLStore    = $remoteResult.LegacySLStore
                    LegacySLCache    = $remoteResult.LegacySLCache
                    OperatingConfigs = $remoteResult.OperatingConfigs
                    ServiceStates    = $remoteResult.ServiceStates
                    EndpointFailures = $remoteResult.EndpointFailures
                    Comment          = ''
                }
            }
            catch {
                [PSCustomObject]@{
                    ComputerName     = $computer
                    Status           = 'Online'
                    AcrobatVersion   = $null
                    CrashCount       = $null
                    LastCrash        = $null
                    CrashLines       = @()
                    NglLogsFound     = $null
                    NglErrorLines    = @()
                    LegacySLStore    = $null
                    LegacySLCache    = $null
                    OperatingConfigs = @()
                    ServiceStates    = @()
                    EndpointFailures = @()
                    Comment          = ($_.Exception.Message).Trim()
                }
            }
        }

        # --- Execute via RunspacePool ---
        $results = @(Invoke-RunspacePool `
            -ScriptBlock    $scriptBlock `
            -ArgumentList   $argList `
            -ThrottleLimit  $ThrottleLimit `
            -TimeoutMinutes $TimeoutMinutes `
            -ActivityName   "Get-AcrobatLicenseDiagnostics"
        )

        # --- Add offline machines to results ---
        foreach ($pc in $offline) {
            $results += [PSCustomObject]@{
                ComputerName     = $pc
                Status           = 'Offline'
                AcrobatVersion   = $null
                CrashCount       = $null
                LastCrash        = $null
                CrashLines       = @()
                NglLogsFound     = $null
                NglErrorLines    = @()
                LegacySLStore    = $null
                LegacySLCache    = $null
                OperatingConfigs = @()
                ServiceStates    = @()
                EndpointFailures = @()
                Comment          = ''
            }
        }

        # --- Summary output ---
        Write-Host ""
        Write-Host "========================================"
        Write-Host "  Get-AcrobatLicenseDiagnostics -- Results"
        Write-Host "========================================"
        Write-Host ""

        $results | Format-Table ComputerName, Status, AcrobatVersion, CrashCount, LastCrash,
            NglLogsFound, LegacySLStore, LegacySLCache, Comment -AutoSize | Out-Host

        # --- Per-machine findings ---
        foreach ($r in @($results | Where-Object { $_.Status -eq 'Online' -and -not $_.Comment })) {
            Write-Host "========================================"
            Write-Host "  $($r.ComputerName) -- Findings"
            Write-Host "========================================"

            $flagged = $false

            if ($r.LegacySLStore -or $r.LegacySLCache -or @($r.OperatingConfigs).Count -gt 0) {
                $flagged = $true
                Write-Host ""
                Write-Host "  [!] Leftover pre-migration licensing state detected:"
                if ($r.LegacySLStore) { Write-Host "        - SLStore present (legacy serial activation data)" }
                if ($r.LegacySLCache) { Write-Host "        - SLCache present (legacy serial activation data)" }
                foreach ($cfg in $r.OperatingConfigs) {
                    Write-Host "        - OperatingConfig: $cfg (FRL/SDL package license)"
                }
                Write-Host "      Acrobat tries these entitlements before the current one. A dead"
                Write-Host "      leftover entitlement produces exactly the flash-popup-then-exit"
                Write-Host "      pattern, intermittently. Clean with Adobe's licensing toolkit:"
                Write-Host "        adobe_licensing_toolkit.exe --deactivate   (legacy serial)"
                Write-Host "      then relaunch and re-sign-in under the current license system."
            }

            if (@($r.NglErrorLines).Count -gt 0) {
                $flagged = $true
                Write-Host ""
                Write-Host "  [!] NGL licensing errors found ($($r.NglLogsFound) log file(s) in range):"
                foreach ($line in ($r.NglErrorLines | Select-Object -First 12)) {
                    Write-Host "        $line"
                }
                Write-Host "      LICENSE_EXPIRED / NOT_ENTITLED lines mean the license system is"
                Write-Host "      reclaiming or failing to renew the seat -- check the user's"
                Write-Host "      assignment on the license-management side, not the workstation."
            }

            if (@($r.EndpointFailures).Count -gt 0) {
                $flagged = $true
                Write-Host ""
                Write-Host "  [!] Licensing endpoints unreachable on 443 from this machine:"
                foreach ($ep in $r.EndpointFailures) {
                    Write-Host "        - $ep"
                }
                Write-Host "      If the refresh can't reach these, the cached token eventually ages"
                Write-Host "      out and Acrobat shuts down after launch on exactly the bad days."
            }

            if ($r.CrashCount -gt 0) {
                Write-Host ""
                Write-Host "  [i] $($r.CrashCount) Acrobat/Adobe event(s) in the last window, most recent $($r.LastCrash):"
                foreach ($line in ($r.CrashLines | Select-Object -First 6)) {
                    Write-Host "        $line"
                }
            }
            elseif (-not $flagged) {
                Write-Host ""
                Write-Host "  [i] Nothing flagged. If the failure is intermittent, re-run on a bad"
                Write-Host "      day, or pull NGL logs immediately after a failure occurrence."
            }

            if ($r.ServiceStates.Count -gt 0) {
                Write-Host ""
                Write-Host "  [i] Adobe services: $($r.ServiceStates -join ', ')"
            }
            Write-Host ""
        }

        if ($PassThru) { return $results }
    }
}
