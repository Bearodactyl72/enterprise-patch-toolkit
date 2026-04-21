# DOTS formatting comment

function Get-StaleAsset {
    <#
        .SYNOPSIS
            Checks remote machines against AD, DNS, and DHCP to identify stale assets.

        .DESCRIPTION
            For each target machine, pings to determine reachability then queries Active
            Directory for the last logon date, password last set, and account status.
            Checks DNS for a valid forward lookup and optionally queries DHCP for active
            leases. Machines that have not checked into AD within the specified threshold
            are flagged as stale.

            Uses Invoke-RunspacePool for parallel execution across many machines.

            Data sources checked:
              - Ping: Online/Offline reachability
              - Active Directory: LastLogonDate, PasswordLastSet, Enabled, OperatingSystem
              - DNS: Forward lookup (A record) and IP match against AD
              - DHCP: Active lease lookup across discovered DHCP servers (optional)

        .PARAMETER ComputerName
            One or more computer names to check. Accepts pipeline input.

        .PARAMETER StaleDays
            Number of days since last AD logon to consider a machine stale.
            Default is 30.

        .PARAMETER IncludeDHCP
            When specified, queries DHCP servers for active leases. Requires the
            DhcpServer module (RSAT-DHCP feature).

        .PARAMETER ThrottleLimit
            Maximum concurrent runspaces. Default 25.

        .PARAMETER Timeout
            Timeout in minutes before a runspace is stopped. Default 5.

        .EXAMPLE
            Get-StaleAsset -ComputerName "PC001","PC002"

        .EXAMPLE
            Get-Content machines.txt | Get-StaleAsset -StaleDays 60

        .EXAMPLE
            Get-StaleAsset -ComputerName $TargetMachines -IncludeDHCP

        .NOTES
            Written by Skyler Werner
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [string[]]
        $ComputerName,

        [Parameter()]
        [ValidateRange(1, 365)]
        [int]
        $StaleDays = 30,

        [Parameter()]
        [switch]
        $IncludeDHCP,

        [Parameter()]
        [ValidateRange(1, 300)]
        [int]
        $ThrottleLimit = 25,

        [Parameter()]
        [Alias("TimeoutMinutes")]
        [ValidateRange(1, 60)]
        [int]
        $Timeout = 5,

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


    # --- Module availability checks ---

    $adAvailable = [bool](Get-Module -ListAvailable -Name ActiveDirectory -ErrorAction SilentlyContinue)
    if (-not $adAvailable) {
        Write-Warning "ActiveDirectory module is not installed. Install RSAT to use this script."
        return
    }

    $dhcpAvailable = $false
    $dhcpServers = @()
    if ($IncludeDHCP) {
        if (Get-Module -ListAvailable -Name DhcpServer -ErrorAction SilentlyContinue) {
            $dhcpAvailable = $true
            try {
                $dhcpServers = @(Get-DhcpServerInDC -ErrorAction Stop |
                    Where-Object { $_.DnsName -match 'fstr' })
                Write-Host "Found $($dhcpServers.Count) DHCP server(s)." -ForegroundColor DarkGray
            }
            catch {
                Write-Warning "DHCP server discovery failed: $_"
                $dhcpAvailable = $false
            }
        }
        else {
            Write-Warning "-IncludeDHCP requires the DhcpServer module (Install-WindowsFeature RSAT-DHCP)."
        }
    }


    # --- Pre-load DHCP leases into a hashtable for fast lookup ---

    $dhcpLeaseMap = @{}
    if ($dhcpAvailable -and $dhcpServers.Count -gt 0) {
        Write-Host "Loading DHCP leases (this may take a moment)..." -ForegroundColor DarkGray
        foreach ($dhcp in $dhcpServers) {
            $dhcpName = $dhcp.DnsName
            try {
                $scopes = @(Get-DhcpServerv4Scope -ComputerName $dhcpName -ErrorAction Stop)
                foreach ($scope in $scopes) {
                    $leases = @(Get-DhcpServerv4Lease -ComputerName $dhcpName `
                        -ScopeId $scope.ScopeId -ErrorAction Stop)
                    foreach ($lease in $leases) {
                        $leaseHost = ($lease.HostName -split '\.')[0].ToUpper()
                        if (-not $dhcpLeaseMap.ContainsKey($leaseHost)) {
                            $dhcpLeaseMap[$leaseHost] = $lease
                        }
                    }
                }
            }
            catch {
                Write-Warning "Failed to query DHCP server ${dhcpName}: $_"
            }
        }
    }


    # --- Connectivity check ---
    Write-Host ""
    Write-Host "Checking for online machines..."

    $pingResults = Test-ConnectionAsJob -ComputerName $targets

    # Build a lookup table -- ping status is passed into each runspace as a data point,
    # not used as a gate. AD, DNS, and DHCP queries hit infrastructure servers, not
    # the target machine, so offline machines still get full data.
    $pingMap = @{}
    $onlineCount = 0
    foreach ($ping in $pingResults) {
        $pingMap[$ping.ComputerName] = [bool]$ping.Reachable
        if ($ping.Reachable) { $onlineCount++ }
    }

    $allResults = [System.Collections.Generic.List[PSCustomObject]]::new()

    Write-Host "  Online: $onlineCount / $($targets.Count)" -ForegroundColor $(
        if ($onlineCount -eq $targets.Count) { 'Green' } else { 'DarkGray' }
    )


    # --- Scriptblock executed in each runspace (one per machine) ---
    $staleCheckBlock = {

        $computer       = $args[0]
        $staleDays      = $args[1]
        $checkDHCP      = $args[2]
        $dhcpLeaseData  = $args[3]
        $isOnline       = $args[4]

        $PhaseTracker[$computer] = "Querying AD"

        $staleThreshold = (Get-Date).AddDays(-$staleDays)

        # -- AD lookup --
        $adFound      = $false
        $lastLogon    = $null
        $pwdLastSet   = $null
        $enabled      = $null
        $os           = $null
        $adIp         = $null
        $adComment    = $null

        try {
            $adObj = Get-ADComputer $computer -Properties `
                LastLogonDate, PasswordLastSet, Enabled, OperatingSystem, IPv4Address `
                -ErrorAction Stop

            $adFound    = $true
            $lastLogon  = $adObj.LastLogonDate
            $pwdLastSet = $adObj.PasswordLastSet
            $enabled    = $adObj.Enabled
            $os         = $adObj.OperatingSystem
            $adIp       = $adObj.IPv4Address
        }
        catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            $adComment = "Not found in AD"
        }
        catch {
            $adComment = "AD query failed: $_"
        }

        # -- Staleness calculation --
        # LastLogonDate (replicated ~14 days) can be $null on active machines.
        # Fall back to PasswordLastSet (rotates every 30 days) as proof of life.
        $effectiveDate = if ($lastLogon) { $lastLogon } elseif ($pwdLastSet) { $pwdLastSet } else { $null }

        $daysSinceLogon = $null
        if ($effectiveDate) {
            $daysSinceLogon = [math]::Round(((Get-Date) - $effectiveDate).TotalDays)
        }

        $adStatus = if (-not $adFound) {
            "Not Found"
        }
        elseif ($enabled -eq $false) {
            "Disabled"
        }
        elseif ($null -eq $effectiveDate) {
            "Never Logged In"
        }
        elseif ($effectiveDate -lt $staleThreshold) {
            "Stale"
        }
        else {
            "Current"
        }

        if (-not $lastLogon -and $pwdLastSet) {
            $adComment = "LastLogonDate empty -- used PasswordLastSet"
        }

        # -- DNS lookup --
        $PhaseTracker[$computer] = "Querying DNS"

        $dnsIp      = $null
        $dnsMatch   = $null
        $dnsComment = $null

        try {
            $dnsResult = Resolve-DnsName $computer -Type A -DnsOnly -ErrorAction Stop
            $dnsIp = @($dnsResult | Where-Object { $_.Type -eq 'A' } |
                Select-Object -First 1).IPAddress

            if ($dnsIp -and $adIp) {
                $dnsMatch = ($dnsIp -eq $adIp)
                if (-not $dnsMatch) {
                    $dnsComment = "DNS=$dnsIp vs AD=$adIp"
                }
            }
        }
        catch {
            $dnsComment = "No DNS record"
        }

        # -- DHCP lookup (from pre-loaded data) --
        $dhcpIp     = $null
        $dhcpExpiry = $null
        if ($checkDHCP -and $null -ne $dhcpLeaseData) {
            $lease = $dhcpLeaseData[$computer]
            if ($null -ne $lease) {
                $dhcpIp     = "$($lease.IPAddress)"
                $dhcpExpiry = $lease.LeaseExpiryTime
            }
        }

        # -- Build combined comment --
        $comment = if ($adComment) { $adComment }
            elseif ($dnsComment) { $dnsComment }
            else { $null }

        $PhaseTracker[$computer] = "Complete"

        [PSCustomObject][ordered]@{
            ComputerName    = $computer
            Status          = if ($isOnline) { "Online" } else { "Offline" }
            ADStatus        = $adStatus
            Enabled         = $enabled
            LastLogon       = $lastLogon
            DaysSinceLogon  = $daysSinceLogon
            PasswordLastSet = $pwdLastSet
            OS              = $os
            ADIPAddress     = $adIp
            DNSIPAddress    = $dnsIp
            DNSMatchesAD    = $dnsMatch
            DHCPIPAddress   = $dhcpIp
            DHCPLeaseExpiry = $dhcpExpiry
            Comment         = $comment
        }
    }


    # --- Build argument sets (one per machine -- ALL machines, not just online) ---
    $argumentSets = @(
        foreach ($machine in $targets) {
            , @($machine, $StaleDays, [bool]$IncludeDHCP, $dhcpLeaseMap, $pingMap[$machine])
        }
    )


    # --- Execute via Invoke-RunspacePool ---
    $runspaceParams = @{
        ScriptBlock    = $staleCheckBlock
        ArgumentList   = $argumentSets
        ThrottleLimit  = $ThrottleLimit
        TimeoutMinutes = $Timeout
        ActivityName   = "Get Stale Assets"
    }

    $poolResults = Invoke-RunspacePool @runspaceParams


    # --- Timeout guard: normalize incomplete results to full-width objects ---
    foreach ($result in $poolResults) {
        if ($result -isnot [PSCustomObject]) { continue }

        if ($null -ne $result.PSObject.Properties['ADStatus']) {
            $allResults.Add($result)
            continue
        }

        $allResults.Add([PSCustomObject][ordered]@{
            ComputerName    = $result.ComputerName
            Status          = if ($pingMap[$result.ComputerName]) { "Online" } else { "Offline" }
            ADStatus        = $null
            Enabled         = $null
            LastLogon       = $null
            DaysSinceLogon  = $null
            PasswordLastSet = $null
            OS              = $null
            ADIPAddress     = $null
            DNSIPAddress    = $null
            DNSMatchesAD    = $null
            DHCPIPAddress   = $null
            DHCPLeaseExpiry = $null
            Comment         = if ($null -ne $result.Comment) { $result.Comment } else { "Task Stopped" }
        })
    }


    # --- Sort: Online first (offline last), then by days since logon descending ---
    $sorted = @($allResults | Sort-Object -Property @(
        @{ Expression = "Status";         Descending = $true  }
        @{ Expression = "DaysSinceLogon"; Descending = $true  }
        @{ Expression = "ComputerName";   Descending = $false }
    ))


    # --- Summary ---
    $staleCount    = @($sorted | Where-Object { $_.ADStatus -eq 'Stale' }).Count
    $disabledCount = @($sorted | Where-Object { $_.ADStatus -eq 'Disabled' }).Count
    $notFoundCount = @($sorted | Where-Object { $_.ADStatus -eq 'Not Found' }).Count
    $neverLogon    = @($sorted | Where-Object { $_.ADStatus -eq 'Never Logged In' }).Count
    $currentCount  = @($sorted | Where-Object { $_.ADStatus -eq 'Current' }).Count
    $offlineCount  = @($sorted | Where-Object { $_.Status -eq 'Offline' }).Count
    $dnsMismatch   = @($sorted | Where-Object { $_.DNSMatchesAD -eq $false }).Count

    Write-Host ""
    Write-Host "--- Stale Asset Summary ---" -ForegroundColor Cyan
    Write-Host "  Total checked:     $($sorted.Count)"
    Write-Host "  Online:            $onlineCount" -ForegroundColor Green
    Write-Host "  Offline:           $offlineCount" -ForegroundColor $(if ($offlineCount -gt 0) { 'DarkGray' } else { 'Green' })
    Write-Host ""
    Write-Host "  AD Current:        $currentCount" -ForegroundColor Green
    Write-Host "  AD Stale (>$StaleDays d):  $staleCount" -ForegroundColor $(if ($staleCount -gt 0) { 'Yellow' } else { 'Green' })
    Write-Host "  AD Disabled:       $disabledCount" -ForegroundColor $(if ($disabledCount -gt 0) { 'Yellow' } else { 'Green' })
    Write-Host "  AD Not Found:      $notFoundCount" -ForegroundColor $(if ($notFoundCount -gt 0) { 'Red' } else { 'Green' })
    if ($neverLogon -gt 0) {
        Write-Host "  AD Never Logged In: $neverLogon" -ForegroundColor Yellow
    }
    if ($dnsMismatch -gt 0) {
        Write-Host "  DNS Mismatch:      $dnsMismatch" -ForegroundColor Red
    }
    Write-Host ""


    # --- Output results ---
    $sorted | Format-Table -AutoSize | Out-Host

    if ($PassThru) { return $sorted }

    } # end
}
