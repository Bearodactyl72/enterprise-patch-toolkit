# DOTS formatting comment

<#
    .SYNOPSIS
        Diagnostic script to test which remote access methods work for machines
        that are pingable but unreachable via WinRM.

    .DESCRIPTION
        For each target machine, probes multiple name resolution and connectivity
        methods to determine what works and what does not. Designed to help identify
        viable fix paths for machines with stale DNS or broken WinRM.

        Tests performed:
          - IP resolution via DNS, NetBIOS (WINS), and Active Directory
          - Ping to each discovered IP
          - WinRM (Test-WSMan and Invoke-Command)
          - WMI over DCOM (Get-WmiObject and Invoke-WmiMethod)
          - SMB admin share (Test-Path \\computer\C$)
          - Remote scheduled task query (schtasks /query)
          - DHCP server discovery and lease lookup
          - Hosts file override workaround (opt-in via -TestHostsFile)
          - PsExec backdoor feasibility (auto-discovered or via -PsExecPath)
          - Remote diagnostics via PsExec: clock sync, secure channel, WinRM service

    .PARAMETER ComputerName
        One or more target machine names to test. Use machines that are known
        to exhibit the WinRM error for meaningful results.

    .PARAMETER TestHostsFile
        When specified, tests whether temporarily adding a hosts file entry
        (mapping hostname to the AD-resolved IP) restores WinRM access.
        Requires admin elevation. The entry is removed after testing.

    .PARAMETER PsExecPath
        Full path to PsExec.exe. If omitted, the script searches common locations
        (M:\Regional\VMT\Scripts\PSTools\, C:\PSTools\, PATH).

    .EXAMPLE
        .\Test-RemoteAccess.ps1 -ComputerName "PC001"

    .EXAMPLE
        .\Test-RemoteAccess.ps1 -ComputerName "PC001","PC002" -TestHostsFile

    .EXAMPLE
        .\Test-RemoteAccess.ps1 -ComputerName $TargetMachines -PsExecPath "D:\Tools\PsExec.exe"

    .NOTES
        Written by Skyler Werner
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory, Position = 0)]
    [string[]]
    $ComputerName,

    [Parameter()]
    [switch]
    $TestHostsFile,

    [Parameter()]
    [string]
    $PsExecPath
)


# -----------------------------------------------------------------------
# Helper functions
# -----------------------------------------------------------------------

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host "--- $Title ---" -ForegroundColor Cyan
}

function Write-TestResult {
    param(
        [string]$Label,
        [string]$Result,
        [string]$Detail
    )

    $padded = ($Label + ":").PadRight(22)

    switch -Wildcard ($Result) {
        "OK"       { $color = "Green"  }
        "FAILED*"  { $color = "Red"    }
        "SKIPPED*" { $color = "DarkGray" }
        "PARTIAL*" { $color = "Yellow" }
        default    { $color = "White"  }
    }

    Write-Host "  $padded " -NoNewline
    Write-Host $Result -ForegroundColor $color -NoNewline
    if ($Detail) {
        Write-Host "  ($Detail)" -ForegroundColor DarkGray
    }
    else {
        Write-Host ""
    }
}

function Resolve-NetBIOS {
    param([string]$Name)

    $output = & nbtstat -a $Name 2>&1
    foreach ($line in $output) {
        if ($line -match 'Node IpAddress:\s*\[(.+?)\]') {
            return $Matches[1]
        }
    }
    return $null
}

function Test-Ping {
    param([string]$Target)
    $result = Test-Connection -ComputerName $Target -Count 1 -Quiet -ErrorAction SilentlyContinue
    return [bool]$result
}

function Invoke-PsExecCommand {
    <#
        Runs a command on a remote machine via PsExec (as SYSTEM, elevated).
        Returns a PSCustomObject with Output (cleaned string), ExitCode, and Success bool.
        PsExec writes banner/status to stderr; this function filters that noise out.
    #>
    param(
        [string]$ExePath,
        [string]$Target,
        [string[]]$ArgumentList
    )

    $raw = & $ExePath -accepteula -h -s "\\$Target" @ArgumentList 2>&1
    $exitCode = $LASTEXITCODE

    # PsExec stderr lines arrive as ErrorRecord objects when captured with 2>&1.
    # Convert everything to strings, then filter out PsExec's own output.
    $lines = @(foreach ($item in $raw) {
        $str = if ($item -is [System.Management.Automation.ErrorRecord]) {
            $item.ToString()
        } else { "$item" }

        if ($str -notmatch '^PsExec' -and
            $str -notmatch '^Copyright' -and
            $str -notmatch '^Sysinternals' -and
            $str -notmatch 'exited on .+ with error code' -and
            $str -notmatch '^\s*Connecting to' -and
            $str -notmatch '^\s*Connecting with' -and
            $str -notmatch '^\s*Starting .+ on' -and
            $str -notmatch 'System\.Management\.Automation' -and
            $str.Trim() -ne '') {
            $str
        }
    })

    [PSCustomObject]@{
        Output   = ($lines -join "`n").Trim()
        ExitCode = $exitCode
        Success  = ($exitCode -eq 0)
    }
}


# -----------------------------------------------------------------------
# Elevation check (needed for hosts file test)
# -----------------------------------------------------------------------

$isElevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if ($TestHostsFile -and -not $isElevated) {
    Write-Warning "-TestHostsFile requires admin elevation. That test will be skipped."
    $TestHostsFile = $false
}


# -----------------------------------------------------------------------
# DHCP server discovery (one-time)
# -----------------------------------------------------------------------

Write-Host ""
Write-Host "========================================" -ForegroundColor White
Write-Host " Remote Access Diagnostic Tool"         -ForegroundColor White
Write-Host "========================================" -ForegroundColor White

Write-Section "Environment Probes"

# Test DhcpServer module
$dhcpModuleAvailable = $false
$dhcpServers = @()

if (Get-Module -ListAvailable -Name DhcpServer -ErrorAction SilentlyContinue) {
    Write-TestResult "DhcpServer Module" "OK" "Module is installed"
    $dhcpModuleAvailable = $true

    try {
        $dhcpServers = @(Get-DhcpServerInDC -ErrorAction Stop | Where-Object { $_.DnsName -match 'fstr' })
        Write-TestResult "DHCP Server Discovery" "OK" "$($dhcpServers.Count) server(s) found"
    }
    catch {
        Write-TestResult "DHCP Server Discovery" "FAILED" "$_"
    }
}
else {
    Write-TestResult "DhcpServer Module" "FAILED" "Not installed (Install-WindowsFeature RSAT-DHCP or RSAT)"
}

# Test ActiveDirectory module
$adModuleAvailable = $false
if (Get-Module -ListAvailable -Name ActiveDirectory -ErrorAction SilentlyContinue) {
    Write-TestResult "AD Module" "OK" "Module is installed"
    $adModuleAvailable = $true
}
else {
    Write-TestResult "AD Module" "FAILED" "Not installed"
}

# PsExec discovery
$psExec = $null
if ($PsExecPath -and (Test-Path $PsExecPath)) {
    $psExec = $PsExecPath
}
else {
    $searchPaths = @(
        "M:\Regional\VMT\Scripts\PSTools\PsExec.exe"
        "$env:SystemDrive\PSTools\PsExec.exe"
        "$env:SystemDrive\SysinternalsSuite\PsExec.exe"
        "$env:ProgramFiles\PSTools\PsExec.exe"
    )
    foreach ($p in $searchPaths) {
        if (Test-Path $p) { $psExec = $p; break }
    }
    if (-not $psExec) {
        $found = Get-Command PsExec.exe -ErrorAction SilentlyContinue
        if ($found) { $psExec = $found.Source }
    }
}
if ($psExec) {
    Write-TestResult "PsExec" "OK" $psExec
}
else {
    Write-TestResult "PsExec" "SKIPPED" "Not found (use -PsExecPath to specify)"
}


# -----------------------------------------------------------------------
# Per-machine testing
# -----------------------------------------------------------------------

foreach ($computer in $ComputerName) {

    Write-Host ""
    Write-Host "========================================" -ForegroundColor White
    Write-Host " Testing: $computer"                      -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor White


    # === IP Resolution ===
    Write-Section "IP Resolution"

    # DNS
    $dnsIp = $null
    try {
        $dnsResult = Resolve-DnsName $computer -Type A -DnsOnly -ErrorAction Stop
        $dnsIp = @($dnsResult | Where-Object { $_.Type -eq 'A' } |
            Select-Object -First 1).IPAddress
        Write-TestResult "DNS (A Record)" "OK" $dnsIp
    }
    catch {
        Write-TestResult "DNS (A Record)" "FAILED" "$_"
    }

    # NetBIOS / WINS
    $nbtIp = Resolve-NetBIOS $computer
    if ($nbtIp) {
        Write-TestResult "NetBIOS (WINS)" "OK" $nbtIp
    }
    else {
        Write-TestResult "NetBIOS (WINS)" "FAILED" "No response or WINS unavailable"
    }

    # Active Directory
    $adIp = $null
    if ($adModuleAvailable) {
        try {
            $adObj = Get-ADComputer $computer -Properties IPv4Address -ErrorAction Stop
            $adIp = $adObj.IPv4Address
            if ($adIp) {
                Write-TestResult "AD IPv4Address" "OK" $adIp
            }
            else {
                Write-TestResult "AD IPv4Address" "PARTIAL" "Computer found but no IPv4Address attribute"
            }
        }
        catch {
            Write-TestResult "AD IPv4Address" "FAILED" "$_"
        }
    }
    else {
        Write-TestResult "AD IPv4Address" "SKIPPED" "AD module not available"
    }

    # Compare -- use AD as ground truth since WINS entries can be stale
    $ipMismatch = $false
    if ($adIp) {
        if ($dnsIp -and ($dnsIp -ne $adIp)) {
            $ipMismatch = $true
            Write-Host ""
            Write-Host "  ** STALE DNS: DNS=$dnsIp  AD=$adIp **" -ForegroundColor Red
        }
        elseif ($dnsIp -and $dnsIp -eq $adIp) {
            Write-Host ""
            Write-Host "  DNS and AD agree: $dnsIp" -ForegroundColor Green
        }
        if ($nbtIp -and ($nbtIp -ne $adIp)) {
            Write-Host "  ** WINS IP ($nbtIp) differs from AD IP ($adIp) -- WINS entry is stale." -ForegroundColor Yellow
        }
    }


    # === Ping Tests ===
    Write-Section "Ping Tests"

    if ($dnsIp) {
        $pingDns = Test-Ping $dnsIp
        Write-TestResult "Ping DNS IP ($dnsIp)" $(if ($pingDns) { "OK" } else { "FAILED" })
    }

    if ($nbtIp -and $nbtIp -ne $dnsIp) {
        $pingNbt = Test-Ping $nbtIp
        Write-TestResult "Ping NBT IP ($nbtIp)" $(if ($pingNbt) { "OK" } else { "FAILED" })
    }

    # Ping by name -- also capture the resolved IP (what the OS resolver actually uses)
    $resolvedIp = $null
    $pingNameResult = Test-Connection -ComputerName $computer -Count 1 -ErrorAction SilentlyContinue
    if ($pingNameResult) {
        $resolvedIp = $pingNameResult.IPV4Address.IPAddressToString
        Write-TestResult "Ping by Name" "OK" "Resolved to $resolvedIp"

        # Flag if the resolved IP differs from AD -- means ping is hitting a different machine
        if ($adIp -and $resolvedIp -ne $adIp) {
            Write-Host "  ** Resolved IP ($resolvedIp) differs from AD IP ($adIp) **" -ForegroundColor Yellow
        }
    }
    else {
        Write-TestResult "Ping by Name" "FAILED"
    }


    # === Connectivity via Hostname ===
    Write-Section "Connectivity via Hostname ($computer)"

    # WinRM - Test-WSMan (lightweight check)
    try {
        $wsman = Test-WSMan -ComputerName $computer -ErrorAction Stop
        Write-TestResult "WinRM (Test-WSMan)" "OK"
    }
    catch {
        $wsmanErr = $_.Exception.Message
        # Truncate long WinRM errors
        if ($wsmanErr.Length -gt 80) { $wsmanErr = $wsmanErr.Substring(0, 80) + "..." }
        Write-TestResult "WinRM (Test-WSMan)" "FAILED" $wsmanErr
    }

    # WinRM - Invoke-Command (full test)
    $winrmByNameOk  = $false
    $winrmByNameErr = $null
    try {
        $icmResult = Invoke-Command -ComputerName $computer -ErrorAction Stop -ScriptBlock {
            $env:COMPUTERNAME
        }
        $winrmByNameOk = $true
        Write-TestResult "WinRM (Invoke-Cmd)" "OK" "Returned: $icmResult"
    }
    catch {
        $winrmByNameErr = $_.Exception.Message
        $icmErr = $winrmByNameErr
        if ($icmErr.Length -gt 80) { $icmErr = $icmErr.Substring(0, 80) + "..." }
        Write-TestResult "WinRM (Invoke-Cmd)" "FAILED" $icmErr
    }

    # WMI over DCOM
    try {
        $wmiResult = Get-WmiObject Win32_OperatingSystem -ComputerName $computer -ErrorAction Stop
        Write-TestResult "WMI/DCOM" "OK" $wmiResult.Caption
    }
    catch {
        $wmiErr = $_.Exception.Message
        if ($wmiErr.Length -gt 80) { $wmiErr = $wmiErr.Substring(0, 80) + "..." }
        Write-TestResult "WMI/DCOM" "FAILED" $wmiErr
    }

    # SMB admin share
    $smbPath = "\\$computer\C$"
    try {
        $smbOk = Test-Path $smbPath -ErrorAction Stop
        Write-TestResult "SMB (C$ share)" $(if ($smbOk) { "OK" } else { "FAILED" }) $smbPath
    }
    catch {
        $smbErr = $_.Exception.Message
        if ($smbErr.Length -gt 80) { $smbErr = $smbErr.Substring(0, 80) + "..." }
        Write-TestResult "SMB (C$ share)" "FAILED" $smbErr
    }

    # Remote schtasks query (RPC)
    $schtasksOutput = & schtasks /query /s $computer /tn "\Microsoft\Windows\Time Synchronization\SynchronizeTime" 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-TestResult "RPC (schtasks)" "OK"
    }
    else {
        $stErr = ($schtasksOutput | Select-Object -First 1) -as [string]
        if ($stErr.Length -gt 80) { $stErr = $stErr.Substring(0, 80) + "..." }
        Write-TestResult "RPC (schtasks)" "FAILED" $stErr
    }


    # === Connectivity via Raw IP ===
    # Prefer AD IP as ground truth; fall back to WINS only if AD unavailable.
    # WINS entries can be stale and point to wrong machines.
    $altIp = $null
    if ($adIp) {
        $altIp = $adIp
    }
    elseif ($nbtIp -and $nbtIp -ne $dnsIp) {
        $altIp = $nbtIp
    }

    $winrmByIpOk = $false

    if ($altIp) {
        Write-Section "Connectivity via Raw IP ($altIp)"

        # WinRM via IP -- Kerberos won't work with a raw IP; falls back to NTLM
        try {
            $null = Test-WSMan -ComputerName $altIp -ErrorAction Stop
            $winrmByIpOk = $true
            Write-TestResult "WinRM (IP)" "OK" "NTLM (Kerberos requires hostname)"
        }
        catch {
            $wsman2Err = $_.Exception.Message
            if ($wsman2Err.Length -gt 80) { $wsman2Err = $wsman2Err.Substring(0, 80) + "..." }
            Write-TestResult "WinRM (IP)" "FAILED" $wsman2Err
        }

        # WMI via IP
        try {
            $wmi2 = Get-WmiObject Win32_OperatingSystem -ComputerName $altIp -ErrorAction Stop
            Write-TestResult "WMI/DCOM (IP)" "OK" $wmi2.Caption
        }
        catch {
            $wmi2Err = $_.Exception.Message
            if ($wmi2Err.Length -gt 80) { $wmi2Err = $wmi2Err.Substring(0, 80) + "..." }
            Write-TestResult "WMI/DCOM (IP)" "FAILED" $wmi2Err
        }

        # SMB via IP
        $smbPathIp = "\\$altIp\C$"
        try {
            $smb2 = Test-Path $smbPathIp -ErrorAction Stop
            Write-TestResult "SMB (IP)" $(if ($smb2) { "OK" } else { "FAILED" }) $smbPathIp
        }
        catch {
            $smb2Err = $_.Exception.Message
            if ($smb2Err.Length -gt 80) { $smb2Err = $smb2Err.Substring(0, 80) + "..." }
            Write-TestResult "SMB (IP)" "FAILED" $smb2Err
        }

        # schtasks via IP
        $st2Output = & schtasks /query /s $altIp /tn "\Microsoft\Windows\Time Synchronization\SynchronizeTime" 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-TestResult "RPC/schtasks (IP)" "OK"
        }
        else {
            $st2Err = ($st2Output | Select-Object -First 1) -as [string]
            if ($st2Err.Length -gt 80) { $st2Err = $st2Err.Substring(0, 80) + "..." }
            Write-TestResult "RPC/schtasks (IP)" "FAILED" $st2Err
        }
    }


    # === Remediation Feasibility ===
    # Only runs for machines where WinRM by hostname failed and we have an AD IP.
    # Tests whether PsExec and WMI can reach the machine via IP -- these are
    # the backdoor methods Repair-RemoteAccess.ps1 would use to fix the machine.

    $psExecToIpOk      = $false
    $psExecHostMatch   = $false
    $wmiExecToIpOk     = $false
    $secureChannelOk   = $null
    $winrmServiceUp    = $null
    $clockSyncInfo     = $null

    if (-not $winrmByNameOk -and $adIp) {
        Write-Section "Remediation Feasibility (via AD IP: $adIp)"

        # -- WMI process creation (DCOM backdoor, no PsExec needed) --
        try {
            $wmiTest = Invoke-WmiMethod -ComputerName $adIp -Class Win32_Process `
                -Name Create -ArgumentList "cmd /c echo test > NUL" -ErrorAction Stop
            if ($wmiTest.ReturnValue -eq 0) {
                $wmiExecToIpOk = $true
                Write-TestResult "WMI RemoteExec (IP)" "OK" "Can create processes via DCOM"
            }
            else {
                Write-TestResult "WMI RemoteExec (IP)" "FAILED" "ReturnValue=$($wmiTest.ReturnValue)"
            }
        }
        catch {
            $wmiExecErr = $_.Exception.Message
            if ($wmiExecErr.Length -gt 80) { $wmiExecErr = $wmiExecErr.Substring(0, 80) + "..." }
            Write-TestResult "WMI RemoteExec (IP)" "FAILED" $wmiExecErr
        }

        # -- PsExec backdoor tests --
        if ($psExec) {

            # Test 1: Basic connectivity -- does PsExec reach the right machine?
            $psResult = Invoke-PsExecCommand -ExePath $psExec -Target $adIp -ArgumentList @("hostname")
            if ($psResult.Success) {
                $psExecToIpOk = $true
                $remoteHost = $psResult.Output.Trim()
                Write-TestResult "PsExec (hostname)" "OK" "Returned: $remoteHost"

                # Verify we are talking to the expected machine
                if ($remoteHost -eq $computer) {
                    $psExecHostMatch = $true
                }
                else {
                    Write-Host "  ** WARNING: Expected $computer but got $remoteHost -- AD IP may point to wrong host **" -ForegroundColor Red
                }
            }
            else {
                $psErr = $psResult.Output
                if ($psErr.Length -gt 80) { $psErr = $psErr.Substring(0, 80) + "..." }
                Write-TestResult "PsExec (hostname)" "FAILED" "Exit code $($psResult.ExitCode) -- $psErr"
            }

            # Only run deeper diagnostics if PsExec works and hostname matches
            if ($psExecToIpOk -and $psExecHostMatch) {

                # Test 2: Clock synchronization
                $clockResult = Invoke-PsExecCommand -ExePath $psExec -Target $adIp `
                    -ArgumentList @("w32tm", "/query", "/status")
                if ($clockResult.Success) {
                    $lastSync = ($clockResult.Output -split "`n" |
                        Where-Object { $_ -match 'Last Successful Sync' }) -replace 'Last Successful Sync Time:\s*', ''
                    $timeSource = ($clockResult.Output -split "`n" |
                        Where-Object { $_ -match '^Source:' }) -replace 'Source:\s*', ''
                    $clockSyncInfo = if ($lastSync) { $lastSync.Trim() } else { "Unknown" }
                    Write-TestResult "Clock (w32tm)" "OK" "Last sync: $clockSyncInfo"
                    if ($timeSource) {
                        Write-Host "                         Time source: $($timeSource.Trim())" -ForegroundColor DarkGray
                    }
                }
                else {
                    $clockSyncInfo = "FAILED"
                    $clockErr = $clockResult.Output
                    if ($clockErr.Length -gt 80) { $clockErr = $clockErr.Substring(0, 80) + "..." }
                    Write-TestResult "Clock (w32tm)" "FAILED" $clockErr
                }

                # Test 3: Secure channel (domain trust)
                $domain = $env:USERDNSDOMAIN
                $scResult = Invoke-PsExecCommand -ExePath $psExec -Target $adIp `
                    -ArgumentList @("nltest", "/sc_query:$domain")
                if ($scResult.Output -match 'NERR_Success') {
                    $secureChannelOk = $true
                    Write-TestResult "Secure Channel" "OK" "Domain trust is healthy"
                }
                else {
                    $secureChannelOk = $false
                    # Extract the status line from nltest output
                    $scStatus = ($scResult.Output -split "`n" |
                        Where-Object { $_ -match 'Status' -or $_ -match 'ERROR' } |
                        Select-Object -First 1)
                    if (-not $scStatus) { $scStatus = "nltest returned exit code $($scResult.ExitCode)" }
                    if ($scStatus.Length -gt 80) { $scStatus = $scStatus.Substring(0, 80) + "..." }
                    Write-TestResult "Secure Channel" "FAILED" $scStatus
                }

                # Test 4: WinRM service state on the remote machine
                $svcResult = Invoke-PsExecCommand -ExePath $psExec -Target $adIp `
                    -ArgumentList @("sc", "query", "WinRM")
                if ($svcResult.Output -match 'RUNNING') {
                    $winrmServiceUp = $true
                    Write-TestResult "WinRM Service" "OK" "Running on remote machine"
                }
                elseif ($svcResult.Output -match 'STOPPED') {
                    $winrmServiceUp = $false
                    Write-TestResult "WinRM Service" "FAILED" "Service is STOPPED"
                }
                else {
                    $winrmServiceUp = $false
                    Write-TestResult "WinRM Service" "FAILED" "Unable to determine state"
                }
            }
        }
        else {
            Write-TestResult "PsExec Diagnostics" "SKIPPED" "PsExec not found"
        }
    }


    # === Hosts File Override Test ===
    if ($TestHostsFile -and $altIp) {
        Write-Section "Hosts File Override Test"

        $hostsPath  = "$env:SystemRoot\System32\drivers\etc\hosts"
        $hostsEntry = "$altIp`t$computer`t# Test-RemoteAccess temp entry"

        Write-Host "  Adding temporary hosts entry: $altIp -> $computer"

        try {
            # Backup and add entry
            $hostsContent = Get-Content $hostsPath -Raw -ErrorAction Stop
            Add-Content -Path $hostsPath -Value $hostsEntry -ErrorAction Stop

            # Flush local DNS cache so the hosts entry takes effect
            & ipconfig /flushdns 2>&1 | Out-Null

            # Test WinRM with hosts override active
            try {
                $hostsWsman = Test-WSMan -ComputerName $computer -ErrorAction Stop
                Write-TestResult "WinRM (hosts)" "OK" "Hosts file workaround WORKS"
            }
            catch {
                $hwErr = $_.Exception.Message
                if ($hwErr.Length -gt 80) { $hwErr = $hwErr.Substring(0, 80) + "..." }
                Write-TestResult "WinRM (hosts)" "FAILED" $hwErr
            }

            # Test Invoke-Command with hosts override
            try {
                $hostsIcm = Invoke-Command -ComputerName $computer -ErrorAction Stop -ScriptBlock {
                    $env:COMPUTERNAME
                }
                Write-TestResult "Invoke-Cmd (hosts)" "OK" "Returned: $hostsIcm"
            }
            catch {
                $hiErr = $_.Exception.Message
                if ($hiErr.Length -gt 80) { $hiErr = $hiErr.Substring(0, 80) + "..." }
                Write-TestResult "Invoke-Cmd (hosts)" "FAILED" $hiErr
            }
        }
        catch {
            Write-TestResult "Hosts File Edit" "FAILED" "$_"
        }
        finally {
            # Always clean up the hosts file
            Write-Host "  Removing temporary hosts entry..."
            try {
                $cleanLines = @(Get-Content $hostsPath -ErrorAction Stop) |
                    Where-Object { $_ -notmatch '# Test-RemoteAccess temp entry$' }
                $cleanLines | Set-Content $hostsPath -ErrorAction Stop
                & ipconfig /flushdns 2>&1 | Out-Null
                Write-Host "  Hosts file restored." -ForegroundColor Green
            }
            catch {
                Write-Host "  WARNING: Failed to clean up hosts file! Manual cleanup required." -ForegroundColor Red
                Write-Host "  Remove the line containing '# Test-RemoteAccess temp entry' from:" -ForegroundColor Red
                Write-Host "  $hostsPath" -ForegroundColor Red
            }
        }
    }
    elseif ($TestHostsFile -and -not $altIp) {
        Write-Section "Hosts File Override Test"
        Write-TestResult "Hosts Override" "SKIPPED" "No alternate IP available (AD IP not found)"
    }


    # === DHCP Lease Lookup ===
    if ($dhcpModuleAvailable -and $dhcpServers.Count -gt 0) {
        Write-Section "DHCP Lease Lookup"

        $leaseFound = $false
        foreach ($dhcp in $dhcpServers) {
            $dhcpName = $dhcp.DnsName
            try {
                $scopes = @(Get-DhcpServerv4Scope -ComputerName $dhcpName -ErrorAction Stop)
                foreach ($scope in $scopes) {
                    $leases = @(Get-DhcpServerv4Lease -ComputerName $dhcpName -ScopeId $scope.ScopeId -ErrorAction Stop |
                        Where-Object { $_.HostName -match "^$computer(\.|$)" })

                    foreach ($lease in $leases) {
                        Write-TestResult "DHCP Lease" "OK" "$($lease.IPAddress) from $dhcpName (scope $($scope.ScopeId))"
                        $leaseFound = $true
                    }
                }
            }
            catch {
                $dhcpErr = $_.Exception.Message
                if ($dhcpErr.Length -gt 80) { $dhcpErr = $dhcpErr.Substring(0, 80) + "..." }
                Write-TestResult "DHCP ($dhcpName)" "FAILED" $dhcpErr
            }

            if ($leaseFound) { break }
        }

        if (-not $leaseFound) {
            Write-TestResult "DHCP Lease" "FAILED" "No lease found on any DHCP server"
        }
    }


    # === Summary ===
    Write-Section "Summary"

    if ($winrmByNameOk) {
        Write-Host "  WinRM is functional. No remediation needed." -ForegroundColor Green
        Write-Host ""
        continue
    }

    # -- What's broken --

    if ($ipMismatch) {
        Write-Host "  STALE DNS: DNS=$dnsIp does not match AD=$adIp" -ForegroundColor Red
        Write-Host "  DNS server is returning the wrong IP for this machine." -ForegroundColor Yellow
    }

    if ($winrmByNameErr -match '0x80090322') {
        Write-Host "  KERBEROS ERROR (0x80090322): Authentication failed despite machine being reachable." -ForegroundColor Red
        if ($secureChannelOk -eq $false) {
            Write-Host "  ROOT CAUSE: Broken secure channel -- machine account password is out of sync with AD." -ForegroundColor Yellow
        }
        elseif ($secureChannelOk -eq $true) {
            Write-Host "  Secure channel is healthy. Issue may be clock skew, stale tickets, or SPN." -ForegroundColor Yellow
        }
    }
    elseif ($winrmByNameErr -match 'cannot complete the operation') {
        Write-Host "  CONNECTIVITY ERROR: WinRM cannot reach the machine by hostname." -ForegroundColor Red
        Write-Host "  Likely stale DNS or firewall blocking WS-Man (5985/5986)." -ForegroundColor Yellow
    }
    elseif ($winrmByNameErr) {
        Write-Host "  WinRM FAILED: $winrmByNameErr" -ForegroundColor Red
    }

    if ($winrmServiceUp -eq $false) {
        Write-Host "  WinRM SERVICE IS STOPPED on the remote machine." -ForegroundColor Red
    }

    # -- Repair feasibility --

    if ($psExecToIpOk -and $psExecHostMatch) {
        Write-Host ""
        Write-Host "  REPAIRABLE: PsExec backdoor to AD IP ($adIp) is confirmed working." -ForegroundColor Green
        Write-Host "  Repair-RemoteAccess.ps1 can reach this machine and attempt automated fixes:" -ForegroundColor Green

        $fixes = @()
        if ($secureChannelOk -eq $false) {
            $fixes += "    - Reset secure channel (Test-ComputerSecureChannel -Repair)"
        }
        if ($clockSyncInfo -eq 'FAILED' -or $clockSyncInfo -eq 'Unknown') {
            $fixes += "    - Resync clock (w32tm /resync)"
        }
        if ($winrmServiceUp -eq $false) {
            $fixes += "    - Start WinRM service (sc start WinRM)"
        }
        # Always include these as general maintenance
        $fixes += "    - Purge stale Kerberos tickets (klist -li 0x3e7 purge)"
        $fixes += "    - Attempt DNS re-registration (ipconfig /registerdns)"

        foreach ($fix in $fixes) {
            Write-Host $fix -ForegroundColor DarkGray
        }
    }
    elseif ($psExecToIpOk -and -not $psExecHostMatch) {
        Write-Host ""
        Write-Host "  WARNING: PsExec reached a DIFFERENT machine at $adIp." -ForegroundColor Red
        Write-Host "  AD IPv4Address is stale. Cannot safely remediate." -ForegroundColor Yellow
    }
    elseif ($wmiExecToIpOk -and -not $psExecToIpOk) {
        Write-Host ""
        Write-Host "  PARTIALLY REPAIRABLE: WMI/DCOM backdoor works but PsExec does not." -ForegroundColor Yellow
        Write-Host "  Can create remote processes via Invoke-WmiMethod but output capture is limited." -ForegroundColor DarkGray
    }
    elseif (-not $winrmByNameOk -and $adIp -and -not $psExecToIpOk -and -not $wmiExecToIpOk) {
        Write-Host ""
        Write-Host "  NOT REPAIRABLE REMOTELY: No backdoor method reached the machine." -ForegroundColor Red
        Write-Host "  Machine may be offline, firewalled, or the AD IP ($adIp) is wrong." -ForegroundColor Yellow
        Write-Host "  Requires hands-on-keyboard or reimaging." -ForegroundColor DarkGray
    }

    Write-Host ""
}


Write-Host ""
Write-Host "========================================" -ForegroundColor White
Write-Host " Diagnostic Complete"                     -ForegroundColor White
Write-Host "========================================" -ForegroundColor White
Write-Host ""
