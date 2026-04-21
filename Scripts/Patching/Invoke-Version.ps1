# DOTS formatting comment

function Invoke-Version { # Start Function
    [CmdletBinding()]
    param(
        # Enter the software name. Make sure the software exists in the Main Switch.
        [Parameter(Mandatory, Position = 0)]
        [Alias("SoftwareName","Target", "SN", "TS")]
        [String]
        $TargetSoftware,

        # Enter a ComputerName to target only one machine.
        [Parameter(Position = 1)]
        [Alias("ComputerName", "CN", "TM")]
        [String]
        $TargetMachine,

        # Overrides the Main-Switch listPath with a custom .txt file of target machines
        [Parameter(Position = 2)]
        [Alias("ListFile", "TL")]
        [String]
        $TargetList,

        # Emit the sorted result objects on the pipeline in addition to the
        # terminal table. Used by the GUI to capture results from a background
        # runspace. Default off so interactive terminal use is unchanged.
        [Switch]
        $PassThru

    )


begin {

#region --- Extract data from Main Switch ---

# Clears variables to prevent conflicts after Ctrl + C
$varriableArray = @(
    "Tag"
    "Software"
    "ListPath"
    "CompliantVer"
    "PatchPath"
    "PatchName"
    "ProcessName"
    "PatchScript"
    "SoftwarePaths"
    "InstallLine"
    "KB"
    "TargetMachine"
    "SoftwareName"
)

Clear-Variable $varriableArray -Scope Global -ErrorAction 0

# Determines switch type
$switch = "Main-Switch.ps1"

# Pulls data from the switch
Set-ExecutionPolicy Bypass -Scope Process -Force *> $null
$switchPath = "$scriptPath\$switch"
$mainSwitch = (Get-Command $switchPath).ScriptBlock

$switchArguments = & $mainSwitch

# Variables from Main Switch are imported
$switchArguments.GetEnumerator() | ForEach-Object {
    # Pulls key only if it has a value in hash table
    if ($null -ne $($_.Value)) {
        # Creates variable from key name
        [string]$key = $($_.Key)
        $value = $($_.Value)

        New-Variable -Name $key -Value $value -Scope Script -Force
    }
}

#endregion --- Extract data from Main Switch ---


#region --- Setup before version check ---

# Override the Main-Switch listPath with a custom target list if provided
if ($TargetList) {
    if (Test-Path -LiteralPath $TargetList) {
        $listPath = $TargetList
    }
    else {
        Write-Warning "TargetList path '$TargetList' does not exist."
        break
    }
}

# Defines the verbose variable if entered as a parameter
$verbose = $PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent
if ($verbose) {
    Write-Host ""
    Write-Verbose "Verbose has been selected"
    Write-Host ""
}


# Catches mistakes in TargetSoftware param
if ($null -eq $softwarePaths) {
    $string = "$switch didn't contain anything matching " + '"' + $targetSoftware + '"' +
    ". You probably misspelled the software you were trying to query."
    Write-Warning $string
    break
}

# Length must be used to determine if Target Machine has data
if ($targetMachine.Length -eq 0) {

    # Checks that the list path is valid
    if (!(Test-Path $listPath)) {
        Write-Warning "'$listPath' is not a valid path."
        if (!(Test-Path (Split-Path $listPath))) {
            Write-Host "Creating new directory '$(Split-Path $listPath)'... Put your list in here!"
            Write-Host "The list needs to be called '$(Split-Path $listPath -Leaf)'"
            mkdir (Split-Path $listPath) > $null
        }
        break
    }

    # Checks that your list is populated with entries
    if ((Get-Content $listPath).Count -lt 1) {
        Write-Warning "The list at '$listPath' is empty."
        break
    }

    # Uses Format-ComputerList custom module to clean up list
    $list = Get-Content $listPath
    $listUniq = Format-ComputerList $list -ToUpper
}
else {
    $listUniq = Format-ComputerList $targetMachine -ToUpper
}


# Pulls date, admin name, and software (or KB) name
$date = Get-Date -Format "yyyy/MM/dd HH:mm"
$user = $env:USERNAME


# Populates $softwareName
if ($null -eq $softwareName) {
    $softwareName = $KB
}

if ($null -eq $softwareName) {
    $softwareName = $software
}


# Changes the name of the window -- only when hosted by a real console.
# The minimal runspace host used by Invoke-PatchGUI refuses user-
# interaction calls and throws on WindowTitle assignment, so guard this
# the same way Invoke-Patch does.
if ($Host.Name -eq "ConsoleHost") {
    $Host.Ui.RawUI.WindowTitle = "$software"
}

Write-Host "Beginning version fetch sequence on $software..."


#endregion --- Setup before version check ---


} # End begin



process {



    #region --- Build config for pipeline ---

    # Set version type (defaults to "File" unless explicitly "Product")
    if ($versionType -ne "Product") {
        $versionType = "File"
    }

    $config = @{
        Tag           = $tag
        Software      = $software
        SoftwareName  = $softwareName
        CompliantVer  = $compliantVer
        SoftwarePaths = $softwarePaths
        VersionType   = $versionType
        RegistryKey   = $registryKey
    }

    # Derive DNS suffix from the active network profile for IP-to-hostname
    # resolution. Empty string on a workgroup / unmatched host.
    $activeNetwork = if (Get-Command Get-RSLActiveNetwork -ErrorAction SilentlyContinue) {
        Get-RSLActiveNetwork
    } else { $null }
    $dnsSuffix = if ($activeNetwork) { "." + $activeNetwork.DomainFqdn } else { "" }

    #endregion --- Build config for pipeline ---



    #region --- Define version-check-only pipeline scriptblock ---

    $versionPipelineScriptBlock = {

        # All data enters via $args (no $Using: in runspaces)
        $computer       = $args[0]
        $config         = $args[1]
        $dnsSuffix      = $args[2]
        $date           = $args[3]
        $adminName      = $args[4]

        # Inline helper -- strips WinRM boilerplate from exception messages.
        # Must be defined inside the scriptblock; runspaces cannot see the
        # caller's imported modules (InitialSessionState::CreateDefault).
        function _CompressError ([string]$Msg) {
            $Msg = $Msg -replace 'Processing data from remote server \S+ failed with the following error message:\s*', ''
            $Msg = $Msg -replace 'Connecting to remote server \S+ failed with the following error message\s*:\s*', ''
            $Msg = $Msg -replace '\s*For more information, see the about_Remote_Troubleshooting Help topic\.', ''
            $Msg = $Msg -replace '^\[.+?\]\s*', ''
            $Msg = $Msg -replace '\r?\n', ' '
            $Msg = $Msg -replace '\s{2,}', ' '
            return $Msg.Trim()
        }
        $partialResults = $args[5]


        # Initialize result object (same schema as Invoke-Patch results)
        $result = [PSCustomObject]@{
            IPAddress    = $null
            ComputerName = $null
            Status       = $null
            SoftwareName = $config.SoftwareName
            Version      = $null
            Compliant    = $null
            NewVersion   = $null
            ExitCode     = $null
            Comment      = $null
            AdminName    = $adminName
            Date         = $date
        }

        # Determine if input is IP or hostname
        if ($computer -match '\.') {
            $result.IPAddress = $computer
        }
        else {
            $result.ComputerName = $computer
        }


        #--- PHASE 1: Reachability (ICMP with WinRM-port fallback for isolation) ---
        $PhaseTracker[$computer] = "Pinging"

        $pingResult = Test-Connection -ComputerName $computer -Count 1 -ErrorAction SilentlyContinue
        $ipAddr     = $null

        if ($null -ne $pingResult) {
            # ICMP succeeded -- machine is Online
            $result.Status = "Online"
            if ($null -ne $pingResult.IPV4Address) {
                $ipAddr = $pingResult.IPV4Address.IPAddressToString
            }
            elseif ($null -ne $pingResult.ProtocolAddress) {
                $ipAddr = $pingResult.ProtocolAddress
            }
        }
        else {
            # ICMP failed -- probe WinRM port 5985 via TCP SYN. Tanium-quarantined
            # machines block ICMP (via IPsec policy) but allowlist 5985 for admin
            # traffic, so a successful SYN here means the machine is reachable
            # and a version check can proceed via WinRM. Short 3s timeout keeps
            # the cost to truly-offline machines bounded.
            $PhaseTracker[$computer] = "Probing WinRM"
            $reachable = $false
            $tcp       = New-Object System.Net.Sockets.TcpClient
            try {
                $connectTask = $tcp.ConnectAsync($computer, 5985)
                $reachable   = $connectTask.Wait(3000)
            }
            catch {
                $reachable = $false
            }
            finally {
                $tcp.Close()
            }

            if ($reachable) {
                $result.Status = "Isolated"
            }
            else {
                $result.Status = "Offline"
                return $result
            }
        }


        #--- PHASE 2: DNS Resolution ---
        $PhaseTracker[$computer] = "DNS Lookup"

        if ($computer -match '\.') {
            # Input was an IP - resolve to hostname
            $result.IPAddress = $computer
            try {
                $dnsName = [System.Net.Dns]::GetHostByAddress($computer)
                if ($null -ne $dnsName) {
                    $result.ComputerName = $dnsName.HostName.Replace($dnsSuffix, "")
                }
            }
            catch {
                $result.Comment = "DNS Request Failed"
            }
        }
        else {
            # Input was a hostname - record IP from ping
            $result.ComputerName = $computer
            if ($null -ne $ipAddr) {
                $result.IPAddress = $ipAddr
            }
        }

        # Determine target name for WinRM operations
        $targetName = if ($null -ne $result.ComputerName) { $result.ComputerName } else { $result.IPAddress }
        if ($null -eq $targetName) {
            $result.Comment = "DNS Request Failed"
            return $result
        }


        # Save partial data so it survives if the runspace is stopped later
        $partialResults[$computer] = @{
            IPAddress    = $result.IPAddress
            ComputerName = $result.ComputerName
            Status       = $result.Status
        }


        #--- PHASE 3: Version Check ---
        $PhaseTracker[$computer] = "Version Check"

        $tag          = $config.Tag
        $compliantVer = $config.CompliantVer

        if ($tag -contains "RegVersion") {

            # Query registry for version info
            $registryKeys = $config.RegistryKey
            if ($null -eq $registryKeys) {
                $registryKeys = @(
                    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
                    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
                )
            }

            try {
                $regResult = Invoke-Command -ComputerName $targetName -ScriptBlock {
                    param($SwName, $RegKeys)
                    $versions = @()
                    foreach ($regKey in $RegKeys) {
                        $children = Get-ChildItem $regKey -ErrorAction SilentlyContinue -Force
                        if ($null -eq $children) { continue }
                        $props = Get-ItemProperty $children.PSPath -ErrorAction SilentlyContinue
                        foreach ($prop in $props) {
                            if ($prop.DisplayName -match $SwName) {
                                if ($null -ne $prop.DisplayVersion) {
                                    $versions += $prop.DisplayVersion
                                }
                            }
                        }
                    }
                    [PSCustomObject]@{ Version = $versions }
                } -ArgumentList $config.Software, $registryKeys -ErrorAction Stop

                if ($null -eq $regResult.Version -or $regResult.Version.Count -eq 0) {
                    $result.Version = "Not Installed"
                    $result.Compliant = $true
                }
                else {
                    [array]$result.Version = $regResult.Version
                }
            }
            catch {
                $result.Comment = "Version Check Failed: $(_CompressError "$_")"
                return $result
            }
        }
        else {
            # Query file version via Invoke-Command on remote machine
            $pathsStr    = [string]$config.SoftwarePaths
            $versionType = $config.VersionType

            try {
                $verResult = Invoke-Command -ComputerName $targetName -ScriptBlock {
                    param($PathsString, $VerType)

                    $versions = @()
                    $targetUsers = @()

                    # Reconstruct path array from space-separated string (split on "C:")
                    $paths = @()
                    foreach ($chunk in ($PathsString -split "C:")) {
                        if ($chunk -eq "") { continue }
                        $paths += "C:" + $chunk.Trim()
                    }

                    foreach ($path in $paths) {

                        # Handle USER paths (e.g. C:\Users\USER\AppData\...)
                        if ($path -cmatch 'USER') {
                            $userArray = (Get-ChildItem "C:\Users" -Force -Directory -ErrorAction SilentlyContinue).Name
                            $excludeUsers = @('Public', 'ADMINI~1')
                            $userArray = $userArray | Where-Object {
                                ($_ -notin $excludeUsers) -and ($_ -notmatch 'svc\d*\$')
                            }

                            foreach ($usr in $userArray) {
                                $userPath = $path.Replace('USER', $usr)

                                if (Test-Path $userPath) {
                                    $item = Get-Item $userPath -Force -ErrorAction SilentlyContinue
                                    if ($null -eq $item) { continue }

                                    if (($item.Mode -match 'a') -or ($item.Mode -eq '------')) {
                                        $fileItem = Get-ChildItem $userPath -Force -ErrorAction SilentlyContinue
                                        if ($VerType -eq 'Product') {
                                            $ver = $fileItem.VersionInfo.ProductVersion
                                        }
                                        else {
                                            $ver = $fileItem.VersionInfo.FileVersionRaw
                                        }
                                        if ($null -eq $ver) { $ver = $fileItem.VersionInfo.ProductVersion }
                                        if ($null -eq $ver) { $ver = $fileItem.VersionInfo.FileVersion }
                                        if ($null -eq $ver) { continue }
                                        if ($ver.GetType().Name -match 'string') {
                                            $ver = [version]($ver.Replace(',','.'))
                                        }
                                        $versions += $ver
                                        $targetUsers += $usr
                                    }
                                    elseif ($item.Mode -match 'd') {
                                        $targetUsers += $usr
                                    }
                                }
                            }
                        }
                        # Handle standard paths
                        elseif (Test-Path $path) {
                            $item = Get-Item $path -Force -ErrorAction SilentlyContinue
                            if ($null -eq $item) { continue }

                            if (($item.Mode -match 'a') -or ($item.Mode -eq '------')) {
                                $fileItem = Get-ChildItem $path -Force -ErrorAction SilentlyContinue
                                if ($VerType -eq 'Product') {
                                    $ver = $fileItem.VersionInfo.ProductVersion
                                }
                                else {
                                    $ver = $fileItem.VersionInfo.FileVersionRaw
                                }
                                if ($null -eq $ver) { $ver = $fileItem.VersionInfo.ProductVersion }
                                if ($null -eq $ver) { $ver = $fileItem.VersionInfo.FileVersion }
                                if ($null -eq $ver) { continue }
                                if ($ver.GetType().Name -match 'string') {
                                    $ver = [version]($ver.Replace(',','.'))
                                }
                                $versions += $ver
                            }
                            elseif ($item.Mode -match 'd') {
                                # Directory exists (folder-based detection)
                            }
                        }
                    }

                    [PSCustomObject]@{
                        Version     = $versions
                        TargetUsers = $targetUsers
                    }
                } -ArgumentList $pathsStr, $versionType -ErrorAction Stop

                if ($null -eq $verResult.Version -or $verResult.Version.Count -eq 0) {
                    $result.Version = "Not Installed"
                    $result.Compliant = $true
                }
                else {
                    [array]$result.Version = $verResult.Version
                }
            }
            catch {
                $result.Comment = "Version Check Failed: $(_CompressError "$_")"
                return $result
            }
        }


        # Compliance check
        if ($result.Version -ne "Not Installed" -and $null -ne $result.Version) {
            $result.Compliant = $true
            foreach ($ver in @($result.Version)) {
                # Version-string contains a failure marker -- skip version
                # math (which would throw on non-numeric input) and surface
                # the bad value as a Comment. Must match Invoke-Patch's
                # filter so the two tools agree on what counts as "error".
                if ("$ver" -match "Job|Failed|Error") {
                    $result.Comment = "Version $ver"
                    continue
                }
                try {
                    if ([Version]"$ver" -lt [Version]$compliantVer) {
                        $result.Compliant = $false
                    }
                }
                catch {}
            }
        }

        return $result

    } # End version pipeline scriptblock

    #endregion --- Define version-check-only pipeline scriptblock ---



    #region --- Build argument sets and execute ---

    # Thread-safe dictionary for partial results from stopped/timed-out runspaces.
    # Each runspace writes its progress here after completing ping/DNS,
    # so the data survives even if the runspace is killed mid-version-check.
    $partialResults = [System.Collections.Concurrent.ConcurrentDictionary[string, hashtable]]::new()

    # Build one argument array per machine
    $argumentSets = @(
        foreach ($machine in $listUniq) {
            , @(
                $machine,           # $args[0] = Computer
                $config,            # $args[1] = Config hashtable
                $dnsSuffix,         # $args[2] = DNS Suffix
                $date,              # $args[3] = Date
                $user,              # $args[4] = AdminName
                $partialResults     # $args[5] = Shared partial results dictionary
            )
        }
    )


    # Build parameters for Invoke-RunspacePool
    $runspaceParams = @{
        ScriptBlock    = $versionPipelineScriptBlock
        ArgumentList   = $argumentSets
        ThrottleLimit  = 32
        TimeoutMinutes = 5
        ActivityName   = "$software Version Check"
    }

    # Execute the version-check pipeline
    Write-Host "Checking online computers for software versions..."
    $pipelineResults = Invoke-RunspacePool @runspaceParams


    # Collect results
    $results = @($pipelineResults | Where-Object {
        $_ -is [PSCustomObject] -and $null -ne $_.Status
    })

    # Capture timed-out or failed machines
    $incompleteResults = @($pipelineResults | Where-Object {
        $_ -is [PSCustomObject] -and $null -eq $_.Status -and $null -ne $_.ComputerName
    })

    foreach ($incomplete in $incompleteResults) {
        # Check shared dictionary for partial data saved before the runspace was stopped
        $partial = $null
        $partialResults.TryGetValue($incomplete.ComputerName, [ref]$partial) > $null

        $results += [PSCustomObject]@{
            IPAddress    = if ($partial) { $partial.IPAddress }  else { $null }
            ComputerName = $incomplete.ComputerName
            Status       = if ($null -ne $incomplete.Status) { $incomplete.Status } else { "Online" }
            SoftwareName = $softwareName
            Version      = $null
            Compliant    = $null
            NewVersion   = $null
            ExitCode     = $null
            Comment      = $incomplete.Comment
            AdminName    = $user
            Date         = $date
        }
    }

    #endregion --- Build argument sets and execute ---



    #region --- Results formatting and output ---

    $displayResults = Add-Delimiter $results -Property @("Version", "NewVersion", "ExitCode")


    # Determines if the results output to terminal needs to be resized
    $resizeArray = @("Version", "NewVersion")

    $resizeRequired = $false
    foreach ($displayResult in $displayResults) {
        foreach ($prop in $resizeArray) {
            if (($displayResult.$prop -match '\d') -and ($displayResult.$prop -match " ")) {
                $resizeRequired = $true
                break
            }
        }
    }


    # If required, resizes table output to terminal
    if ($resizeRequired) {

        # Calculates the max length of certain outputs based on the current window size
        if ($null -ne (Get-Host).UI.RawUI.WindowSize) {
            $windowSize = (Get-Host).UI.RawUI.WindowSize.Width
            $maxLength = [math]::Round(($windowSize - 100) / 2)

            if ($maxLength -ge 100) {
                $maxLength = 100
            }
            elseif ($maxLength -le 15) {
                $maxLength = 15
            }
        }
        else {
            $maxLength = 50
        }

        # Formats the terminal output table
        foreach ($displayResult in $displayResults) {
            foreach ($prop in $resizeArray) {

                $stringArray = $displayResult.$prop.Split(",")

                # Shortens strings if they are too long
                $n = 0
                $outputString = $null
                foreach ($string in $stringArray) {
                    if ($string.Length + $outputString.Length -le $maxLength) {
                        if ($n -ne 0) {
                            $outputString = $outputString + "," + $string
                        }
                        else {
                            $outputString = $string
                            $n++
                        }
                    }
                    else {
                        $outputString = $outputString + "..."
                        break
                    }
                }

                $displayResult.$prop = $outputString
            }
        }
    }


    # For terminal output
    $displayProperties = @("IPAddress", "Computername", "Status", "SoftwareName",
        "Version", "Compliant", "NewVersion", "ExitCode", "Comment")

    # Output to Terminal only (no CSV - this is a situational awareness tool)
    $sortedResults = $displayResults | Select-Object $displayProperties | Sort-Object -Property (
        @{Expression = "Status"; Descending = $true},
        @{Expression = "Version"; Descending = $false}
    )
    $sortedResults | Format-Table -AutoSize | Out-Host

    # Emit structured results on the pipeline when requested (GUI capture)
    if ($PassThru) {
        $sortedResults
    }


    #endregion --- Results formatting and output ---


} # End process

} # End Function
