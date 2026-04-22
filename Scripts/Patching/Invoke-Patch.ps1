# DOTS formatting comment

<#
    .SYNOPSIS
        Deploys a patch across a fleet of Windows endpoints concurrently.
    .DESCRIPTION
        Primary orchestrator for the toolkit. Resolves a software definition from
        Main-Switch.ps1, builds the target list, and fans deployment out across
        the fleet via Invoke-RunspacePool. Per machine, runs the pipeline:

            ping -> DNS resolution -> version check -> copy patch files
                 -> execute deploy script -> post-install verify

        Timeout is derived dynamically from patch file size (small = 35 min,
        large = up to 120 min) and can be overridden. Results come back as a
        uniform table regardless of per-machine success, failure, or timeout,
        so downstream Format-Table / Export-Csv consumers never have to
        special-case missing rows.

        Written by Skyler Werner
    .EXAMPLE
        Invoke-Patch -TargetSoftware Edge
        Patches every machine listed in Desktop\Lists\Microsoft_Edge.txt.
    .EXAMPLE
        Invoke-Patch -TS Chrome -TM WORKSTATION01
        Patches a single machine, using the short parameter aliases.
    .EXAMPLE
        Invoke-Patch -TargetSoftware Edge -ConfirmTimeout -CollectLogs
        Deploys with interactive confirmation before killing timed-out tasks
        and copies per-machine install logs back to the operator's desktop.
#>

function Invoke-Patch {
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

        # Forces patch to run on all machines (Necessary if some file paths are folders with no version info)
        [Parameter(Position = 2)]
        [Switch]
        $Force,

        # Skips copying files to machines
        [Parameter(Position = 3)]
        [Alias("SkipCopy")]
        [Switch]
        $NoCopy,

        # Sets forced timeout for copying in minutes instead of default dynamic timeout based on file size
        [Parameter(Position = 4)]
        [ValidateRange(0, 120)]
        [Int]
        $CopyTimeout,

        # Sets forced timeout for patching in minutes
        [Parameter(Position = 5)]
        [ValidateRange(0, 240)]
        [Int]
        $Timeout,

        # Prompts for confirmation before stopping timed-out tasks (default is auto-stop)
        [Parameter(Position = 6)]
        [Alias("CT")]
        [Switch]
        $ConfirmTimeout,

        # Retrieves detailed per-machine patch logs from remote machines after patching
        [Parameter(Position = 7)]
        [Alias("GetLogs", "CopyLogs")]
        [Switch]
        $CollectLogs,

        # Overrides the Main-Switch listPath with a custom .txt file of target machines
        [Parameter(Position = 8)]
        [Alias("ListFile", "TL")]
        [String]
        $TargetList,

        # Emits the results array on the pipeline (in addition to the host table).
        # Used by Invoke-PatchGUI to capture results from a background runspace.
        [Parameter(Position = 9)]
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


        #region --- Setup before patching ---

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


        # Ensures the item path for Copy is valid (skip check when patchPath is null, e.g. uninstall-only)
        if ($NoCopy.IsPresent -eq $false -and $null -ne $patchPath) {
            if (-not (Test-Path "$patchPath\$patchName")) {
                Write-Warning "Patch path '$patchPath\$patchName' does not exist. Check $switch for errors."
                break
            }
        }

        # Catches mistakes in TargetSoftware param
        if ($null -eq $softwarePaths) {
            $string = "$switch didn't contain anything matching " + '"' + $targetSoftware + '"' +
            ". You probably misspelled the software you were trying to patch."
            Write-Warning $string
            break
        }

        # Checks that the Patch Script is valid
        if (!($patchScript)) {
            Write-Warning ('$PathScript' + " string '$patchScript' is not valid. Check $switch for errors.")
            break
        }

        # Checks for PSTools
        if ($tag -contains "PSExec") {
            if (!(Test-Path "$env:USERPROFILE\Desktop\PSTools\PsExec.exe")) {
                Write-Warning ("$env:USERPROFILE\Desktop\PSTools\PsExec.exe was not found! Microsoft " +
                    ".msu updates require this file.")
                break
            }
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
            $listFormatted = Format-ComputerList $list -ToUpper
        }
        else {
            $listFormatted = Format-ComputerList $targetMachine -ToUpper
        }


        # Pulls date, admin name, and software (or KB) name
        $dateOutput = Get-Date -Format "yyyy-MM-dd-HHmm"
        $date = Get-Date -Format "yyyy/MM/dd HH:mm"
        $user = $env:USERNAME


        # Populates $softwareName
        if ($null -eq $softwareName) {
            $softwareName = $KB
        }

        if ($null -eq $softwareName) {
            $softwareName = $software
        }


        # Changes the name of the window
        if ($Host.Name -eq "ConsoleHost") {
            $Host.Ui.RawUI.WindowTitle = "$software"
        }

        Write-Host "Beginning patching sequence for " -NoNewline
        Write-Host "$software" -ForegroundColor Cyan -NoNewline
        Write-Host "..."


        #endregion --- Setup before patching ---


    } # End begin



    process {



        #region --- Build config and arguments ---

        # Convert PatchScript ScriptBlock to string (scriptblocks cannot cross runspace boundaries)
        $patchScriptStr = $patchScript.ToString()

        # Build config hashtable with all Main-Switch values for the pipeline
        $config = @{
            Tag           = $tag
            Software      = $software
            SoftwareName  = $softwareName
            CompliantVer  = $compliantVer
            PatchPath     = $patchPath
            PatchName     = $patchName
            ProcessName   = $processName
            SoftwarePaths = $softwarePaths
            InstallLine   = $installLine
            KB            = $KB
            RegistryKey   = $registryKey
            AdminName     = $user
        }

        # Set version type (defaults to "File" unless explicitly "Product")
        if ($versionType -ne "Product") {
            $versionType = "File"
        }
        $config.VersionType = $versionType

        # Build argument list for Default.ps1 / Default-PSExec.ps1
        if ($tag -match "PsExec") {
            $scriptArgList = @(
                $patchPath
                $patchName
                [string]$installLine
            )
        }
        else {
            # Pass $config hashtable directly -- Invoke-Command serializes it natively.
            # Default.ps1 / Default-NoUninstall.ps1 unpack values from $Args[0].
            $scriptArgList = @($config)
        }

        if ($verbose) {
            Write-Host ""
            Write-Verbose "Argument List:"
            Write-Host ""
            $scriptArgList
            Write-Host ""
        }

        # Derive DNS suffix from the active network profile for IP-to-hostname
        # resolution. Empty string on a workgroup / unmatched host.
        $activeNetwork = if (Get-Command Get-RSLActiveNetwork -ErrorAction SilentlyContinue) {
            Get-RSLActiveNetwork
        } else { $null }
        $dnsSuffix = if ($activeNetwork) { "." + $activeNetwork.DomainFqdn } else { "" }

        #endregion --- Build config and arguments ---



        #region --- Calculate dynamic timeout ---

        # Install timeout from Main-Switch (per-region), defaults to 30 if not set
        if (-not $installTimeout) { $installTimeout = 30 }

        if ($timeout -ge 1) {
            # User provided explicit timeout
            $dynamicTimeout = [int]$timeout
        }
        elseif ($copyTimeout -ge 1) {
            # User provided copy timeout; add install time from Main-Switch
            $dynamicTimeout = [int]$copyTimeout + $installTimeout
        }
        else {
            # Calculate dynamic timeout based on patch size (preserves Copy-ItemAsJob logic)
            $patchSizeBytes = 0
            if ($null -ne $patchPath -and (Test-Path $patchPath)) {
                (Get-ChildItem $patchPath -Recurse -ErrorAction SilentlyContinue) | ForEach-Object {
                    $patchSizeBytes += $_.Length
                }
            }
            $divideBy100MB = $patchSizeBytes / 105000000

            # Copy time based on file size + install time from Main-Switch
            if     ($patchSizeBytes -eq 0)                                 { $copyTime = 0  }   # No files to copy
            elseif ($divideBy100MB -lt 0.1)                                { $copyTime = 5  }   # < 10 MB
            elseif (($divideBy100MB -ge 0.1) -and ($divideBy100MB -lt 1))  { $copyTime = 10 }   # < 100 MB
            elseif (($divideBy100MB -ge 1)   -and ($divideBy100MB -lt 2))  { $copyTime = 20 }   # < 200 MB
            elseif (($divideBy100MB -ge 2)   -and ($divideBy100MB -lt 5))  { $copyTime = 30 }   # < 500 MB
            elseif (($divideBy100MB -ge 5)   -and ($divideBy100MB -lt 10)) { $copyTime = 45 }   # < 1 GB
            elseif (($divideBy100MB -ge 10)  -and ($divideBy100MB -lt 30)) { $copyTime = 60 }   # < 3 GB
            elseif (($divideBy100MB -ge 30)  -and ($divideBy100MB -lt 50)) { $copyTime = 75 }   # < 5 GB
            elseif ($divideBy100MB -ge 50)                                 { $copyTime = 90 }   # >= 5 GB

            $dynamicTimeout = $copyTime + $installTimeout
        }

        if ($verbose) {
            Write-Verbose "Dynamic Timeout: $dynamicTimeout minutes (install budget: $installTimeout)"
        }

        #endregion --- Calculate dynamic timeout ---



        #region --- Pre-calculate origin file info for copy verification ---

        $originFileCount = 0
        $originFileSize = 0
        $originHashArray = @()

        if (-not $NoCopy.IsPresent -and $null -ne $patchPath -and (Test-Path $patchPath)) {
            $originFiles = Get-ChildItem $patchPath -Recurse -ErrorAction SilentlyContinue
            $originMeasure = $originFiles | Measure-Object -Sum Length
            $originFileCount = $originMeasure.Count
            $originFileSize = $originMeasure.Sum

            # Pre-calculate hashes (top-level files only, matching Copy-ItemAsJob behavior)
            $originTopFiles = Get-ChildItem $patchPath -ErrorAction SilentlyContinue | Where-Object { -not $_.PSIsContainer }
            if ($originTopFiles) {
                $originHashArray = @($originTopFiles | Get-FileHash | Select-Object @{N='FileName';E={Split-Path $_.Path -Leaf}}, Hash)
            }
        }

        $config.OriginFileCount = $originFileCount
        $config.OriginFileSize  = $originFileSize
        $config.OriginHashes    = $originHashArray

        #endregion --- Pre-calculate origin file info for copy verification ---



        #region --- Define per-machine pipeline scriptblock ---

        $pipelineScriptBlock = {

            # All data enters via $args (no $Using: in runspaces)
            $computer       = $args[0]
            $config         = $args[1]
            $force          = $args[2]
            $noCopy         = $args[3]
            $patchScriptStr = $args[4]
            $scriptArgList  = $args[5]
            $dnsSuffix      = $args[6]
            $date           = $args[7]
            $adminName      = $args[8]
            $partialResults = $args[9]

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

            $_scriptStart = [DateTime]::Now

            # Initialize result object (same schema as original results table)
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
                # and patching can proceed via WinRM. Short 3s timeout keeps the
                # cost to truly-offline machines bounded.
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
                        if (-not $force) { return $result }
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
                        if (-not $force) { return $result }
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
                    if ("$ver" -match "Failed|Error") {
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

                if ($result.Compliant -and -not $force) {
                    return $result
                }
            }


            # Update partial data with version/compliance info
            $partialResults[$computer] = @{
                IPAddress    = $result.IPAddress
                ComputerName = $result.ComputerName
                Status       = $result.Status
                Version      = $result.Version
                Compliant    = $result.Compliant
            }


            #--- PHASE 4: Copy via Robocopy ---
            $PhaseTracker[$computer] = "Copying Files"

            if ((-not $noCopy) -and ($null -ne $config.PatchPath)) {
                $patchPath  = $config.PatchPath
                $itemFolder = Split-Path $patchPath -Leaf
                $remoteDest = "\\$targetName\C`$\Temp"
                $destPath   = "$remoteDest\$itemFolder"

                $copyRequired = $true

                # Size verification (fast metadata check via UNC)
                if (Test-Path $destPath) {
                    try {
                        $destMeasure = Get-ChildItem $destPath -Recurse -ErrorAction SilentlyContinue | Measure-Object -Sum Length
                        if (($destMeasure.Count -eq $config.OriginFileCount) -and ($destMeasure.Sum -eq $config.OriginFileSize)) {

                            # Sizes match - verify hashes on remote machine (avoids reading files over network)
                            if ($config.OriginHashes.Count -gt 0) {
                                $hashMismatch = $false
                                try {
                                    $remoteHashes = Invoke-Command -ComputerName $targetName -ScriptBlock {
                                        param($LocalDest, $Folder)
                                        $destItemPath = "$LocalDest\$Folder"
                                        $files = Get-ChildItem $destItemPath -ErrorAction SilentlyContinue | Where-Object { -not $_.PSIsContainer }
                                        if ($null -eq $files) { return @() }
                                        $files | Get-FileHash | Select-Object @{N='FileName';E={Split-Path $_.Path -Leaf}}, Hash
                                    } -ArgumentList "C:\Temp", $itemFolder -ErrorAction Stop

                                    foreach ($originHash in $config.OriginHashes) {
                                        $matching = $remoteHashes | Where-Object { $_.FileName -eq $originHash.FileName }
                                        if ($null -eq $matching -or $matching.Hash -ne $originHash.Hash) {
                                            $hashMismatch = $true
                                            break
                                        }
                                    }
                                }
                                catch {
                                    $hashMismatch = $true
                                }

                                if (-not $hashMismatch) {
                                    $copyRequired = $false
                                }
                            }
                            else {
                                # No hashes to check, size match is sufficient
                                $copyRequired = $false
                            }
                        }
                    }
                    catch {}
                }

                if ($copyRequired) {
                    # Remove stale file if a file exists where a directory is expected
                    if (Test-Path $destPath -PathType Leaf) {
                        Remove-Item $destPath -Force > $null
                    }

                    # robocopy source=contents destination=full path (unlike Copy-Item which nests automatically)
                    $robocopyArgs = @(
                        "`"$patchPath`""           # source directory
                        "`"$destPath`""            # destination directory (includes folder name)
                        '/E'                       # copy subdirectories including empty ones
                        '/R:3'                     # retry 3 times on failed copies
                        '/W:5'                     # wait 5 seconds between retries
                        '/MT:4'                    # multi-threaded copy (4 threads)
                        '/NP'                      # no progress percentage (cleaner output)
                        '/NDL'                     # no directory listing in output
                        '/NFL'                     # no file listing in output (keep output concise)
                        '/NJH'                     # no job header
                        '/NJS'                     # no job summary
                    )
                    $robocopyOutput = & robocopy @robocopyArgs 2>&1
                    $robocopyExit   = $LASTEXITCODE

                    # Robocopy exit codes: 0-7 = success (bitmask), 8+ = failure
                    if ($robocopyExit -ge 8) {
                        $exitMeaning = switch ($robocopyExit) {
                            8  { "Some files could not be copied" }
                            16 { "Fatal error - no files were copied" }
                            default { "Unexpected error" }
                        }
                        $result.Comment = "Copy Failed (robocopy exit $robocopyExit): $exitMeaning"
                        return $result
                    }
                }
            }


            #--- PHASE 5: Install ---
            $PhaseTracker[$computer] = "Patching"

            $patchScriptBlock = [ScriptBlock]::Create($patchScriptStr)

            if ($tag -match "PsExec") {
                # PSExec: invoke scriptblock locally on admin workstation
                # Default-PSExec.ps1 expects: $patchPath, $patchName, $installLine, [verbose], $computerName (last)
                $psexecArgs = @($scriptArgList) + @($targetName)
                try {
                    $installResult = & $patchScriptBlock @psexecArgs
                }
                catch {
                    $result.Comment = "Patch Failed: $(_CompressError "$_")"
                    return $result
                }

                $result.ExitCode = $installResult.ExitCode
                if ($null -ne $installResult.Comment) {
                    $result.Comment = $installResult.Comment
                }
            }
            else {
                # Standard: Invoke-Command to remote machine with Default.ps1 / Default-NoUninstall.ps1
                try {
                    $installResult = Invoke-Command -ComputerName $targetName `
                        -ScriptBlock $patchScriptBlock `
                        -ArgumentList $scriptArgList `
                        -ErrorAction Stop `
                        -InformationAction Ignore

                    $result.ExitCode = $installResult.ExitCode
                    if ($null -ne $installResult.Comment) {
                        $result.Comment = $installResult.Comment
                    }
                }
                catch {
                    $result.Comment = "Patch Failed: $(_CompressError "$_")"
                    return $result
                }
            }


            #--- PHASE 6: Post-Install Version Check ---
            $PhaseTracker[$computer] = "Verifying"

            if ($tag -match "PsExec") {
                # PSExec: check version via UNC paths (local read, same as original Invoke-Patch lines 694-710)
                $softwarePaths = $config.SoftwarePaths
                $versionType   = $config.VersionType

                $remotePaths = @()
                foreach ($sp in @($softwarePaths)) {
                    $remotePath = "$sp".Replace(":", "`$")
                    $remotePaths += "\\$targetName\$remotePath"
                }

                $newVersions = @()
                foreach ($rp in $remotePaths) {
                    $items = Get-ChildItem $rp -Force -ErrorAction SilentlyContinue
                    foreach ($fileItem in $items) {
                        if ($versionType -eq "Product") {
                            $ver = $fileItem.VersionInfo.ProductVersion
                        }
                        else {
                            $ver = $fileItem.VersionInfo.FileVersionRaw
                        }
                        if ($null -eq $ver) { $ver = $fileItem.VersionInfo.ProductVersion }
                        if ($null -eq $ver) { $ver = $fileItem.VersionInfo.FileVersion }
                        if ($null -ne $ver) { $newVersions += $ver }
                    }
                }

                [array]$result.NewVersion = $newVersions

                if ([string]$result.NewVersion -eq [string]$result.Version) {
                    if ([string]$result.NewVersion -ne "") {
                        $result.NewVersion = "No Change"
                    }
                }
            }
            elseif ($tag -contains "RegVersion") {
                # RegVersion: re-query registry for new version
                $registryKeys = $config.RegistryKey
                if ($null -eq $registryKeys) {
                    $registryKeys = @(
                        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
                        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
                    )
                }

                try {
                    $newRegResult = Invoke-Command -ComputerName $targetName -ScriptBlock {
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

                    if ($null -eq $newRegResult.Version -or $newRegResult.Version.Count -eq 0) {
                        [array]$result.NewVersion = "Removed"
                    }
                    else {
                        [array]$result.NewVersion = $newRegResult.Version
                    }
                }
                catch {
                    # Keep existing comment, append version check failure
                    if ($null -ne $result.Comment) {
                        $result.Comment = $result.Comment + " | New Version Check Failed: $(_CompressError "$_")"
                    }
                    else {
                        $result.Comment = "New Version Check Failed: $(_CompressError "$_")"
                    }
                }
            }
            else {
                # Standard: NewVersion was returned by Default.ps1 via Invoke-Command
                if ($null -ne $installResult.NewVersion) {
                    [array]$result.NewVersion = $installResult.NewVersion
                }
            }

            # Update Avg Success in the progress display (exit code 0 or 3010 = success)
            if ($result.ExitCode -eq 0 -or $result.ExitCode -eq 3010) {
                $dur = ([DateTime]::Now - $_scriptStart).TotalSeconds
                $StatusMessage['_count'] = [int]$StatusMessage['_count'] + 1
                $StatusMessage['_sum']   = [double]$StatusMessage['_sum'] + $dur
                $avg = $StatusMessage['_sum'] / $StatusMessage['_count']
                $avgSpan = [TimeSpan]::FromSeconds($avg)
                if ($avgSpan.TotalMinutes -ge 1) {
                    $avgStr = "~{0}m {1:D2}s" -f [math]::Floor($avgSpan.TotalMinutes), $avgSpan.Seconds
                } else {
                    $avgStr = "~{0}s" -f [math]::Floor($avgSpan.TotalSeconds)
                }
                $StatusMessage['Text'] = "Avg Success: $avgStr"
            }

            return $result

        } # End pipeline scriptblock

        #endregion --- Define per-machine pipeline scriptblock ---



        #region --- Build argument sets and execute ---

        # Thread-safe dictionary for partial results from stopped/timed-out runspaces.
        # Each runspace writes its progress here after completing ping/DNS and version
        # check phases, so the data survives even if the runspace is killed mid-install.
        $partialResults = [System.Collections.Concurrent.ConcurrentDictionary[string, hashtable]]::new()

        # Build one argument array per machine
        $argumentSets = @(
            foreach ($machine in $listFormatted) {
                , @(
                    $machine,              # $args[0] = Computer
                    $config,               # $args[1] = Config hashtable
                    $force.IsPresent,      # $args[2] = Force
                    $NoCopy.IsPresent,     # $args[3] = NoCopy
                    $patchScriptStr,       # $args[4] = PatchScript as string
                    $scriptArgList,        # $args[5] = ArgumentList for remote script
                    $dnsSuffix,            # $args[6] = DNS Suffix
                    $date,                 # $args[7] = Date
                    $user,                 # $args[8] = AdminName
                    $partialResults        # $args[9] = Shared partial results dictionary
                )
            }
        )


        # Build parameters for Invoke-RunspacePool
        $runspaceParams = @{
            ScriptBlock    = $pipelineScriptBlock
            ArgumentList   = $argumentSets
            ThrottleLimit  = 50
            TimeoutMinutes = $dynamicTimeout
            ActivityName   = $software
        }

        if ($confirmTimeout.IsPresent) {
            $runspaceParams.ConfirmTimeout = $true
        }
        if ($verbose) {
            Write-Host ""
            Write-Verbose "Runspace Parameters:"
            Write-Host "  ThrottleLimit  = 50"
            Write-Host "  TimeoutMinutes = $dynamicTimeout"
            Write-Host "  Machines       = $($listFormatted.Count)"
            Write-Host ""
        }


        # Execute the per-machine pipeline
        $pipelineResults = Invoke-RunspacePool @runspaceParams


        # Filter valid results (complete PSCustomObjects from the pipeline scriptblock)
        $results = @($pipelineResults | Where-Object {
            $_ -is [PSCustomObject] -and $null -ne $_.PSObject.Properties['SoftwareName']
        })

        # Capture incomplete results (timed-out or failed objects from the runspace module)
        $incompleteResults = @($pipelineResults | Where-Object {
            $_ -is [PSCustomObject] -and $null -ne $_.ComputerName -and
            $null -eq $_.PSObject.Properties['SoftwareName']
        })

        foreach ($incomplete in $incompleteResults) {
            # Check shared dictionary for partial data saved before the runspace was stopped
            $partial = $null
            $partialResults.TryGetValue($incomplete.ComputerName, [ref]$partial) > $null

            $results += [PSCustomObject]@{
                IPAddress    = if ($partial) { $partial.IPAddress }    else { $null }
                ComputerName = $incomplete.ComputerName
                Status       = if ($null -ne $incomplete.Status) { $incomplete.Status } else { "Online" }
                SoftwareName = $softwareName
                Version      = if ($partial) { $partial.Version }     else { $null }
                Compliant    = if ($partial) { $partial.Compliant }   else { $null }
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
        $loggingResults = Add-Delimiter $results -Property @("Version", "NewVersion", "ExitCode") -Delimiter ";"


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

        # For logging
        $logProperties = @("IPAddress", "Computername", "Status", "Version", "Compliant",
            "NewVersion", "ExitCode", "AdminName", "Date")


        if (!(Test-Path "$env:USERPROFILE\Desktop\Patch-Results" -PathType Container)) {
            mkdir "$env:USERPROFILE\Desktop\Patch-Results" -Force > $null
        }

        # Log to the active network's central PatchLog share when reachable;
        # fall back to a local backlog folder on the operator's desktop so
        # no results are lost when the share is offline.
        $loggingRoot = $null
        if ($activeNetwork -and $activeNetwork.PatchLogShareUnc -and
            (Test-Path $activeNetwork.PatchLogShareUnc)) {
            $loggingRoot = $activeNetwork.PatchLogShareUnc
        }
        else {
            $loggingRoot = "$env:USERPROFILE\Desktop\Patch-Results\ShareBacklog"
            if (-not (Test-Path $loggingRoot -PathType Container)) {
                mkdir $loggingRoot -Force > $null
            }
            Write-Warning "Network share unavailable - logging to $loggingRoot"
        }

        $resultsRoot = "$env:USERPROFILE\Desktop\Patch-Results"
        $fileName = "$software" + "_" + "$dateOutput"

        # Output to Terminal
        # Sort cascade surfaces the most actionable rows first:
        #   1. Status      custom -- Isolated (rare + high-signal,
        #                  often Tanium quarantine or WinRM down)
        #                  first, Online bulk second, Offline tail.
        #   2. NewVersion  DESC -- "No Change" ('N') sorts above
        #                  numeric versions so failures bubble above
        #                  successes within each Status block.
        #   3. ExitCode    DESC -- within an outcome: 3010 > 0
        #                  (reboot required is more actionable than
        #                  clean success); 1618 > 1603 (transient
        #                  "install in progress" before real fatal).
        #   4. Version     ASC -- oldest-version grouping as the
        #                  final outcome tiebreaker.
        #   5. Comment     ASC -- stable ordering for matching text.
        #   6. ComputerName ASC -- deterministic within-group order.
        $displayResults | Select-Object $displayProperties | Sort-Object -Property (
            @{ Expression = {
                switch ($_.Status) {
                    'Isolated' { 1 }
                    'Online'   { 2 }
                    'Offline'  { 3 }
                    default    { 99 }
                }
            }; Descending = $false },
            @{ Expression = "NewVersion";   Descending = $true  },
            @{ Expression = "ExitCode";     Descending = $true  },
            @{ Expression = "Version";      Descending = $false },
            @{ Expression = "Comment";      Descending = $false },
            @{ Expression = "ComputerName"; Descending = $false }
        ) | Format-Table -AutoSize | Out-Host

        # Saves to Server Desktop
        $loggingResults | Select-Object $logProperties | Sort-Object Version |
            Export-Csv "$resultsRoot\$fileName.csv" -Append -Force -NoTypeInformation

        # Saves to Share
        $loggingResults | Select-Object $logProperties | Sort-Object Version |
            Export-Csv "$loggingRoot\$fileName.csv" -Append -Force -NoTypeInformation


        # Emit delimited results on the pipeline for programmatic consumers
        # (e.g. Invoke-PatchGUI's DataGrid). Use $displayResults so Version
        # and NewVersion are comma-separated strings instead of raw arrays
        # -- the DataGrid renders arrays as "System.Object[]" otherwise.
        if ($PassThru) {
            $displayResults
        }


        # --- Retrieve detailed per-machine logs from remote machines ---
        if ($CollectLogs) {
            $safeName = ($softwareName -replace '[^a-zA-Z0-9._-]', '_')
            $logFilter = "*_${safeName}_*.log"
            $onlineNames = @($results |
                Where-Object { $_.Status -eq "Online" } |
                ForEach-Object { $_.ComputerName })

            if ($onlineNames.Count -gt 0) {
                Copy-Log -ComputerName $onlineNames `
                    -Filter $logFilter `
                    -RemoteLogPath 'C:\Temp\PatchRemediation\Logs' `
                    -DestinationFolder 'Patch-Logs'
            }
            else {
                Write-Warning "No online machines to retrieve logs from."
            }
        }

        #endregion --- Results formatting and output ---


    } # End process



    <#
.SYNOPSIS
This function utilizes a library of modules to update or uninstall software for a list of computers.

.DESCRIPTION
This function pulls software data from the "Main Switch". This file is
where all information concerning patching or uninstalling each piece of software is
stored. Specifying "Reader" for the parameter TargetSoftware will import the following
information:

$tag           = "3rdParty"
$listPath      = "$listPathRoot\Adobe_Reader.txt"
$software      = "Adobe Acrobat Reader"
$processName   = "AcroRd32"
$compliantVer  = "20.13.20064" # "20.13.20064"
$patchPath     = "D:\VMT\Patches\AdobeReader_20.013.20064"
$patchScript   = (Get-Command "$ScriptRepoPath\Default.ps1").ScriptBlock
$softwarePaths = "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
$installLine   = '& cmd /c "C:\Temp\AdobeReader_20.013.20064\setup.exe"'

For more information on the Main Switch, read the Get-Help for Main-Switch.ps1

Each machine progresses independently through a per-machine pipeline using a RunspacePool:
  Ping -> DNS Resolution -> Version Check -> Copy -> Install -> Post-Version Check

This eliminates the batch-phase bottleneck where all machines had to wait for the
slowest machine at each step.

.EXAMPLE
Invoke-Patch Chrome

Runs the function against computers targeting Google Chrome. "Chrome" exists as a
switch in the Main Switch as a group of arguments for Google Chrome. The switch
is not case sensitive.

.EXAMPLE
Invoke-Patch -TargetSoftware FireFox -Force

Runs the function against computers targeting Mozilla Firefox. "FireFox" exists as a
switch in the Main Switch as a group of arguments for Mozilla FireFox.

It will run even if the Firefox version is already compliant or Firefox not installed.

.EXAMPLE
Invoke-Patch -TargetSoftware FlashKB -CopyTimeout 20

Runs the function against computers targeting Adobe Flash Player. "FlashKB" exists as
one of several switches in the Main Switch that contain arguments for Adobe Flash.
This switch utilizes a Microsoft Update KB file.

The timeout will be calculated as CopyTimeout + install budget from Main-Switch.
Otherwise, the timeout is dynamic based on file or folder size + install budget.

.INPUTS
This function does not support inputs from the pipeline.

.OUTPUTS
A summary of the results is output into the terminal. More detailed information is
saved to the server and the Systems Share for long term logging.

Terminal Output example:

ComputerName    Status Version     Compliant NewVersion ExitCode Comment
------------    ------ -------     --------- ---------- -------- -------
WKSTN-000000123 Online 32.0.0.387      False 32.0.0.445        0 Completed Successfully

Logging Example:

ComputerName    Status Version     Compliant NewVersion ExitCode AdminName           Date
------------    ------ -------     --------- ---------- -------- ---------           ----
WKSTN-000000123 Online 32.0.0.387      False 32.0.0.445        0 first.last.w.admin  2020/12/11 13:46

.NOTES
Written by Skyler Werner
Version 2.0 - Rebuilt with RunspacePool per-machine pipeline for concurrent execution.
Replaces batch-phase architecture (Test-ConnectionAsJob, Get-VersionAsJob, Copy-ItemAsJob,
Get-InvokeAsJob/Get-StartJob) with a single Invoke-RunspacePool call.

Version 2.1 - Replaced Copy-Item with robocopy for patch file transfers.
Provides multi-threaded copying (/MT:4), automatic retries (/R:3 /W:5), and
descriptive error reporting via robocopy exit codes in the Comment field.

.LINK
Modules:
https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_modules?view=powershell-7.1

.FUNCTIONALITY
Use this script to mitigate software vulnarabilities on many machines simultaneously.
#>

} # End Function
