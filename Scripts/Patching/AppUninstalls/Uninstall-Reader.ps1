# DOTS formatting comment


# --- Local exit code lookup (this script runs remotely where modules are unavailable) ---
function Get-ExitCodeComment {
    param([int]$Code)
    switch ($Code) {
        0       { "Completed Successfully" }
        1       { "Incorrect Function" }
        2       { "File cannot be found" }
        3       { "Cannot find the specified path" }
        5       { "Access denied" }
        8       { "Not enough memory resources are available to process this command" }
        13      { "The data is invalid" }
        23      { "The component store has been corrupted / Generic Error - Check logs" }
        38      { "Windows is unable to load the device driver because a previous version is still in memory, resulting in conflicts" }
        53      { "The network path was not found" }
        59      { "Unexpected network error" }
        87      { "The parameter is incorrect" }
        112     { "There is not enough space on the disk." }
        184     { "A necessary file is locked by another process" }
        233     { "No process is on the other end of the pipe" }
        267     { "Directory name is invalid" }
        1060    { "The specified service does not exist as an installed service" }
        1392    { "A file or files are corrupt" }
        1450    { "Insufficient system resources exist to complete the requested service" }
        1602    { "User canceled installation" }
        1603    { "Fatal error during installation" }
        1605    { "This action is only valid for products that are currently installed" }
        1612    { "The installation source for this product is not available" }
        1618    { "Another installation is in progress" }
        1619    { "The installation package could not be opened" }
        1620    { "This installation package could not be opened. Contact the application vendor to verify that this is a valid Windows Installer package." }
        1635    { "This update package could not be opened" }
        1636    { "Update package cannot be opened" }
        1641    { "Successful - Restart required" }
        1642    { "Upgrade cannot be installed - Missing software" }
        1648    { "No valid sequence could be found for the set of patches" }
        1726    { "The remote procedure call failed" }
        3010    { "Successful - Restart required" }
        14001   { "The application failed to start" }
        14098   { "The component store has been corrupted" }
        2359302 { "The patch has already been installed" }
        -1073741510 { "Cmd.exe window was closed" }
        -2067919934 { "SQL server related error" }
        -2145116147 { "The update handler did not install the update because it needs to be downloaded again" }
        -2146959355 { ".NET framework installation failure" }
        -2145124329 { "Patch installation failure" }
        -2145124322 { "Generic error" }
        -2145124330 { "Another install is ongoing or reboot is pending" }
        -2146498167 { "The device is missing important security and quality fixes" }
        -2146498172 { "The matching component directory exists but binary is missing" }
        -2146498299 { "DISM Package Manager processed the command line but failed" }
        -2146498304 { "An unknown error occurred" }
        -2146498511 { "Corruption in the windows component store" }
        -2147417839 { "OLE received a packet with an invalid header (RPC Error)" }
        -2147467259 { "A file that the Windows Product Activation (WPA) requires is damaged or missing" }
        -2147956498 { "The component store has been corrupted" }
        default { $null }
    }
}


# --- Event log capture for action correlation ---
function Get-ActionEvents {
    param(
        [datetime]$Start,
        [datetime]$End
    )
    # Add small buffer for events logged slightly after process exits
    $End = $End.AddSeconds(5)
    $providers = @(
        'MsiInstaller'
        'Microsoft-Windows-WindowsUpdateClient'
        'Microsoft-Windows-Servicing'
    )
    $allEvents = @()
    foreach ($provider in $providers) {
        $filter = @{
            ProviderName = $provider
            StartTime    = $Start
            EndTime      = $End
        }
        try {
            $found = @(Get-WinEvent -FilterHashtable $filter -ErrorAction Stop)
            if ($found.Count -gt 0) {
                $allEvents += $found
            }
        }
        catch {
            # No events found for this provider -- continue
        }
    }
    $sorted = @($allEvents | Sort-Object TimeCreated | Select-Object -First 25)
    $result = @($sorted | ForEach-Object {
        [PSCustomObject]@{
            Time    = $_.TimeCreated.ToString('HH:mm:ss')
            Source  = $_.ProviderName
            Id      = $_.Id
            Message = (($_.Message -split "`n")[0]).Trim()
        }
    })
    return $result
}


# --- Safely parse a registry UninstallString into executable + args ---
# Registry UninstallString values are untrusted (any installed app can write
# arbitrary values). Handing the whole string to 'cmd /c' lets shell
# metacharacters (&, |, ;, etc.) chain arbitrary commands. This helper splits
# the string into a literal FilePath and a raw ArgumentList for use with
# Start-Process, which calls CreateProcess directly and bypasses the shell.
#
# Returns $null if no valid absolute-path executable can be resolved.
function Split-UninstallString {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$CommandLine
    )

    $trimmed = $CommandLine.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        return $null
    }

    # Case 1: quoted executable. Extract up to the closing quote.
    if ($trimmed.StartsWith('"')) {
        $closeIdx = $trimmed.IndexOf('"', 1)
        if ($closeIdx -gt 1) {
            $filePath = $trimmed.Substring(1, $closeIdx - 1)
            $argList  = $trimmed.Substring($closeIdx + 1).Trim()
            if (Test-Path -LiteralPath $filePath -PathType Leaf) {
                return [PSCustomObject]@{
                    FilePath     = $filePath
                    ArgumentList = $argList
                }
            }
        }
        return $null
    }

    # Case 2: unquoted. Walk whitespace-delimited tokens, progressively
    # joining until Test-Path succeeds. Require absolute path (drive letter
    # or UNC) to reject bare command names like 'cmd.exe' that could resolve
    # via CWD or PATH in a remote runspace.
    $tokens = @($trimmed -split '\s+')
    for ($i = 0; $i -lt $tokens.Count; $i++) {
        $candidate = ($tokens[0..$i] -join ' ')
        if ($candidate -notmatch '^(?:[A-Za-z]:\\|\\\\)') {
            continue
        }
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            $argList = ''
            if ($i -lt ($tokens.Count - 1)) {
                $argList = ($tokens[($i + 1)..($tokens.Count - 1)] -join ' ')
            }
            return [PSCustomObject]@{
                FilePath     = $candidate
                ArgumentList = $argList
            }
        }
    }

    return $null
}


# Unpack config hashtable passed from Invoke-Patch.ps1
$config        = $Args[0]
$softwareName  = $config.Software
$softwarePaths = @($config.SoftwarePaths)
$version       = $config.CompliantVer
$processName   = $config.ProcessName
$installLine   = @($config.InstallLine)
$versionType   = $config.VersionType



# --- Incremental log setup (write-through so timed-out machines still have logs) ---

$script:logPath = $null
try {
    $logDir = 'C:\Temp\PatchRemediation\Logs'
    $null = New-Item -Path $logDir -ItemType Directory -Force -ErrorAction Stop

    $logAdmin = if ($config.AdminName) { $config.AdminName } else { $env:USERNAME }
    $logTimestamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
    $safeName = ($softwareName -replace '[^a-zA-Z0-9._-]', '_')
    $logFileName = '{0}_{1}_{2}.log' -f $logTimestamp, $safeName, $logAdmin
    $script:logPath = Join-Path $logDir $logFileName

    $divider = '=' * 60
    $thinDiv = '-' * 60
    $header = @(
        $divider
        'Patch Remediation Log'
        $divider
        "Software    : $softwareName"
        "Target Ver  : $version"
        "Admin       : $logAdmin"
        "Date        : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "Computer    : $env:COMPUTERNAME"
        $thinDiv
        ''
    )
    $header | Out-File -FilePath $script:logPath -Encoding ASCII -Force
}
catch {
    # Log init failure must not break patch operation
}

function Write-ActionToLog {
    param($Action)
    if (-not $script:logPath) { return }
    try {
        $codeComment = ''
        if ($null -ne $Action.ExitCode) {
            $c = Get-ExitCodeComment $Action.ExitCode
            if ($c) { $codeComment = " - $c" }
        }
        $lines = @(
            "[$($Action.Phase)] $($Action.Command)"
            "  Exit Code : $($Action.ExitCode)$codeComment"
            "  Duration  : $($Action.Duration)"
            "  Source    : $($Action.Source)"
        )
        if ($Action.Error) {
            $lines += "  Error     : $($Action.Error)"
        }
        if ($Action.StartTime -and $Action.EndTime) {
            $events = @(Get-ActionEvents -Start $Action.StartTime -End $Action.EndTime)
            if ($events.Count -gt 0) {
                $evtLabel = "  Events ($($events.Count)):"
                if ($events.Count -ge 25) {
                    $evtLabel = "  Events (25+ showing first 25):"
                }
                $lines += $evtLabel
                foreach ($evt in $events) {
                    $evtMsg = $evt.Message
                    if ($evtMsg.Length -gt 80) {
                        $evtMsg = $evtMsg.Substring(0, 77) + '...'
                    }
                    $lines += "    $($evt.Time) | $($evt.Source) ($($evt.Id)) : $evtMsg"
                }
            }
            else {
                $lines += "  Events    : None captured"
            }
        }
        $lines += ''
        $lines | Out-File -FilePath $script:logPath -Encoding ASCII -Append
    }
    catch {
        # Log append failure must not break patch operation
    }
}

function Write-ProcessSnapshot {
    param(
        [string]$Label,
        [string[]]$Extra
    )
    if (-not $script:logPath) { return }
    try {
        $watchList = @($processName -split " ") + @('msiexec')
        if ($Extra) { $watchList += $Extra }
        $watchList = $watchList | Select-Object -Unique
        $found = @()
        foreach ($name in $watchList) {
            $procs = @(Get-Process $name -ErrorAction 0)
            foreach ($p in $procs) {
                $cmdLine = '(unavailable)'
                $cim = Get-CimInstance Win32_Process -Filter "ProcessId=$($p.Id)" -ErrorAction 0
                if ($cim.CommandLine) { $cmdLine = $cim.CommandLine }
                $found += "    $($p.Name) (PID $($p.Id)) | $cmdLine"
            }
        }
        $lines = @("[Snapshot] $Label")
        if ($found.Count -gt 0) {
            $lines += "  Processes ($($found.Count)):"
            $lines += $found
        }
        else {
            $lines += "  Processes : None found"
        }
        $lines += ''
        $lines | Out-File -FilePath $script:logPath -Encoding ASCII -Append
    }
    catch { }
}

# --- End incremental log setup ---


$actionLog = [System.Collections.ArrayList]::new()
$localVerArray = @()
$softwarePathsFull = @()
$softwarePathsComment = @()

foreach ($path in $softwarePaths) {

    # For user file paths like C:\Users\First.Last\AppData\...
    if ($path -cmatch 'USER') {

        # Builds array of users based on profiles in the C:\Users directory
        [System.Collections.ArrayList]$userArray = (Get-ChildItem "C:\Users").Name
        $userArray.Remove('Public')
        $userArray.Remove('ADMINI~1')
        $userArray.Remove('svcpatch01$')

        foreach ($user in $userArray) {
            $userPathRaw = $path.Replace('USER',"$user")
            $userPathGCI = Get-ChildItem $userPathRaw -Force -ErrorAction 0

            if ((Test-Path $userPathRaw) -and ($null -ne $userPathGCI)) {

                $item = Get-Item $userPathRaw -Force -ErrorAction 0

                # For files
                if (($item).Mode -match "a") {

                    if ($versionType -eq "Product") {
                        $userPathVer1 = $userPathGCI.VersionInfo.ProductVersion
                    }
                    else {
                        $userPathVer1 = $userPathGCI.VersionInfo.FileVersionRaw
                    }

                    if ($null -eq $userPathVer1) {
                        $userPathVer1 = $userPathGCI.VersionInfo.ProductVersion
                    }

                    if ($null -eq $userPathVer1) {
                        $userPathVer1 = $userPathGCI.VersionInfo.FileVersion
                    }

                    if ($null -eq $userPathVer1) {
                        continue
                    }

                    if (($userPathVer1).gettype().name -match "string") {
                        [version]$userPathVer1 = $userPathVer1.Replace(",",".")
                    }

                    $softwarePathsFull += $userPathRaw
                    $softwarePathsComment += $userPathGCI.FullName
                    $localVerArray += $userPathVer1

                    # Need to use try/catch in case multiple files are found
                    # due to wildcard use in file path. Will cause -lt operator
                    # to error out otherwise
                    try {
                        if ([version]$userPathVer1 -lt [version]$version) {
                            [array]$targetUsers += $user
                        }
                    }
                    catch {
                        [array]$targetUsers += $user
                    }
                }

                # For folders
                if (($item).Mode -match "d") {
                    [array]$targetUsers += $user
                }
            }
        }
    }

    # For system file paths like C:\Program Files\...
    else {
        $softwarePathsFull += $path

        $pathGCI = Get-ChildItem $path -Force -ErrorAction 0

        if ($versionType -eq "Product") {
            $pathVer1 = $pathGCI.VersionInfo.ProductVersion
        }
        else {
            $pathVer1 = $pathGCI.VersionInfo.FileVersionRaw
        }

        if ($null -eq $pathVer1) {
            $pathVer1 = $pathGCI.VersionInfo.ProductVersion
        }

        if ($null -eq $pathVer1) {
            $pathVer1 = $pathGCI.VersionInfo.FileVersion
        }

        if ($null -eq $pathVer1) {
            continue
        }

        if (($pathVer1).gettype().name -match "string") {
            [version]$pathVer1 = $pathVer1.Replace(",",".")
        }

        $localVerArray += $pathVer1
    }
}



$a = 0
[array]$installLineFull = $installLine
foreach ($line in $installLineFull) {
    if ($line -cmatch 'USER') {
        $userInstallLines = @()
        foreach ($targetUser in $targetUsers) {
            $userInstallLines += $line.Replace('USER',"$targetUser")
        }
        $installLineFull[$a] = $userInstallLines
    }
    $a++
}

$installLineArrayRaw = @()
foreach ($parentLine in $installLineFull) {
    foreach ($line0 in $parentLine) {
        $installLineArrayRaw += $line0
    }
}

$installLineArray = $installLineArrayRaw





# --- End Process Start ---

Set-ExecutionPolicy Bypass -Scope Process -Force

Write-ProcessSnapshot "Before Process Kill"

foreach ($process in $processName -split " ") {
    Get-Process $process -ErrorAction 0 | Stop-Process -Force
}

# Reader-specific processes only -- do NOT kill Acrobat.exe or FrameMaker.exe
$readerProcs = @('AcroRd32', 'AcroCEF', 'AcroBroker', 'AdobeCollabSync', 'AdobeARM')
foreach ($proc in $readerProcs) {
    Get-Process $proc -ErrorAction 0 | Stop-Process -Force
}

Write-ProcessSnapshot "After Process Kill"

# --- End Process End ---



# --- Uninstall Section Start ---


$regKeys = @(
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall'
    'HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall'
)

$uninstallKeys = Get-ItemProperty (Get-ChildItem $regKeys -ErrorAction SilentlyContinue -Force).PSPath |
    Where-Object {$_ -match "$softwareName"}

if ($null -eq $uninstallKeys -and $localVerArray.length -gt 0) {
    $regComment = "No Registry Keys"
}
else {
    foreach ($string in $uninstallKeys) {
        if ($string -match "msiexec.exe") {

            $string1 = $string.UninstallString -replace "msiexec.exe","" -replace "/I","" -replace "/x",""
            $string1 = $string1.trim()

            $msiArgs = "/X $string1 /quiet /norestart"
            $actionStart = Get-Date
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $proc = Start-Process "C:\Windows\System32\msiexec.exe" -ArgumentList $msiArgs -PassThru

            if ($proc.WaitForExit(600000)) {

                $code = $proc.ExitCode

            } else {

                $proc | Stop-Process -Force -ErrorAction SilentlyContinue

                $code = -1

            }
            $sw.Stop()
            $actionEnd = Get-Date

            $actionEntry = [PSCustomObject]@{
                Phase     = 'Uninstall'
                Command   = "msiexec.exe $msiArgs"
                ExitCode  = $code
                Duration  = $sw.Elapsed.ToString()
                Source    = 'Registry-MSI'
                StartTime = $actionStart
                EndTime   = $actionEnd
            }
            $actionLog.Add($actionEntry) > $null
            Write-ActionToLog $actionEntry

        }
        elseif ($string.QuietUninstallString) {

            $actionStart = Get-Date
            $parts = Split-UninstallString -CommandLine $string.QuietUninstallString

            if ($null -eq $parts) {
                $actionEnd = Get-Date
                $actionEntry = [PSCustomObject]@{
                    Phase     = 'Uninstall'
                    Command   = $string.QuietUninstallString
                    ExitCode  = $null
                    Duration  = '00:00:00'
                    Source    = 'Registry-QuietUninstall'
                    Error     = 'Unparseable UninstallString - no valid executable found'
                    StartTime = $actionStart
                    EndTime   = $actionEnd
                }
                $actionLog.Add($actionEntry) > $null
                Write-ActionToLog $actionEntry
                continue
            }

            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $spParams = @{
                FilePath = $parts.FilePath
                PassThru = $true
            }
            if (-not [string]::IsNullOrWhiteSpace($parts.ArgumentList)) {
                $spParams.ArgumentList = $parts.ArgumentList
            }
            $proc = Start-Process @spParams

            if ($proc.WaitForExit(600000)) {

                $uninstallExit = $proc.ExitCode

            } else {

                $proc | Stop-Process -Force -ErrorAction SilentlyContinue

                $uninstallExit = -1

            }
            $sw.Stop()
            $actionEnd = Get-Date

            $actionEntry = [PSCustomObject]@{
                Phase     = 'Uninstall'
                Command   = $string.QuietUninstallString
                ExitCode  = $uninstallExit
                Duration  = $sw.Elapsed.ToString()
                Source    = 'Registry-QuietUninstall'
                StartTime = $actionStart
                EndTime   = $actionEnd
            }
            $actionLog.Add($actionEntry) > $null
            Write-ActionToLog $actionEntry

        }
        elseif ($string.UninstallString) {

            $actionStart = Get-Date
            $parts = Split-UninstallString -CommandLine $string.UninstallString

            if ($null -eq $parts) {
                $actionEnd = Get-Date
                $actionEntry = [PSCustomObject]@{
                    Phase     = 'Uninstall'
                    Command   = $string.UninstallString
                    ExitCode  = $null
                    Duration  = '00:00:00'
                    Source    = 'Registry-Uninstall'
                    Error     = 'Unparseable UninstallString - no valid executable found'
                    StartTime = $actionStart
                    EndTime   = $actionEnd
                }
                $actionLog.Add($actionEntry) > $null
                Write-ActionToLog $actionEntry
                continue
            }

            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $spParams = @{
                FilePath = $parts.FilePath
                PassThru = $true
            }
            if (-not [string]::IsNullOrWhiteSpace($parts.ArgumentList)) {
                $spParams.ArgumentList = $parts.ArgumentList
            }
            $proc = Start-Process @spParams

            if ($proc.WaitForExit(600000)) {

                $uninstallExit = $proc.ExitCode

            } else {

                $proc | Stop-Process -Force -ErrorAction SilentlyContinue

                $uninstallExit = -1

            }
            $sw.Stop()
            $actionEnd = Get-Date

            $actionEntry = [PSCustomObject]@{
                Phase     = 'Uninstall'
                Command   = $string.UninstallString
                ExitCode  = $uninstallExit
                Duration  = $sw.Elapsed.ToString()
                Source    = 'Registry-Uninstall'
                StartTime = $actionStart
                EndTime   = $actionEnd
            }
            $actionLog.Add($actionEntry) > $null
            Write-ActionToLog $actionEntry

        }
    }
}

# --- Uninstall Section End ---



# =============================================================================
# Reader-Specific Cleanup
# =============================================================================
# Surgical removal of Adobe Acrobat Reader ONLY.
# Does NOT touch Adobe Acrobat (Pro/Standard), FrameMaker, or other Adobe products.
# =============================================================================


# --- Stop Adobe update service ---
# AdobeARMservice (Adobe Acrobat Update Service) locks Reader files and can
# block uninstall. We stop it but do NOT delete it -- if Acrobat Pro is also
# installed, it still needs this service.

$armService = Get-Service 'AdobeARMservice' -ErrorAction 0
if ($null -ne $armService -and $armService.Status -eq 'Running') {
    $actionStart = Get-Date
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Stop-Service 'AdobeARMservice' -Force -ErrorAction SilentlyContinue
    $sw.Stop()
    $actionEnd = Get-Date

    $actionEntry = [PSCustomObject]@{
        Phase     = 'Cleanup'
        Command   = 'Stop-Service AdobeARMservice'
        ExitCode  = 0
        Duration  = $sw.Elapsed.ToString()
        Source    = 'Reader-ServiceStop'
        StartTime = $actionStart
        EndTime   = $actionEnd
    }
    $actionLog.Add($actionEntry) > $null
    Write-ActionToLog $actionEntry
}


# --- Kill Reader processes again ---
# Stopping the service can take a moment; processes may have respawned.

foreach ($process in $processName -split " ") {
    Get-Process $process -ErrorAction 0 | Stop-Process -Force
}
foreach ($proc in $readerProcs) {
    Get-Process $proc -ErrorAction 0 | Stop-Process -Force
}


# --- Registry cleanup ---
# Remove Reader-specific registry keys only. Scoped to "Acrobat Reader" to
# avoid touching Acrobat Pro/Standard keys under the same Adobe parent.

$actionStart = Get-Date
$sw = [System.Diagnostics.Stopwatch]::StartNew()

New-PSDrive -PSProvider registry -Root HKEY_CLASSES_ROOT -Name HKCR -ErrorAction SilentlyContinue > $null

$readerRegKeys = @(
    # Standard uninstall entries (searched by $softwareName filter)
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    # Adobe vendor keys -- search children, filter to Reader only
    'HKLM:\Software\Adobe'
    'HKLM:\Software\WOW6432Node\Adobe'
    'HKCU:\SOFTWARE\Adobe'
    # HKCR installer products
    'HKCR:\Installer\Products'
)

$regDeleteCount = 0

foreach ($regKey in $readerRegKeys) {
    if (-not (Test-Path $regKey -ErrorAction 0)) { continue }

    $children = Get-ChildItem $regKey -ErrorAction 0 -Force
    if ($null -eq $children) { continue }

    $matchingKeys = @(Get-ItemProperty $children.PSPath -ErrorAction 0 |
        Where-Object { $_ -match $softwareName })

    foreach ($key in $matchingKeys) {
        if ($key.PSPath -match 'RegisteredApplications') { continue }
        Remove-Item $key.PSPath -Recurse -Force -ErrorAction SilentlyContinue
        $regDeleteCount++
    }
}

# Clean up Reader-specific vendor subkeys directly
# Use "Acrobat Reader" pattern to avoid matching "Acrobat DC" (Pro)
$readerVendorPaths = @(
    'HKLM:\SOFTWARE\Adobe\Acrobat Reader'
    'HKLM:\SOFTWARE\WOW6432Node\Adobe\Acrobat Reader'
    'HKCU:\SOFTWARE\Adobe\Acrobat Reader'
)

foreach ($vPath in $readerVendorPaths) {
    if (Test-Path $vPath -ErrorAction 0) {
        Remove-Item $vPath -Recurse -Force -ErrorAction SilentlyContinue
        $regDeleteCount++
    }
}

Remove-PSDrive -Name HKCR -ErrorAction SilentlyContinue

$sw.Stop()
$actionEnd = Get-Date

$actionEntry = [PSCustomObject]@{
    Phase     = 'Cleanup'
    Command   = "Registry cleanup ($regDeleteCount keys removed)"
    ExitCode  = 0
    Duration  = $sw.Elapsed.ToString()
    Source    = 'Reader-Registry'
    StartTime = $actionStart
    EndTime   = $actionEnd
}
$actionLog.Add($actionEntry) > $null
Write-ActionToLog $actionEntry


# --- Dynamic Reader folder detection ---
# Find all "Acrobat Reader*" folders under Adobe in both Program Files locations.
# This catches DC, 2017, and any future naming without hardcoding versions.

$readerFolders = @()

foreach ($pfRoot in @('C:\Program Files\Adobe', 'C:\Program Files (x86)\Adobe')) {
    if (Test-Path $pfRoot -ErrorAction 0) {
        $found = @(Get-ChildItem $pfRoot -Directory -ErrorAction 0 |
            Where-Object { $_.Name -match 'Acrobat Reader' })
        foreach ($f in $found) {
            $readerFolders += $f.FullName
        }
    }
}


# --- ACL fixes and file deletion ---
# Adobe's installer can set restrictive ACLs. Take ownership and grant
# full control before attempting deletion.

$actionStart = Get-Date
$sw = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($folder in $readerFolders) {
    if (-not (Test-Path $folder -ErrorAction 0)) { continue }

    # ACL Method 1 -- Take ownership
    $acl = Get-Acl -Path $folder -ErrorAction 0
    if ($null -ne $acl) {
        try {
            $owner = New-Object System.Security.Principal.Ntaccount("$env:USERNAME")
            $acl.SetOwner($owner)
            $acl | Set-Acl -Path $folder -ErrorAction SilentlyContinue
        }
        catch { }
    }

    # ACL Method 2 -- Grant FullControl
    try {
        $newAcl = Get-Acl -Path $folder -ErrorAction 0
        if ($null -ne $newAcl) {
            $identity = "$env:USERNAME"
            $rights = "FullControl"
            $type = "Allow"
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                $identity, $rights, $type)
            $newAcl.SetAccessRule($rule)
            Set-Acl -Path $folder -AclObject $newAcl -ErrorAction SilentlyContinue
        }
    }
    catch { }
}

$sw.Stop()
$actionEnd = Get-Date

if ($readerFolders.Count -gt 0) {
    $actionEntry = [PSCustomObject]@{
        Phase     = 'Cleanup'
        Command   = "ACL ownership and FullControl on Reader folders"
        ExitCode  = 0
        Duration  = $sw.Elapsed.ToString()
        Source    = 'Reader-ACL'
        StartTime = $actionStart
        EndTime   = $actionEnd
    }
    $actionLog.Add($actionEntry) > $null
    Write-ActionToLog $actionEntry
}


# --- Kill Reader processes one final time before deletion ---

foreach ($process in $processName -split " ") {
    Get-Process $process -ErrorAction 0 | Stop-Process -Force
}
foreach ($proc in $readerProcs) {
    Get-Process $proc -ErrorAction 0 | Stop-Process -Force
}


# --- Delete Reader files and folders ---

$actionStart = Get-Date
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$deleteCount = 0

foreach ($folder in $readerFolders) {
    if (-not (Test-Path $folder -ErrorAction 0)) { continue }

    # PowerShell Remove-Item first
    Remove-Item $folder -Recurse -Force -ErrorAction SilentlyContinue

    # Fallback to rmdir for stubborn files
    if (Test-Path $folder -ErrorAction 0) {
        cmd /c "rmdir /s /q `"$folder`"" 2>$null
    }

    # Final sweep for any remaining children
    if (Test-Path $folder -ErrorAction 0) {
        Get-ChildItem $folder -Recurse -Force -ErrorAction 0 |
            Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }

    $deleteCount++
}

$sw.Stop()
$actionEnd = Get-Date

if ($deleteCount -gt 0) {
    $actionEntry = [PSCustomObject]@{
        Phase     = 'Cleanup'
        Command   = "File/folder deletion ($deleteCount Reader folders targeted)"
        ExitCode  = 0
        Duration  = $sw.Elapsed.ToString()
        Source    = 'Reader-FileDelete'
        StartTime = $actionStart
        EndTime   = $actionEnd
    }
    $actionLog.Add($actionEntry) > $null
    Write-ActionToLog $actionEntry
}


# =============================================================================
# End Reader-Specific Cleanup
# =============================================================================



Write-ProcessSnapshot "After Cleanup" -Extra $readerProcs

# --- Install Section Start ---

try {
    foreach ($line in $installLineArray) {

        $actionStart = Get-Date
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $installer = [Scriptblock]::Create($line)
        & $installer
        $installExit = $LASTEXITCODE
        $sw.Stop()
        $actionEnd = Get-Date

        $actionEntry = [PSCustomObject]@{
            Phase     = 'Install'
            Command   = $line
            ExitCode  = $installExit
            Duration  = $sw.Elapsed.ToString()
            Source    = 'InstallLine'
            StartTime = $actionStart
            EndTime   = $actionEnd
        }
        $actionLog.Add($actionEntry) > $null
        Write-ActionToLog $actionEntry

    }

    # Gives the installer a little more time to generate exit code
    $lastAction = $actionLog | Select-Object -Last 1
    if ($null -ne $lastAction) {
        $expInstallTime = New-TimeSpan -Seconds 60
        $lastDuration = [TimeSpan]::Parse($lastAction.Duration)
        if (($null -eq $lastAction.ExitCode) -and ($lastDuration.TotalSeconds -lt 60)) {
            Start-Sleep (60 - [int]$lastDuration.TotalSeconds)
        }
    }
}
catch {
    $actionEnd = Get-Date
    $actionEntry = [PSCustomObject]@{
        Phase     = 'Install'
        Command   = $line
        ExitCode  = $null
        Duration  = $sw.Elapsed.ToString()
        Source    = 'InstallLine-Error'
        Error     = $_.Exception.Message
        StartTime = $actionStart
        EndTime   = $actionEnd
    }
    $actionLog.Add($actionEntry) > $null
    Write-ActionToLog $actionEntry
}

# --- Install Section End ---



$updateVerArray = @()
foreach ($softwarePath in $softwarePathsFull) {
    $updatePathGCI = Get-ChildItem $softwarePath -Force -ErrorAction 0

    if ($versionType -eq "Product") {
        $updatePathVer1 = $updatePathGCI.VersionInfo.ProductVersion
    }
    else {
        $updatePathVer1 = $updatePathGCI.VersionInfo.FileVersionRaw
    }

    if ($null -eq $updatePathVer1) {
        $updatePathVer1 = $updatePathGCI.VersionInfo.ProductVersion
    }

    if ($null -eq $updatePathVer1) {
        $updatePathVer1 = $updatePathGCI.VersionInfo.FileVersion
    }

    if ($null -eq $updatePathVer1) {
        continue
    }

    if (($updatePathVer1).gettype().name -match "string") {
        [version]$updatePathVer1 = $updatePathVer1.Replace(",",".")
    }

    $updateVerArray += $updatePathVer1
}


# Formats new version output
if ($updateVerArray.length -eq 0) {

    $updateVerReg = (Get-ItemProperty (Get-ChildItem $regKeys -ErrorAction SilentlyContinue).PSPath |
        Where-Object {$_ -match $softwareName}).DisplayVersion

    if (($null -ne $updateVerReg) -and ($null -eq $regComment)) {
        $regComment = "Registry Version: $updateVerReg"
    }

    if ($null -eq $updateVerArray[0]) {
        $updateVerArray = "Removed"
    }
}
elseif ($localVerArray.length -ge 1) {
    if ($updateVerArray[0] -eq $localVerArray[0]) {
        $updateVerArray = "No Change"
    }
}

if (($updateVerArray -match "Removed") -and ($localVerArray[0] -eq $null)) {
    $updateVerArray = $null
}



#region --- Service checks ---


$services = @("Netlogon","WinRM")

foreach ($service in $services) {
    if ((Get-Service $service).Status -ne "Running") {
        Start-Service $service -PassThru -ErrorAction SilentlyContinue
    }
}


#endregion --- Service checks ---



# Derive exit code array from action log
$exitCodeArray = @($actionLog | Where-Object { $null -ne $_.ExitCode } |
    ForEach-Object { $_.ExitCode })

# Enrich action log entries with exit code comments
foreach ($action in $actionLog) {
    if ($null -ne $action.ExitCode) {
        $action | Add-Member -NotePropertyName 'Comment' -NotePropertyValue (
            Get-ExitCodeComment $action.ExitCode) -Force
    }
}

# Get comment for the final exit code
$exitCodeComment = $null
if ($exitCodeArray.Count -gt 0) {
    $exitCodeComment = Get-ExitCodeComment $exitCodeArray[-1]
}


# Formats comments
$commentArray = @(
    $exitCodeComment
    $regComment
    $softwarePathsComment
)

$b = 0
foreach ($comment in $commentArray) {
    if ($comment.Length -gt 0) {
        if ($b -eq 0) {
            $finalComment = $comment
            $b++
        }
        else {
            $finalComment = $finalComment, $comment -join " - "
        }
    }
}


#region --- Finalize Patch Remediation Log ---

if ($script:logPath) {
    try {
        $thinDiv = '-' * 60
        $divider = '=' * 60
        $finalExitDisplay = $exitCodeArray[-1]
        $finalCodeComment = ''
        if ($exitCodeComment) { $finalCodeComment = " - $exitCodeComment" }

        @(
            $thinDiv
            "Result      : ExitCode $finalExitDisplay$finalCodeComment"
            $divider
        ) | Out-File -FilePath $script:logPath -Encoding ASCII -Append
    }
    catch {
        # Log finalize failure must not break patch operation
    }
}

#endregion --- Finalize Patch Remediation Log ---


# Generates variable with results
$results = [PSCustomObject]@{
    NewVersion = $updateVerArray
    ExitCode   = $exitCodeArray[-1]
    ExitCodes  = $exitCodeArray
    Comment    = $finalComment
    ActionLog  = @($actionLog)
}

return $results
