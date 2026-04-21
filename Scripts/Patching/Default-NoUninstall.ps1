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


# Unpack config hashtable passed from Invoke-Patch.ps1
$config        = $Args[0]
$softwareName  = $config.Software
$softwarePaths = @($config.SoftwarePaths)
$version       = $config.CompliantVer
$processName   = $config.ProcessName
$installLine   = @($config.InstallLine)
$versionType   = $config.VersionType

# Suppress DISM/cmdlet progress bars (e.g. Add/Remove-WindowsCapability)
$ProgressPreference = 'SilentlyContinue'


# --- Incremental log setup (write-through so timed-out machines still have logs) ---

$script:logPath = $null
try {
    $logDir = if ($config.LogDir) { $config.LogDir } else { 'C:\Temp\PatchRemediation\Logs' }
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

Write-ProcessSnapshot "After Process Kill"

# --- End Process End ---



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



$regKeys = @(
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall'
    'HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall'
)

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
