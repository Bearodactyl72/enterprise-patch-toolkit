# DOTS formatting comment


# --- Local exit code lookup (runs in runspace where modules are unavailable) ---
function Get-ExitCodeComment {
    param([int]$Code)
    switch ($Code) {
        0       { "Success" }
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
        [datetime]$End,
        [string]$ComputerName
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
        $params = @{
            FilterHashtable = $filter
            ErrorAction     = 'Stop'
        }
        if ($ComputerName) {
            $params['ComputerName'] = $ComputerName
        }
        try {
            $found = @(Get-WinEvent @params)
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


$patchPath     = $Args[0]
$patchName     = $Args[1]
$installLine   = $Args[2]
$computerName  = $Args | Select-Object -Last 1



#region --- Patch argument building ----


# Pulls admin username, patch name, and date, for logging
$user = $env:USERNAME
$date = Get-Date -Format "yyyy-MM-dd_HHmm"
$log  = "/log:c:\Temp\$date.$patchName.$user.evt"


# Decides what installer to use based on file extension
switch ($patchName) {
    {$_ -match ".msu"} {
        $installer = "c:\windows\system32\wusa.exe"
    }
    {$_ -match ".msi"} {
        $installer = "msiexec /install"
    }
    {$_ -match ".msp"} {
        $installer = "msiexec /update"
    }
    {$_ -match "MAgent"} {
        $mcAfee = $true
    }
}


# Makes patch path as it will be after it's copied to the Temp folder on a remote computer
$itemFolder = $patchPath.Split("\") | Select-Object -Last 1
$tempPatchPath = "$itemFolder" + "\" + "$patchName"


# Defines argument list
if ($mcAfee) {
    $arguments = @("\\$($computerName)", "-accepteula", "-h", "-s", "C:\Temp\$($tempPatchPath) /Install=Agent /Silent /ForceInstall")
}
else {
    $arguments = @("\\$($computerName)", "-accepteula", "-s", "$installer C:\Temp\$($tempPatchPath) /quiet /norestart $log")
}


#endregion --- Patch argument building ----





#region --- PsExec install ---

# --- Incremental log setup (write-through so timed-out machines still have logs) ---

$script:logPath = $null
try {
    $logDir = "\\$computerName\C`$\Temp\PatchRemediation\Logs"
    $null = New-Item -Path $logDir -ItemType Directory -Force -ErrorAction Stop

    $safeName = ($patchName -replace '[^a-zA-Z0-9._-]', '_')
    $logFileName = '{0}_{1}_{2}.log' -f $date, $safeName, $user
    $script:logPath = Join-Path $logDir $logFileName

    $divider = '=' * 60
    $thinDiv = '-' * 60
    $header = @(
        $divider
        'Patch Remediation Log'
        $divider
        "Software    : $patchName"
        "Target Ver  : N/A"
        "Admin       : $user"
        "Date        : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "Computer    : $computerName"
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
            $events = @(Get-ActionEvents -Start $Action.StartTime -End $Action.EndTime -ComputerName $computerName)
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

# --- End incremental log setup ---


$actionLog = [System.Collections.ArrayList]::new()

# Write breadcrumb before PsExec starts so timed-out machines show what was attempted
$psExecCommand = "PsExec.exe $($arguments -join ' ')"
if ($script:logPath) {
    try {
        $breadcrumb = "[Install] STARTED at $(Get-Date -Format 'HH:mm:ss') -- $psExecCommand"
        @($breadcrumb, '  (awaiting completion)', '') |
            Out-File -FilePath $script:logPath -Encoding ASCII -Append
    }
    catch { }
}

try {
    $actionStart = Get-Date
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $result = (Start-Process "$env:USERPROFILE\Desktop\PSTools\PsExec.exe" -ArgumentList $arguments -WindowStyle Minimized -Wait -PassThru).ExitCode
    $sw.Stop()
    $actionEnd = Get-Date

    $actionEntry = [PSCustomObject]@{
        Phase     = 'Install'
        Command   = $psExecCommand
        ExitCode  = $result
        Duration  = $sw.Elapsed.ToString()
        Source    = 'PsExec'
        StartTime = $actionStart
        EndTime   = $actionEnd
    }
    $actionLog.Add($actionEntry) > $null
    Write-ActionToLog $actionEntry

    $expectedInstallTime = New-TimeSpan -Seconds 60
    if (($null -eq $result) -and ($sw.Elapsed -lt $expectedInstallTime)) {
        Start-Sleep 60
    }
}
catch {
    $sw.Stop()
    $actionEnd = Get-Date
    $actionEntry = [PSCustomObject]@{
        Phase     = 'Install'
        Command   = $psExecCommand
        ExitCode  = $null
        Duration  = $sw.Elapsed.ToString()
        Source    = 'PsExec-Error'
        Error     = $_.Exception.Message
        StartTime = $actionStart
        EndTime   = $actionEnd
    }
    $actionLog.Add($actionEntry) > $null
    Write-ActionToLog $actionEntry
}


#endregion --- PsExec install ---



#region --- Service checks ---


$services = @("Netlogon","WinRM")

foreach ($service in $services) {
    if ((Get-Service $service).Status -ne "Running") {
        Start-Service $service -PassThru -ErrorAction SilentlyContinue
    }
}


#endregion --- Service checks ---



#region --- Builds results table for return ----


# Enrich action log entries with exit code comments
foreach ($action in $actionLog) {
    if ($null -ne $action.ExitCode) {
        $action | Add-Member -NotePropertyName 'Comment' -NotePropertyValue (
            Get-ExitCodeComment $action.ExitCode) -Force
    }
}

# Get comment for the final exit code
$comment = $null
if ($null -ne $result) {
    $comment = Get-ExitCodeComment $result
}


#region --- Finalize Patch Remediation Log ---

if ($script:logPath) {
    try {
        $thinDiv = '-' * 60
        $divider = '=' * 60
        $finalCodeComment = ''
        if ($comment) { $finalCodeComment = " - $comment" }

        @(
            $thinDiv
            "Result      : ExitCode $result$finalCodeComment"
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
    ComputerName = $computerName
    ExitCode     = $result
    Comment      = $comment
    ActionLog    = @($actionLog)
}

return $results


#endregion --- Builds results table for return ----

