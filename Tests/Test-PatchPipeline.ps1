# DOTS formatting comment

<#
    .SYNOPSIS
        Test suite for the Invoke-Patch per-machine pipeline.
    .DESCRIPTION
        Validates that the Invoke-Patch pipeline scriptblock works correctly
        inside runspaces -- the exact environment where it runs in production.
        Tests cover all 6 phases of the pipeline plus error handling, result
        schema, compliance logic, and the Compress-ExceptionMessage regression.

        All tests are local-only; no remote machines or network shares needed.

        Tests 8-13 and 16 require WinRM loopback (Invoke-Command to localhost)
        from runspace threads. If unavailable, these tests are skipped gracefully.

        For full 64/64 coverage, this test script must be run from an elevated
        (Run as Administrator) PowerShell window AND the following one-time
        setup must be in place:

            Set-Item WSMan:\localhost\Client\TrustedHosts -Value 'localhost' -Force
            Enable-PSRemoting -Force

        Network profile must be Private (not Public). Driver updates or adapter
        resets can revert these settings -- re-run the commands above if the
        probe reports "UNAVAILABLE".

        Tests performed:

          Test  1 : Offline Machine Path (Phase 1 - Ping)
          Test  2 : _CompressError Works in Runspace
          Test  3 : WinRM Error Catch Block (Phase 3 - Version Check)
          Test  4 : REGRESSION -- Non-Existent Function in Catch Block
          Test  5 : Result Schema Completeness
          Test  6 : No Silent Drops (Online + Failed = Comment Populated)
          Test  7 : Compliance Logic
          Test  8 : Compliant Machine Exits Early (no patch attempted)
          Test  9 : -Force Flag Overrides Compliance
          Test 10 : File Copy Phase (Phase 4 - Robocopy to localhost)
          Test 11 : Patch Execution Phase (Phase 5 - Dummy patch script)
          Test 12 : Exit Code Handling (0, 3010, 1603, etc.)
          Test 13 : Post-Install Verification (Phase 6)
          Test 14 : Partial Result Recovery (timeout mid-pipeline)
          Test 15 : Isolated Mode (skip ping)
          Test 16 : Localhost Full Pipeline (all phases, if WinRM available)

        Estimated runtime: 2-3 minutes.

        Written by Skyler Werner
        Date: 2026/03/24
        Version 1.0.0

    .PARAMETER ModulePath
        Path to Invoke-RunspacePool.psm1. Auto-detected if not specified.
    .EXAMPLE
        .\Test-PatchPipeline.ps1
        .\Test-PatchPipeline.ps1 -ModulePath "C:\Modules\Invoke-RunspacePool.psm1"
#>

param(
    [string]$ModulePath
)

$ErrorActionPreference = 'Continue'


#region ===== SETUP =====

# --- Test result tracker ---

$script:TestResults = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:TestNumber  = 0

function Assert-Test {
    param(
        [string]$TestName,
        [string]$Description,
        [object]$Expected,
        [object]$Actual,
        [ValidateSet('Equals','NotNull','IsNull','Contains','GreaterThan','NotEquals')]
        [string]$CompareMode = 'Equals'
    )

    $script:TestNumber++

    $passed = switch ($CompareMode) {
        'Equals'      { "$Actual" -eq "$Expected" }
        'NotNull'     { $null -ne $Actual -and "$Actual" -ne "" }
        'IsNull'      { $null -eq $Actual }
        'Contains'    { "$Actual" -match [regex]::Escape("$Expected") }
        'GreaterThan' { [int]$Actual -gt [int]$Expected }
        'NotEquals'   { "$Actual" -ne "$Expected" }
    }

    $status = if ($passed) { 'PASS' } else { 'FAIL' }
    $color  = if ($passed) { 'Green' } else { 'Red' }
    $marker = if ($passed) { '' } else { '  <<<' }

    Write-Host "  [$status] " -ForegroundColor $color -NoNewline
    Write-Host "$Description" -NoNewline

    switch ($CompareMode) {
        'Equals'      { Write-Host "  (Expected: $Expected | Actual: $Actual)$marker" }
        'NotEquals'   { Write-Host "  (Expected NOT: $Expected | Actual: $Actual)$marker" }
        'Contains'    { Write-Host "  (Pattern: '$Expected' in '$Actual')$marker" }
        'GreaterThan' { Write-Host "  (Actual: $Actual > $Expected)$marker" }
        default       { Write-Host "  (Value: $Actual)$marker" }
    }

    $script:TestResults.Add([PSCustomObject]@{
        Number      = $script:TestNumber
        Test        = $TestName
        Description = $Description
        Status      = $status
        Expected    = $Expected
        Actual      = $Actual
    })
}


# --- Locate and import module ---

if (-not $ModulePath) {
    $candidates = @(
        (Join-Path (Split-Path $PSScriptRoot -Parent) "Modules\Invoke-RunspacePool\Invoke-RunspacePool.psm1")
        (Join-Path $PSScriptRoot "Modules\Invoke-RunspacePool\Invoke-RunspacePool.psm1")
        ".\Modules\Invoke-RunspacePool\Invoke-RunspacePool.psm1"
    )
    foreach ($c in $candidates) {
        if (Test-Path $c) { $ModulePath = (Resolve-Path $c).Path; break }
    }
}

if (-not $ModulePath -or -not (Test-Path $ModulePath)) {
    Write-Host "ERROR: Cannot find Invoke-RunspacePool.psm1" -ForegroundColor Red
    Write-Host "  Searched:"
    foreach ($c in $candidates) { Write-Host "    $c" }
    Write-Host "  Use -ModulePath to specify the location."
    return
}

Write-Host ""
Write-Host ("=" * 72) -ForegroundColor Cyan
Write-Host "  Invoke-Patch Pipeline  --  Test Suite" -ForegroundColor Cyan
Write-Host ("=" * 72) -ForegroundColor Cyan
Write-Host ""
Write-Host "  Module : $ModulePath"
Write-Host "  PS Ver : $($PSVersionTable.PSVersion)"
Write-Host "  Host   : $($Host.Name)"
Write-Host "  Date   : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host ""

if ($PSVersionTable.PSVersion.Major -ne 5) {
    Write-Host "  NOTE: Production targets PS 5.1. Tests should pass on any version," -ForegroundColor Yellow
    Write-Host "        but run on PS 5.1 for authoritative results." -ForegroundColor Yellow
    Write-Host ""
}

Import-Module $ModulePath -Force
Write-Host "  Module imported." -ForegroundColor Gray
Write-Host ""


# --- Check WinRM loopback availability (used by several tests) ---
$winrmAvailable = $false
try {
    $null = Test-WSMan -ComputerName localhost -ErrorAction Stop
    $winrmAvailable = $true
    Write-Host "  WinRM loopback: " -NoNewline
    Write-Host "Available" -ForegroundColor Green
}
catch {
    Write-Host "  WinRM loopback: " -NoNewline
    Write-Host "Unavailable (some tests will validate error paths instead)" -ForegroundColor Yellow
}
Write-Host ""

# --- Check WinRM loopback from INSIDE a runspace (the actual test path) ---
$winrmInRunspace = $false
if ($winrmAvailable) {
    Import-Module $ModulePath -Force
    $winrmTestBlock = {
        $computer = $args[0]
        try {
            $r = Invoke-Command -ComputerName localhost -ScriptBlock { $env:COMPUTERNAME } -ErrorAction Stop
            [PSCustomObject]@{ ComputerName = $computer; Success = $true; Result = $r; Error = $null }
        }
        catch {
            [PSCustomObject]@{ ComputerName = $computer; Success = $false; Result = $null; Error = "$_" }
        }
    }
    $winrmTestArgs = @( , @("WINRM-PROBE") )
    $probe = @(Invoke-RunspacePool -ScriptBlock $winrmTestBlock -ArgumentList $winrmTestArgs -ThrottleLimit 1 -TimeoutMinutes 1 -ActivityName "Probe")
    if ($probe.Count -gt 0 -and $probe[0].Success -eq $true) {
        $winrmInRunspace = $true
        Write-Host "  WinRM in runspace: " -NoNewline
        Write-Host "Available" -ForegroundColor Green
    }
    else {
        Write-Host "  WinRM in runspace: " -NoNewline
        Write-Host "UNAVAILABLE" -ForegroundColor Red
        $errMsg = if ($probe.Count -gt 0 -and $probe[0].Error) { $probe[0].Error }
                  elseif ($probe.Count -gt 0 -and $probe[0].Comment) { $probe[0].Comment }
                  else { "(no probe result returned)" }
        Write-Host "    Error: $errMsg" -ForegroundColor DarkRed
        Write-Host ""
        Write-Host "    Tests 8-13, 16 require Invoke-Command -ComputerName localhost" -ForegroundColor Yellow
        Write-Host "    from a runspace thread and will be skipped." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "    To fix:" -ForegroundColor Yellow
        Write-Host "      1. Run this test from an elevated (Run as Admin) PowerShell" -ForegroundColor White
        Write-Host "      2. Ensure the following one-time setup has been run (elevated):" -ForegroundColor White
        Write-Host "         Set-Item WSMan:\localhost\Client\TrustedHosts -Value 'localhost' -Force" -ForegroundColor White
        Write-Host "         Enable-PSRemoting -Force" -ForegroundColor White
        Write-Host ""
        Write-Host "    NOTE: Network profile must be Private (not Public)." -ForegroundColor Yellow
        Write-Host "    Driver updates or adapter resets can revert these settings." -ForegroundColor Yellow
    }
    Write-Host ""
}


$sw = [System.Diagnostics.Stopwatch]::StartNew()


# --- Common parameters ---
$common = @{ ThrottleLimit = 5; TimeoutMinutes = 2; ActivityName = "Test" }


# --- Expected result schema ---
$expectedProperties = @(
    "IPAddress", "ComputerName", "Status", "SoftwareName",
    "Version", "Compliant", "NewVersion", "ExitCode",
    "Comment", "AdminName", "Date"
)


# --- Dummy patch script that returns a configurable exit code ---
# This runs on the REMOTE side (or localhost) via Invoke-Command.
# It expects $Args[0] = config hashtable with .TestExitCode and .TestNewVersion.
$dummyPatchScriptStr = @'
$config = $Args[0]
$exitCode = $config.TestExitCode
if ($null -eq $exitCode) { $exitCode = 0 }

# Simulate install delay
Start-Sleep -Milliseconds 100

[PSCustomObject]@{
    ExitCode   = $exitCode
    NewVersion = $config.TestNewVersion
    Comment    = $null
}
'@


# --- Build argument sets helper ---
function Build-PatchArgs {
    param(
        [string[]]$Machines,
        [hashtable]$Config,
        [bool]$Force          = $false,
        [bool]$NoCopy         = $true,
        [string]$PatchScript  = $dummyPatchScriptStr,
        [array]$ScriptArgList = @(),
        [bool]$Isolated       = $false
    )

    $partial = [System.Collections.Concurrent.ConcurrentDictionary[string, hashtable]]::new()

    @(
        foreach ($m in $Machines) {
            , @(
                $m,               # $args[0]  = Computer
                $Config,          # $args[1]  = Config hashtable
                $Force,           # $args[2]  = Force
                $NoCopy,          # $args[3]  = NoCopy
                $PatchScript,     # $args[4]  = PatchScript as string
                $ScriptArgList,   # $args[5]  = ArgumentList for remote script
                "",               # $args[6]  = DNS suffix
                "2026/03/24",     # $args[7]  = Date
                "TestAdmin",      # $args[8]  = AdminName
                $Isolated,        # $args[9]  = Isolated
                $partial          # $args[10] = Partial results dict
            )
        }
    )
}


# --- Base test configs ---

# File-based version check, software doesn't exist (will return "Not Installed")
$cfgNotInstalled = @{
    Tag           = $null
    Software      = "FakeSoftwareXYZ"
    SoftwareName  = "Fake Software"
    CompliantVer  = "1.0.0.0"
    SoftwarePaths = "C:\NonExistent\Fake\test.exe"
    VersionType   = "File"
    RegistryKey   = $null
    PatchPath     = $null
}

# Registry-based version check for Edge (likely installed on test machine)
$cfgEdgeReg = @{
    Tag           = @("RegVersion")
    Software      = "Microsoft Edge"
    SoftwareName  = "Microsoft Edge"
    CompliantVer  = "1.0.0.0"           # Very low so anything passes
    SoftwarePaths = $null
    VersionType   = $null
    RegistryKey   = $null
    PatchPath     = $null
}

#endregion ===== SETUP =====



#region ===== SCRIPTBLOCKS =====

# ---------------------------------------------------------------------------
# Production-equivalent pipeline scriptblock.
# Mirrors the scriptblock from Invoke-Patch.ps1 (Phases 1-6).
# If the production code changes, this must be updated to match.
# ---------------------------------------------------------------------------

$pipelineScriptBlock = {

    $computer       = $args[0]
    $config         = $args[1]
    $force          = $args[2]
    $noCopy         = $args[3]
    $patchScriptStr = $args[4]
    $scriptArgList  = $args[5]
    $dnsSuffix      = $args[6]
    $date           = $args[7]
    $adminName      = $args[8]
    $isolated       = $args[9]
    $partialResults = $args[10]

    # Inline helper -- strips WinRM boilerplate from exception messages.
    # Must be defined inside the scriptblock; runspaces cannot see the
    # caller's imported modules (InitialSessionState::CreateDefault).
    function _CompressError ([string]$Msg) {
        $Msg = $Msg -replace 'Processing data from remote server \S+ failed with the following error message:\s*', ''
        $Msg = $Msg -replace 'Connecting to remote server \S+ failed with the following error message\s*:\s*', ''
        $Msg = $Msg -replace '\s*For more information, see the about_Remote_Troubleshooting Help topic\.', ''
        $Msg = $Msg -replace '\r?\n', ' '
        $Msg = $Msg -replace '\s{2,}', ' '
        return $Msg.Trim()
    }

    $_scriptStart = [DateTime]::Now

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

    if ($computer -match '\.') {
        $result.IPAddress = $computer
    }
    else {
        $result.ComputerName = $computer
    }


    #--- PHASE 1: Ping ---
    $PhaseTracker[$computer] = "Pinging"

    if ($isolated) {
        $result.Status = "Isolated"
        $ipAddr = $null
    }
    else {
        $pingResult = Test-Connection -ComputerName $computer -Count 1 -ErrorAction SilentlyContinue
        if ($null -eq $pingResult) {
            $result.Status = "Offline"
            return $result
        }
        $result.Status = "Online"

        $ipAddr = $null
        if ($null -ne $pingResult.IPV4Address) {
            $ipAddr = $pingResult.IPV4Address.IPAddressToString
        }
        elseif ($null -ne $pingResult.ProtocolAddress) {
            $ipAddr = $pingResult.ProtocolAddress
        }
    }


    #--- PHASE 2: DNS Resolution ---
    $PhaseTracker[$computer] = "DNS Lookup"

    if ($computer -match '\.') {
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
        $result.ComputerName = $computer
        if ($null -ne $ipAddr) {
            $result.IPAddress = $ipAddr
        }
    }

    $targetName = if ($null -ne $result.ComputerName) { $result.ComputerName } else { $result.IPAddress }
    if ($null -eq $targetName) {
        $result.Comment = "DNS Request Failed"
        return $result
    }

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
        $pathsStr    = [string]$config.SoftwarePaths
        $versionType = $config.VersionType

        try {
            $verResult = Invoke-Command -ComputerName $targetName -ScriptBlock {
                param($PathsString, $VerType)

                $versions = @()
                $targetUsers = @()

                $paths = @()
                foreach ($chunk in ($PathsString -split "C:")) {
                    if ($chunk -eq "") { continue }
                    $paths += "C:" + $chunk.Trim()
                }

                foreach ($path in $paths) {
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
                                    if ($VerType -eq 'Product') { $ver = $fileItem.VersionInfo.ProductVersion }
                                    else { $ver = $fileItem.VersionInfo.FileVersionRaw }
                                    if ($null -eq $ver) { $ver = $fileItem.VersionInfo.ProductVersion }
                                    if ($null -eq $ver) { $ver = $fileItem.VersionInfo.FileVersion }
                                    if ($null -eq $ver) { continue }
                                    if ($ver.GetType().Name -match 'string') { $ver = [version]($ver.Replace(',','.')) }
                                    $versions += $ver
                                    $targetUsers += $usr
                                }
                                elseif ($item.Mode -match 'd') { $targetUsers += $usr }
                            }
                        }
                    }
                    elseif (Test-Path $path) {
                        $item = Get-Item $path -Force -ErrorAction SilentlyContinue
                        if ($null -eq $item) { continue }
                        if (($item.Mode -match 'a') -or ($item.Mode -eq '------')) {
                            $fileItem = Get-ChildItem $path -Force -ErrorAction SilentlyContinue
                            if ($VerType -eq 'Product') { $ver = $fileItem.VersionInfo.ProductVersion }
                            else { $ver = $fileItem.VersionInfo.FileVersionRaw }
                            if ($null -eq $ver) { $ver = $fileItem.VersionInfo.ProductVersion }
                            if ($null -eq $ver) { $ver = $fileItem.VersionInfo.FileVersion }
                            if ($null -eq $ver) { continue }
                            if ($ver.GetType().Name -match 'string') { $ver = [version]($ver.Replace(',','.')) }
                            $versions += $ver
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

        if (Test-Path $destPath) {
            try {
                $destMeasure = Get-ChildItem $destPath -Recurse -ErrorAction SilentlyContinue | Measure-Object -Sum Length
                if (($destMeasure.Count -eq $config.OriginFileCount) -and ($destMeasure.Sum -eq $config.OriginFileSize)) {
                    $copyRequired = $false
                }
            }
            catch {}
        }

        if ($copyRequired) {
            if (Test-Path $destPath -PathType Leaf) {
                Remove-Item $destPath -Force > $null
            }

            $robocopyArgs = @(
                "`"$patchPath`""
                "`"$destPath`""
                '/E'
                '/R:3'
                '/W:5'
                '/MT:4'
                '/NP'
                '/NDL'
                '/NFL'
                '/NJH'
                '/NJS'
            )
            $robocopyOutput = & robocopy @robocopyArgs 2>&1
            $robocopyExit   = $LASTEXITCODE

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

    if ($tag -contains "RegVersion") {
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
            if ($null -ne $result.Comment) {
                $result.Comment = $result.Comment + " | New Version Check Failed: $_"
            }
            else {
                $result.Comment = "New Version Check Failed: $_"
            }
        }
    }
    else {
        if ($null -ne $installResult.NewVersion) {
            [array]$result.NewVersion = $installResult.NewVersion
        }
    }

    # Update Avg Success in the progress display
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
}


# --- Focused test scriptblocks ---

$compressErrorTestBlock = {
    function _CompressError ([string]$Msg) {
        $Msg = $Msg -replace 'Processing data from remote server \S+ failed with the following error message:\s*', ''
        $Msg = $Msg -replace 'Connecting to remote server \S+ failed with the following error message\s*:\s*', ''
        $Msg = $Msg -replace '\s*For more information, see the about_Remote_Troubleshooting Help topic\.', ''
        $Msg = $Msg -replace '\r?\n', ' '
        $Msg = $Msg -replace '\s{2,}', ' '
        return $Msg.Trim()
    }

    $sampleError = "Connecting to remote server FAKEPC failed with the following error message : " +
        "The WinRM client cannot process the request. " +
        "For more information, see the about_Remote_Troubleshooting Help topic."

    [PSCustomObject]@{
        ComputerName = $args[0]
        Compressed   = (_CompressError $sampleError)
    }
}

$catchPatternTestBlock = {
    function _CompressError ([string]$Msg) {
        $Msg = $Msg -replace 'Processing data from remote server \S+ failed with the following error message:\s*', ''
        $Msg = $Msg -replace 'Connecting to remote server \S+ failed with the following error message\s*:\s*', ''
        $Msg = $Msg -replace '\s*For more information, see the about_Remote_Troubleshooting Help topic\.', ''
        $Msg = $Msg -replace '\r?\n', ' '
        $Msg = $Msg -replace '\s{2,}', ' '
        return $Msg.Trim()
    }

    $result = [PSCustomObject]@{
        ComputerName = $args[0]
        Status       = "Online"
        Comment      = $null
    }

    try {
        throw "Connecting to remote server $($args[0]) failed with the following error message : " +
            "The client cannot connect to the destination specified in the request. " +
            "For more information, see the about_Remote_Troubleshooting Help topic."
    }
    catch {
        $result.Comment = "Version Check Failed: $(_CompressError "$_")"
        return $result
    }
}

# REGRESSION: calls non-existent function in catch block (the exact bug)
$regressionBugBlock = {
    $result = [PSCustomObject]@{
        ComputerName = $args[0]
        Status       = "Online"
        Comment      = $null
    }

    try {
        throw "Simulated WinRM error"
    }
    catch {
        $result.Comment = "Failed: $(Compress-ExceptionMessage "$_")"
        return $result
    }
}

# Compliance logic test block (no network, tests version comparison only)
$complianceTestBlock = {
    function _CompressError ([string]$Msg) { return $Msg }

    $config  = $args[1]
    $testVer = $args[2]

    $result = [PSCustomObject]@{
        ComputerName = $args[0]
        Status       = "Online"
        SoftwareName = $config.SoftwareName
        Version      = $testVer
        Compliant    = $null
        Comment      = $null
    }

    $compliantVer = $config.CompliantVer

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
    }

    return $result
}

#endregion ===== SCRIPTBLOCKS =====



# ======================================================================
#  TEST 1 : Offline Machine Path (Phase 1 - Ping)
# ======================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 1: Offline Machine Path (Phase 1)" -ForegroundColor Yellow
Write-Host "  Bogus hostname -> ping fails -> Status = Offline, early exit." -ForegroundColor Gray
Write-Host ""

$a1 = Build-PatchArgs -Machines @("FAKE-OFFLINE-99999") -Config $cfgNotInstalled
$r1 = @(Invoke-RunspacePool -ScriptBlock $pipelineScriptBlock -ArgumentList $a1 @common)

Assert-Test "Test 1" "Returns 1 result"          -Expected 1                    -Actual $r1.Count
Assert-Test "Test 1" "Status = Offline"           -Expected "Offline"            -Actual $r1[0].Status
Assert-Test "Test 1" "ComputerName populated"     -Expected "FAKE-OFFLINE-99999" -Actual $r1[0].ComputerName
Assert-Test "Test 1" "SoftwareName populated"     -Expected "Fake Software"      -Actual $r1[0].SoftwareName
Assert-Test "Test 1" "Version is null (no check)" -Actual $r1[0].Version        -CompareMode IsNull
Assert-Test "Test 1" "ExitCode is null (no patch)" -Actual $r1[0].ExitCode      -CompareMode IsNull

$r1Props = @($r1[0].PSObject.Properties.Name)
$missing1 = @($expectedProperties | Where-Object { $_ -notin $r1Props })
Assert-Test "Test 1" "All 11 properties present"  -Expected 0 -Actual $missing1.Count
Write-Host ""



# ======================================================================
#  TEST 2 : _CompressError Works in Runspace
# ======================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 2: _CompressError Works in Runspace" -ForegroundColor Yellow
Write-Host "  Inline function must strip WinRM boilerplate inside a runspace." -ForegroundColor Gray
Write-Host ""

$a2 = @( , @("COMPRESS-TEST") )
$r2 = @(Invoke-RunspacePool -ScriptBlock $compressErrorTestBlock -ArgumentList $a2 @common)

Assert-Test "Test 2" "Returns 1 result"           -Expected 1 -Actual $r2.Count
Assert-Test "Test 2" "Boilerplate stripped"        -Expected "The WinRM client cannot process the request." -Actual $r2[0].Compressed
Assert-Test "Test 2" "No 'Connecting to' prefix"   -Expected "True" -Actual "$($r2[0].Compressed -notmatch 'Connecting to remote server')"
Assert-Test "Test 2" "No troubleshooting footer"   -Expected "True" -Actual "$($r2[0].Compressed -notmatch 'about_Remote')"
Write-Host ""



# ======================================================================
#  TEST 3 : WinRM Error Catch Block (Phase 3)
# ======================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 3: WinRM Error Catch Block (Phase 3)" -ForegroundColor Yellow
Write-Host "  Simulated throw -> catch with _CompressError -> result returned." -ForegroundColor Gray
Write-Host ""

$a3 = @( , @("CATCH-TEST") )
$r3 = @(Invoke-RunspacePool -ScriptBlock $catchPatternTestBlock -ArgumentList $a3 @common)

Assert-Test "Test 3" "Returns 1 result"                -Expected 1        -Actual $r3.Count
Assert-Test "Test 3" "Status = Online"                  -Expected "Online" -Actual $r3[0].Status
Assert-Test "Test 3" "Comment is not null"              -Actual $r3[0].Comment -CompareMode NotNull
Assert-Test "Test 3" "Comment has prefix"               -Expected "Version Check Failed:" -Actual $r3[0].Comment -CompareMode Contains
Assert-Test "Test 3" "Boilerplate stripped from Comment" -Expected "True" -Actual "$($r3[0].Comment -notmatch 'about_Remote')"
Assert-Test "Test 3" "Actual error preserved"           -Expected "cannot connect to the destination" -Actual $r3[0].Comment -CompareMode Contains
Write-Host ""



# ======================================================================
#  TEST 4 : REGRESSION -- Non-Existent Function in Catch Block
# ======================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 4: REGRESSION -- Non-Existent Function in Catch Block" -ForegroundColor Yellow
Write-Host "  Old bug: Compress-ExceptionMessage missing -> silent drop." -ForegroundColor Gray
Write-Host "  Safety net must catch this and produce a result." -ForegroundColor Gray
Write-Host ""

$a4 = @( , @("REGRESSION-001") )
$error4 = $null

try {
    $r4 = @(Invoke-RunspacePool -ScriptBlock $regressionBugBlock -ArgumentList $a4 @common)
}
catch {
    $error4 = $_
}

Assert-Test "Test 4" "No unhandled exception"     -Expected "True"            -Actual "$($null -eq $error4)"
Assert-Test "Test 4" "Returns 1 result (not 0)"   -Expected 1                -Actual $r4.Count
Assert-Test "Test 4" "ComputerName preserved"     -Expected "REGRESSION-001" -Actual $r4[0].ComputerName
Assert-Test "Test 4" "Comment is not null"         -Actual $r4[0].Comment     -CompareMode NotNull
Write-Host ""



# ======================================================================
#  TEST 5 : Result Schema Completeness
# ======================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 5: Result Schema Completeness" -ForegroundColor Yellow
Write-Host "  Multiple machines in different states -- all must have full schema." -ForegroundColor Gray
Write-Host ""

$a5 = Build-PatchArgs -Machines @("FAKE-AAA", "FAKE-BBB", "localhost") -Config $cfgNotInstalled
$r5 = @(Invoke-RunspacePool -ScriptBlock $pipelineScriptBlock -ArgumentList $a5 @common)

Assert-Test "Test 5" "3 targets -> 3 results" -Expected 3 -Actual $r5.Count

$schemaOK = $true
foreach ($r in $r5) {
    $props = @($r.PSObject.Properties.Name)
    $miss = @($expectedProperties | Where-Object { $_ -notin $props })
    if ($miss.Count -gt 0) {
        $schemaOK = $false
        Write-Host "    Missing on $($r.ComputerName): $($miss -join ', ')" -ForegroundColor Red
    }
}
Assert-Test "Test 5" "All results have full 11-property schema" -Expected "True" -Actual "$schemaOK"
Write-Host ""



# ======================================================================
#  TEST 6 : No Silent Drops (Online + Failed = Comment Populated)
# ======================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 6: No Silent Drops" -ForegroundColor Yellow
Write-Host "  If Status=Online and Version is null, Comment must explain why." -ForegroundColor Gray
Write-Host ""

$a6 = Build-PatchArgs -Machines @("localhost") -Config $cfgNotInstalled
$r6 = @(Invoke-RunspacePool -ScriptBlock $pipelineScriptBlock -ArgumentList $a6 @common)

Assert-Test "Test 6" "Returns 1 result" -Expected 1 -Actual $r6.Count

$r6i = $r6[0]
$hasVer = ($null -ne $r6i.Version -and "$($r6i.Version)" -ne "")
$hasCom = ($null -ne $r6i.Comment -and "$($r6i.Comment)" -ne "")

if ($r6i.Status -eq "Online" -or $r6i.Status -eq "Isolated") {
    $either = $hasVer -or $hasCom
    Assert-Test "Test 6" "Online -> Version or Comment populated (never both blank)" -Expected "True" -Actual "$either"
}
else {
    Write-Host "    (localhost reported Offline -- unusual but not a test failure)" -ForegroundColor Yellow
}
Write-Host ""



# ======================================================================
#  TEST 7 : Compliance Logic
# ======================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 7: Compliance Logic" -ForegroundColor Yellow
Write-Host "  Version comparison against CompliantVer." -ForegroundColor Gray
Write-Host ""

$compCfg = @{ SoftwareName = "Compliance Test"; CompliantVer = "10.0.0.0" }

$a7a = @( , @("COMP-BELOW", $compCfg, "9.0.0.0") )
$r7a = @(Invoke-RunspacePool -ScriptBlock $complianceTestBlock -ArgumentList $a7a @common)
Assert-Test "Test 7a" "Below compliant -> Compliant = False" -Expected "False" -Actual "$($r7a[0].Compliant)"

$a7b = @( , @("COMP-EXACT", $compCfg, "10.0.0.0") )
$r7b = @(Invoke-RunspacePool -ScriptBlock $complianceTestBlock -ArgumentList $a7b @common)
Assert-Test "Test 7b" "At compliant -> Compliant = True" -Expected "True" -Actual "$($r7b[0].Compliant)"

$a7c = @( , @("COMP-ABOVE", $compCfg, "11.0.0.0") )
$r7c = @(Invoke-RunspacePool -ScriptBlock $complianceTestBlock -ArgumentList $a7c @common)
Assert-Test "Test 7c" "Above compliant -> Compliant = True" -Expected "True" -Actual "$($r7c[0].Compliant)"

$a7d = @( , @("COMP-NONE", $compCfg, "Not Installed") )
$r7d = @(Invoke-RunspacePool -ScriptBlock $complianceTestBlock -ArgumentList $a7d @common)
Assert-Test "Test 7d" "Not Installed -> Compliant is null" -Actual $r7d[0].Compliant -CompareMode IsNull
Write-Host ""



# ======================================================================
#  TEST 8 : Compliant Machine Exits Early
# ======================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 8: Compliant Machine Exits Early" -ForegroundColor Yellow
Write-Host "  If version >= compliant and no -Force, pipeline exits after Phase 3." -ForegroundColor Gray
Write-Host ""

if ($winrmInRunspace) {
    # Use Edge against localhost (should be compliant with CompliantVer 1.0.0.0)
    $a8 = Build-PatchArgs -Machines @("localhost") -Config $cfgEdgeReg -Force $false
    $r8 = @(Invoke-RunspacePool -ScriptBlock $pipelineScriptBlock -ArgumentList $a8 @common)

    Assert-Test "Test 8" "Returns 1 result"         -Expected 1      -Actual $r8.Count
    Assert-Test "Test 8" "Compliant = True"          -Expected "True" -Actual "$($r8[0].Compliant)"
    Assert-Test "Test 8" "ExitCode is null (no patch)" -Actual $r8[0].ExitCode -CompareMode IsNull
    Assert-Test "Test 8" "NewVersion is null (skipped)" -Actual $r8[0].NewVersion -CompareMode IsNull
    if ($r8[0].Compliant -ne $true -and $r8[0].Comment) {
        Write-Host "    Comment: $($r8[0].Comment)" -ForegroundColor DarkYellow
    }
}
else {
    Write-Host "    (Skipped -- WinRM from runspace unavailable)" -ForegroundColor Yellow
}
Write-Host ""



# ======================================================================
#  TEST 9 : -Force Flag Overrides Compliance
# ======================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 9: -Force Flag Overrides Compliance" -ForegroundColor Yellow
Write-Host "  Compliant machine with -Force proceeds to Phase 5 (patching)." -ForegroundColor Gray
Write-Host ""

if ($winrmInRunspace) {
    # Force = true, NoCopy = true, dummy patch script returns exit 0
    $cfgEdgeForce = $cfgEdgeReg.Clone()
    $cfgEdgeForce.TestExitCode  = 0
    $cfgEdgeForce.TestNewVersion = "99.0.0.0"

    $a9 = Build-PatchArgs -Machines @("localhost") -Config $cfgEdgeForce -Force $true -NoCopy $true -ScriptArgList @(,$cfgEdgeForce)
    $r9 = @(Invoke-RunspacePool -ScriptBlock $pipelineScriptBlock -ArgumentList $a9 @common)

    Assert-Test "Test 9" "Returns 1 result"                -Expected 1      -Actual $r9.Count
    Assert-Test "Test 9" "Compliant = True (was compliant)" -Expected "True" -Actual "$($r9[0].Compliant)"
    Assert-Test "Test 9" "ExitCode populated (patch ran)"   -Actual $r9[0].ExitCode -CompareMode NotNull
    Assert-Test "Test 9" "ExitCode = 0"                     -Expected "0"    -Actual "$($r9[0].ExitCode)"
    if ($r9[0].Comment) {
        Write-Host "    Comment: $($r9[0].Comment)" -ForegroundColor DarkYellow
    }
}
else {
    Write-Host "    (Skipped -- WinRM from runspace unavailable)" -ForegroundColor Yellow
}
Write-Host ""



# ======================================================================
#  TEST 10 : File Copy Phase (Phase 4 - Robocopy)
# ======================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 10: File Copy Phase (Phase 4)" -ForegroundColor Yellow
Write-Host "  Robocopy a temp folder to localhost C$ share." -ForegroundColor Gray
Write-Host ""

if ($winrmInRunspace) {
    # Create a temp source dir with a test file
    $tempSource = Join-Path $env:TEMP "TestPatchPipeline_Source_$(Get-Random)"
    $tempDest   = Join-Path "C:\Temp" (Split-Path $tempSource -Leaf)
    mkdir $tempSource -Force > $null
    "test content" | Set-Content "$tempSource\testfile.txt"

    $cfgCopy = $cfgEdgeReg.Clone()
    $cfgCopy.PatchPath       = $tempSource
    $cfgCopy.OriginFileCount = 1
    $cfgCopy.OriginFileSize  = (Get-Item "$tempSource\testfile.txt").Length
    $cfgCopy.OriginHashes    = @()
    $cfgCopy.TestExitCode    = 0
    $cfgCopy.TestNewVersion  = "99.0.0.0"

    $a10 = Build-PatchArgs -Machines @("localhost") -Config $cfgCopy -Force $true -NoCopy $false -ScriptArgList @(,$cfgCopy)
    $r10 = @(Invoke-RunspacePool -ScriptBlock $pipelineScriptBlock -ArgumentList $a10 @common)

    Assert-Test "Test 10" "Returns 1 result"     -Expected 1 -Actual $r10.Count
    Assert-Test "Test 10" "No copy failure"       -Expected "True" -Actual "$($r10[0].Comment -notmatch 'Copy Failed')"

    # Check that the file arrived
    $fileCopied = Test-Path "$tempDest\testfile.txt"
    Assert-Test "Test 10" "File copied to destination" -Expected "True" -Actual "$fileCopied"
    if ($r10[0].Comment) {
        Write-Host "    Comment: $($r10[0].Comment)" -ForegroundColor DarkYellow
    }

    # Cleanup
    Remove-Item $tempSource -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item $tempDest   -Recurse -Force -ErrorAction SilentlyContinue
}
else {
    Write-Host "    (Skipped -- WinRM from runspace unavailable)" -ForegroundColor Yellow
}
Write-Host ""



# ======================================================================
#  TEST 11 : Patch Execution Phase (Phase 5)
# ======================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 11: Patch Execution Phase (Phase 5)" -ForegroundColor Yellow
Write-Host "  Dummy patch script returns configurable exit code via Invoke-Command." -ForegroundColor Gray
Write-Host ""

if ($winrmInRunspace) {
    $cfgPatch = $cfgEdgeReg.Clone()
    $cfgPatch.TestExitCode  = 42
    $cfgPatch.TestNewVersion = "99.0.0.0"

    $a11 = Build-PatchArgs -Machines @("localhost") -Config $cfgPatch -Force $true -NoCopy $true -ScriptArgList @(,$cfgPatch)
    $r11 = @(Invoke-RunspacePool -ScriptBlock $pipelineScriptBlock -ArgumentList $a11 @common)

    Assert-Test "Test 11" "Returns 1 result"      -Expected 1    -Actual $r11.Count
    Assert-Test "Test 11" "ExitCode = 42"          -Expected "42" -Actual "$($r11[0].ExitCode)"
    if ($r11[0].Comment) {
        Write-Host "    Comment: $($r11[0].Comment)" -ForegroundColor DarkYellow
    }
}
else {
    Write-Host "    (Skipped -- WinRM from runspace unavailable)" -ForegroundColor Yellow
}
Write-Host ""



# ======================================================================
#  TEST 12 : Exit Code Handling
# ======================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 12: Exit Code Handling" -ForegroundColor Yellow
Write-Host "  Various exit codes: 0, 3010, 1603." -ForegroundColor Gray
Write-Host ""

if ($winrmInRunspace) {
    foreach ($testCase in @(
        @{ Code = 0;    Label = "success (0)" },
        @{ Code = 3010; Label = "reboot required (3010)" },
        @{ Code = 1603; Label = "fatal install error (1603)" }
    )) {
        $cfgEC = $cfgEdgeReg.Clone()
        $cfgEC.TestExitCode  = $testCase.Code
        $cfgEC.TestNewVersion = "99.0.0.0"

        $aEC = Build-PatchArgs -Machines @("localhost") -Config $cfgEC -Force $true -NoCopy $true -ScriptArgList @(,$cfgEC)
        $rEC = @(Invoke-RunspacePool -ScriptBlock $pipelineScriptBlock -ArgumentList $aEC @common)

        Assert-Test "Test 12" "Exit $($testCase.Label) -> ExitCode = $($testCase.Code)" `
            -Expected "$($testCase.Code)" -Actual "$($rEC[0].ExitCode)"
    }
}
else {
    Write-Host "    (Skipped -- WinRM from runspace unavailable)" -ForegroundColor Yellow
}
Write-Host ""



# ======================================================================
#  TEST 13 : Post-Install Verification (Phase 6)
# ======================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 13: Post-Install Verification (Phase 6)" -ForegroundColor Yellow
Write-Host "  After patch, NewVersion is populated from re-query or install result." -ForegroundColor Gray
Write-Host ""

if ($winrmInRunspace) {
    # RegVersion path: Phase 6 re-queries registry, so NewVersion should match
    # whatever Edge version is actually installed
    $cfgVerify = $cfgEdgeReg.Clone()
    $cfgVerify.TestExitCode  = 0
    $cfgVerify.TestNewVersion = $null  # RegVersion path ignores this

    $a13 = Build-PatchArgs -Machines @("localhost") -Config $cfgVerify -Force $true -NoCopy $true -ScriptArgList @(,$cfgVerify)
    $r13 = @(Invoke-RunspacePool -ScriptBlock $pipelineScriptBlock -ArgumentList $a13 @common)

    Assert-Test "Test 13" "Returns 1 result"         -Expected 1 -Actual $r13.Count
    Assert-Test "Test 13" "NewVersion is populated"   -Actual $r13[0].NewVersion -CompareMode NotNull
    Assert-Test "Test 13" "Version is populated"      -Actual $r13[0].Version -CompareMode NotNull
    if ($r13[0].Comment) {
        Write-Host "    Comment: $($r13[0].Comment)" -ForegroundColor DarkYellow
    }
}
else {
    Write-Host "    (Skipped -- WinRM from runspace unavailable)" -ForegroundColor Yellow
}
Write-Host ""



# ======================================================================
#  TEST 14 : Partial Result Recovery (timeout mid-pipeline)
# ======================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 14: Partial Result Recovery (Timeout)" -ForegroundColor Yellow
Write-Host "  Scriptblock that hangs mid-pipeline -> partial data survives." -ForegroundColor Gray
Write-Host ""

# Scriptblock that completes Phase 1-2, saves partial data, then hangs
$hangingBlock = {
    $computer       = $args[0]
    $config         = $args[1]
    $partialResults = $args[10]

    function _CompressError ([string]$Msg) { return $Msg }

    $result = [PSCustomObject]@{
        IPAddress    = $null
        ComputerName = $computer
        Status       = "Online"
        SoftwareName = $config.SoftwareName
        Version      = $null
        Compliant    = $null
        NewVersion   = $null
        ExitCode     = $null
        Comment      = $null
        AdminName    = "TestAdmin"
        Date         = "2026/03/24"
    }

    # Save partial data (simulates completing Phase 2)
    $partialResults[$computer] = @{
        IPAddress    = "127.0.0.1"
        ComputerName = $computer
        Status       = "Online"
    }

    $PhaseTracker[$computer] = "Patching"

    # Hang forever (will be killed by timeout)
    Start-Sleep -Seconds 300
    return $result
}

$partialDict = [System.Collections.Concurrent.ConcurrentDictionary[string, hashtable]]::new()
$a14 = @(
    , @(
        "TIMEOUT-001",      # $args[0]
        $cfgNotInstalled,   # $args[1]
        $false,             # $args[2]  Force
        $true,              # $args[3]  NoCopy
        "",                 # $args[4]  PatchScript
        @(),                # $args[5]  ScriptArgList
        "",                 # $args[6]  DNS suffix
        "2026/03/24",       # $args[7]  Date
        "TestAdmin",        # $args[8]  AdminName
        $false,             # $args[9]  Isolated
        $partialDict        # $args[10] Partial results
    )
)

$r14 = @(Invoke-RunspacePool -ScriptBlock $hangingBlock -ArgumentList $a14 -ThrottleLimit 1 -TimeoutMinutes 1 -ActivityName "Test")

Assert-Test "Test 14" "Returns 1 result (not lost)"     -Expected 1             -Actual $r14.Count
Assert-Test "Test 14" "ComputerName = TIMEOUT-001"       -Expected "TIMEOUT-001" -Actual $r14[0].ComputerName
Assert-Test "Test 14" "Comment contains 'Task Stopped'"  -Expected "Task Stopped" -Actual $r14[0].Comment -CompareMode Contains

# Check that partial data was saved
$p14 = $null
$partialDict.TryGetValue("TIMEOUT-001", [ref]$p14) > $null
Assert-Test "Test 14" "Partial data saved before timeout" -Expected "True" -Actual "$($null -ne $p14)"
if ($null -ne $p14) {
    Assert-Test "Test 14" "Partial IP preserved"    -Expected "127.0.0.1"   -Actual $p14.IPAddress
    Assert-Test "Test 14" "Partial Status preserved" -Expected "Online"      -Actual $p14.Status
}
Write-Host ""



# ======================================================================
#  TEST 15 : Isolated Mode (skip ping)
# ======================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 15: Isolated Mode" -ForegroundColor Yellow
Write-Host "  Isolated = true -> ping is skipped, Status = 'Isolated'." -ForegroundColor Gray
Write-Host ""

if ($winrmAvailable) {
    $a15 = Build-PatchArgs -Machines @("localhost") -Config $cfgNotInstalled -Isolated $true
    $r15 = @(Invoke-RunspacePool -ScriptBlock $pipelineScriptBlock -ArgumentList $a15 @common)

    Assert-Test "Test 15" "Returns 1 result"     -Expected 1          -Actual $r15.Count
    Assert-Test "Test 15" "Status = Isolated"     -Expected "Isolated" -Actual $r15[0].Status
}
else {
    Write-Host "    (Skipped -- WinRM loopback unavailable)" -ForegroundColor Yellow
}
Write-Host ""



# ======================================================================
#  TEST 16 : Localhost Full Pipeline (all phases)
# ======================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 16: Localhost Full Pipeline" -ForegroundColor Yellow
Write-Host "  All 6 phases against localhost with Edge + dummy patch." -ForegroundColor Gray
Write-Host ""

if ($winrmInRunspace) {
    $cfgFull = $cfgEdgeReg.Clone()
    $cfgFull.TestExitCode    = 0
    $cfgFull.TestNewVersion  = $null       # Phase 6 re-queries registry
    $cfgFull.PatchPath       = $null       # Skip copy (NoCopy = true)

    $a16 = Build-PatchArgs -Machines @("localhost") -Config $cfgFull -Force $true -NoCopy $true -ScriptArgList @(,$cfgFull)
    $r16 = @(Invoke-RunspacePool -ScriptBlock $pipelineScriptBlock -ArgumentList $a16 @common)

    Assert-Test "Test 16" "Returns 1 result"          -Expected 1        -Actual $r16.Count
    Assert-Test "Test 16" "Status = Online"            -Expected "Online" -Actual $r16[0].Status
    Assert-Test "Test 16" "Version populated (Phase 3)" -Actual $r16[0].Version    -CompareMode NotNull
    Assert-Test "Test 16" "ExitCode populated (Phase 5)" -Actual $r16[0].ExitCode  -CompareMode NotNull
    Assert-Test "Test 16" "NewVersion populated (Phase 6)" -Actual $r16[0].NewVersion -CompareMode NotNull
    Assert-Test "Test 16" "Compliant is set"            -Actual $r16[0].Compliant  -CompareMode NotNull
    if ($r16[0].Comment) {
        Write-Host "    Comment: $($r16[0].Comment)" -ForegroundColor DarkYellow
    }
    Assert-Test "Test 16" "AdminName = TestAdmin"       -Expected "TestAdmin" -Actual $r16[0].AdminName

    # Full schema check
    $r16Props = @($r16[0].PSObject.Properties.Name)
    $miss16 = @($expectedProperties | Where-Object { $_ -notin $r16Props })
    Assert-Test "Test 16" "All 11 properties present" -Expected 0 -Actual $miss16.Count
}
else {
    Write-Host "    (Skipped -- WinRM from runspace unavailable)" -ForegroundColor Yellow
}
Write-Host ""



# ======================================================================
#  SUMMARY
# ======================================================================

$sw.Stop()

$passed = @($script:TestResults | Where-Object { $_.Status -eq 'PASS' }).Count
$failed = @($script:TestResults | Where-Object { $_.Status -eq 'FAIL' }).Count
$total  = $script:TestResults.Count

Write-Host ("=" * 72) -ForegroundColor Cyan
Write-Host "  TEST SUMMARY" -ForegroundColor Cyan
Write-Host ("=" * 72) -ForegroundColor Cyan
Write-Host ""
Write-Host "  Total  : $total assertions"
Write-Host "  Passed : " -NoNewline
Write-Host "$passed" -ForegroundColor Green
Write-Host "  Failed : " -NoNewline
Write-Host "$failed" -ForegroundColor $(if ($failed -gt 0) { 'Red' } else { 'Green' })
if (-not $winrmInRunspace) {
    Write-Host "  Skipped: Tests 8-13, 16 (WinRM loopback unavailable from runspace threads)" -ForegroundColor Yellow
    Write-Host "           Run elevated + 'Enable-PSRemoting -Force' + set TrustedHosts to fix." -ForegroundColor Yellow
    Write-Host "           See probe output above for full instructions." -ForegroundColor Yellow
}
Write-Host "  Time   : $([math]::Round($sw.Elapsed.TotalSeconds, 1)) seconds"
Write-Host ""

if ($failed -gt 0) {
    Write-Host "  FAILED ASSERTIONS:" -ForegroundColor Red
    Write-Host ""
    foreach ($f in ($script:TestResults | Where-Object { $_.Status -eq 'FAIL' })) {
        Write-Host "    [$($f.Test)] $($f.Description)" -ForegroundColor Red
        Write-Host "      Expected: $($f.Expected)  |  Actual: $($f.Actual)" -ForegroundColor DarkRed
    }
    Write-Host ""
}
