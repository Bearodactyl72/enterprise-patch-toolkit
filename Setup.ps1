# DOTS formatting comment

# Pulls powershell script root
if ($psISE) {
    $path = Split-Path -Path $psISE.CurrentFile.FullPath
}
else {
    $path = $PSScriptRoot
}

if ($null -eq $path) {
    Write-Warning "Path could not be determined"
    Start-Sleep 5
    break
}


# Checks for current PowerShell profile
$copyRequired = $false
if (!(Test-Path "$env:USERPROFILE\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1")) {
    $copyRequired = $true
}
else {
    $lengthPSNew = (Get-Item "$path\Profiles\Microsoft.PowerShell_profile.ps1").Length
    $lengthPSOld = (Get-Item "$env:USERPROFILE\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1").Length
    if ($lengthPSNew -ne $lengthPSOld) {
        $copyRequired = $true
    }
}

if (!(Test-Path "$env:USERPROFILE\Documents\WindowsPowerShell\Microsoft.PowerShellISE_profile.ps1")) {
    $copyRequired = $true
}
else {
    $lengthISENew = (Get-Item "$path\Profiles\Microsoft.PowerShellISE_profile.ps1").Length
    $lengthISEOld = (Get-Item "$env:USERPROFILE\Documents\WindowsPowerShell\Microsoft.PowerShellISE_profile.ps1").Length
    if ($lengthISENew -ne $lengthISEOld) {
        $copyRequired = $true
    }
}


# Builds new PowerShell profile if required
if ($copyRequired) {
    Write-Host "Building PowerShell Profiles"
    Copy-Item "$path\Profiles\*" -Destination "$env:USERPROFILE\Documents\WindowsPowerShell" -Recurse -Force
}


# Loads environment-specific config (domains, shares, trusted hosts, etc.)
# so the rest of setup is driven from data, not hardcoded paths. Falls
# back to warnings on a fresh-clone / non-domain machine so the repo
# still installs cleanly for a smoke test.
$envModulePath = "$path\Modules\RSL-Environment\RSL-Environment.psd1"
$rslConfig     = $null
$activeNetwork = $null
if (Test-Path $envModulePath) {
    Import-Module $envModulePath -Force
    $rslConfig     = Import-RSLEnvironment -RepoRoot $path
    $activeNetwork = Get-RSLActiveNetwork
}
else {
    Write-Warning "RSL-Environment module not found; environment-driven setup skipped."
}


# Adds paths to user %appdata%
if (!(Test-Path "$env:APPDATA\Patching" -PathType Container)) {
    mkdir "$env:APPDATA\Patching" > $null
}

$scriptPath = "ScriptPath : " + $path + "\Scripts"
$modulePath = "ModulePath : " + $path + "\Modules"

@($scriptPath, $modulePath) | Out-File -FilePath "$env:APPDATA\Patching\Paths.txt" -Force


# Maps the active network's patch share to the configured drive letter.
# Warn-and-continue on failure so the rest of setup still runs.
if ($rslConfig -and $activeNetwork) {
    $driveLetter = $rslConfig.MappedDriveLetter
    $healthPath  = "${driveLetter}:\$($rslConfig.ShareAnchorPath)"
    if (!(Test-Path $healthPath)) {
        if (Test-Path $activeNetwork.PatchShareUnc) {
            New-PSDrive -Name $driveLetter -PSProvider FileSystem `
                -Root $activeNetwork.PatchShareUnc -Persist -Scope Global | Out-Null
        }
        else {
            Write-Warning "Patch share $($activeNetwork.PatchShareUnc) is unreachable; drive not mapped."
        }
    }
}
else {
    Write-Warning "No active network profile matched USERDNSDOMAIN; patch share drive not mapped."
}


# Copies PSTools from the mapped share to the user's desktop. Only runs
# on hosts whose name matches one of the TrustedRunnerHosts regex
# patterns, so random laptops do not pull centrally-hosted tooling.
if (!(Test-Path "$env:USERPROFILE\Desktop\PSTools\PsExec.exe")) {
    $isTrusted = $false
    if ($rslConfig -and $rslConfig.TrustedRunnerHosts) {
        foreach ($pattern in $rslConfig.TrustedRunnerHosts) {
            if ($env:COMPUTERNAME -match $pattern) { $isTrusted = $true; break }
        }
    }

    if ($isTrusted -and $rslConfig) {
        $pstoolsSource = "$($rslConfig.MappedDriveLetter):\$($rslConfig.ShareAnchorPath)\$($rslConfig.CentralPSToolsPath)"
        if (Test-Path $pstoolsSource) {
            Copy-Item $pstoolsSource -Recurse -Destination "$env:USERPROFILE\Desktop"
        }
        else {
            Write-Warning "PSTools not found at $pstoolsSource; skipping desktop copy."
        }
    }
    else {
        Write-Warning "PSTools desktop copy skipped: host is not a trusted runner."
    }
}


# Unblocks the files in case they were marked as downloaded-from-Internet
# (happens on fresh clone from any remote source-control host).
$files = Get-ChildItem $path -Recurse | Where-Object {$_.Mode -eq "-a----"}

foreach ($file in $files) {
    Unblock-File -Path $file.FullName
}


# Sync Main-Switch with central (self-heals stale base, pulls new entries).
# Skipped silently if central share is offline.
$mainSwitchLocal   = "$path\Scripts\Main-Switch.ps1"
if ($rslConfig) {
    $mainSwitchCentral = "$($rslConfig.MappedDriveLetter):\$($rslConfig.ShareAnchorPath)\$($rslConfig.CentralMainSwitchPath)"
}
else {
    $mainSwitchCentral = ""
}
$mergeModule       = "$path\Modules\Merge-MainSwitch\Merge-MainSwitch.psm1"

if ((Test-Path $mainSwitchLocal) -and (Test-Path $mainSwitchCentral) -and (Test-Path $mergeModule)) {
    try {
        Import-Module $mergeModule -Force -ErrorAction Stop
        Receive-MainSwitch -LocalPath $mainSwitchLocal -CentralPath $mainSwitchCentral -ErrorAction Stop
    }
    catch {
        Write-Warning ("Main-Switch sync skipped: " + $_.Exception.Message)
    }
}


# Create desktop + Start Menu shortcuts for the patching GUI so admins
# can launch it without going through the CLI. Helper is idempotent --
# safe to re-run on every setup pass; it overwrites any existing .lnk
# with current paths so the shortcut self-heals if the repo moves.
$shortcutHelper = "$path\Scripts\Patching\GUI\Install-PatchGUIShortcut.ps1"
if (Test-Path $shortcutHelper) {
    try {
        & $shortcutHelper -RepoRoot $path -ErrorAction Stop
    }
    catch {
        Write-Warning ("Shortcut creation skipped: " + $_.Exception.Message)
    }
}


# Easy test if script has been run before
function Get-Thotfix {Write-Host '*bonk* go to horny jail'}

Write-Host ""
Write-Host "Script run successfully."
Start-Sleep 3
