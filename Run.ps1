# Run.ps1 -- Quick-launch scratch pad for script invocations
#
# Workflow:
#   1. Open this file in PowerShell ISE.
#   2. Press F5 to dot-source all the standalone scripts below (loading
#      their functions into the session) and hit the break statement.
#   3. Edit the arguments in the example blocks below to fit your target.
#   4. Select the line(s) you want to run and press F8.
#
# $TargetMachines is loaded automatically by the TargetMachines module
# from Desktop\Lists\Target_Machines.txt when the profile starts.
#   - To reload after updating the list file:   Import-TargetMachines
#   - To paste a list interactively:            Enter-TargetMachines

$ScriptRoot = $PSScriptRoot
if (-not $ScriptRoot) {
    # Running selected lines in ISE loses $PSScriptRoot; fall back to the
    # currently-open file's directory when ISE is the host.
    if ($psISE) { $ScriptRoot = Split-Path -Path $psISE.CurrentFile.FullPath }
}
. "$ScriptRoot\Scripts\Utility\Discovery\Get-LoggedInUser.ps1"
. "$ScriptRoot\Scripts\Utility\Discovery\Find-PatchContent.ps1"
. "$ScriptRoot\Scripts\Utility\Discovery\Get-StaleAsset.ps1"
. "$ScriptRoot\Scripts\Utility\Maintenance\Restart-Machine.ps1"
. "$ScriptRoot\Scripts\Utility\Maintenance\Repair-MachineHealth.ps1"
. "$ScriptRoot\Scripts\Utility\Maintenance\Repair-WindowsUpdate.ps1"
. "$ScriptRoot\Scripts\Utility\Cleanup\Clear-CcmCache.ps1"
. "$ScriptRoot\Scripts\Utility\Cleanup\Remove-StaleRegistryKey.ps1"
. "$ScriptRoot\Scripts\Utility\Cleanup\Start-UserProfileCleanup.ps1"
. "$ScriptRoot\Scripts\Utility\Remediation\Invoke-SoftwareUninstall.ps1"
. "$ScriptRoot\Scripts\Utility\Remediation\Repair-MSIUninstall.ps1"
. "$ScriptRoot\Scripts\Utility\Remediation\Invoke-Log4JRemediation.ps1"
. "$ScriptRoot\Scripts\Utility\Remediation\Test-TaniumQuarantine.ps1"

break

# =====================================================================
#  Discovery
# =====================================================================

# --- Get logged-in users ---
Get-LoggedInUser -ComputerName $TargetMachines


# --- Find patch content in ccmcache ---
Find-PatchContent -ComputerName $TargetMachines -SearchString "google"


# --- Get stale assets ---
Get-StaleAsset -ComputerName $targetMachines -IncludeDHCP


# =====================================================================
#  Maintenance
# =====================================================================

# --- Restart machines ---
Restart-Machine -ComputerName $TargetMachines -DelayMinutes 15


# --- Repair machine health (SFC, DISM, services, SCCM) ---
Repair-MachineHealth -ComputerName $TargetMachines


# --- Repair Windows Update (aggressive WU agent reset) ---
Repair-WindowsUpdate -ComputerName $TargetMachines



# =====================================================================
#  Cleanup
# =====================================================================

# --- Clear SCCM cache ---
Clear-CcmCache -ComputerName $TargetMachines


# --- Remove stale software registry keys ---
$softwareName = "Google Chrome"
$softwarePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
Remove-StaleRegistryKey -ComputerName $TargetMachines -SoftwareName $softwareName -FilePath $softwarePath


# --- User profile cleanup ---
Start-UserProfileCleanup -ComputerName $TargetMachines


# =====================================================================
#  Remediation
# =====================================================================

# --- Invoke software uninstall via registry uninstall strings---
Invoke-SoftwareUninstall -ComputerName $TargetMachines -SoftwareName "lightroom"


# --- Repair broken MSI uninstall entries ---
Repair-MSIUninstall -ComputerName $TargetMachines -SoftwareName "Google Chrome" -Remove


# --- Log4J remediation ---
Invoke-Log4JRemediation -ComputerName $TargetMachines


# --- Check / remove Tanium quarantine isolation policy ---
Test-TaniumQuarantine -ComputerName "PC01"
Test-TaniumQuarantine -ComputerName "PC01" -RemovePolicy
