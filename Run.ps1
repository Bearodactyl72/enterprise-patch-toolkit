# Run.ps1 -- Quick-launch scratch pad for script invocations
# Open in ISE, F5 to load functions, edit args, select the lines you need, F8 to run selection.
#
# $TargetMachines is loaded automatically by the TargetMachines module.
# If you update the list file mid-session, run:  Import-TargetMachines
# To paste a list interactively, run:            Enter-TargetMachines

$ScriptRoot = "$env:USERPROFILE\Desktop\Remediation-Script-Library"
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
