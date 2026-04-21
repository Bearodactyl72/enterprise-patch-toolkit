# DOTS formatting comment

<#
    .SYNOPSIS
        Flushes and re-registers DNS on the local machine.

    .DESCRIPTION
        Runs ipconfig /flushdns, /registerdns, nbtstat -RR, ipconfig /release,
        ipconfig /renew, and netsh winsock reset to force the local machine to
        flush its DNS cache, re-register its A/PTR records, and renew its DHCP
        lease. Useful when your workstation has a stale DNS record and remote
        tools (WinRM, RDP, etc.) cannot reach you by name.

        Designed to be run locally on admin workstations. Requires admin
        elevation.

    .EXAMPLE
        .\Renew-DNS.ps1

    .NOTES
        Written by Skyler Werner
#>

[CmdletBinding()]
param()


# --- Elevation check ---
$isElevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isElevated) {
    Write-Host "ERROR: This script requires admin elevation." -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as Administrator'." -ForegroundColor Yellow
    return
}


# --- Build command list ---
$commands = [ordered]@{
    "Flush DNS cache"        = "ipconfig /flushdns"
    "Register DNS records"   = "ipconfig /registerdns"
    "Refresh NetBIOS names"  = "nbtstat -RR"
    "Release DHCP lease"     = "ipconfig /release"
    "Renew DHCP lease"       = "ipconfig /renew"
    "Reset Winsock"          = "netsh winsock reset"
}


# --- Execute ---
Write-Host ""
Write-Host "========================================" -ForegroundColor White
Write-Host " DNS Renewal -- $env:COMPUTERNAME"       -ForegroundColor White
Write-Host "========================================" -ForegroundColor White
Write-Host ""

$hasErrors = $false

foreach ($entry in $commands.GetEnumerator()) {
    $label = $entry.Key
    $cmd   = $entry.Value

    Write-Host "  $label... " -NoNewline

    $output = & cmd /c $cmd 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "OK" -ForegroundColor Green
    }
    else {
        Write-Host "FAILED (exit $LASTEXITCODE)" -ForegroundColor Red
        $hasErrors = $true
    }
}

Write-Host ""

if ($hasErrors) {
    Write-Host "  Some commands failed. Check output above." -ForegroundColor Yellow
}
else {
    Write-Host "  DNS renewal complete." -ForegroundColor Green
    Write-Host "  Records should update within a few minutes." -ForegroundColor DarkGray
    Write-Host "  Winsock reset requires a reboot to take full effect." -ForegroundColor Yellow
}

Write-Host ""
