# DOTS formatting comment

<#
    .SYNOPSIS
        Placeholder -- WinRM fix script pending DNS team resolution.

    .DESCRIPTION
        --- INVESTIGATION FINDINGS (March 2026) ---

        Machines exhibiting WinRM error 0x80090322 (Kerberos authentication
        failure) were investigated using Test-RemoteAccess.ps1. All tested
        machines (6/6) turned out to have the SAME root cause:

          Stale DNS A records pointing hostnames to wrong IP addresses.

        The machines are alive and authenticating to AD (LastLogonDate within
        2 weeks), but their DNS dynamic update requests are being REFUSED by
        the DNS server (Event IDs 8012 and 8018). When we connect by hostname,
        DNS sends us to a completely different machine that now occupies the
        old IP -- Kerberos fails because the SPN does not match.

        Every remote workaround was tested and eliminated:
          - PsExec/WMI/SMB by IP --> reaches the wrong machine (stale IP)
          - WinRM by IP --> blocked by TrustedHosts GPO
          - Tanium --> client not checking in, cached data also stale
          - DHCP lookup --> no read permissions on DHCP servers
          - DC event logs --> no access

        The fix must come from the DNS server side:
          1. Delete stale A records so machines can re-register
          2. Fix dynamic update permissions (record ACL ownership)
          3. Enable DNS scavenging to prevent recurrence

        Full report: Docs\DNS_Stale_Record_Investigation.docx
        Diagnostic tool: Scripts\Utility\Discovery\Test-RemoteAccess.ps1

        If the DNS team resolves the server-side issue, this script becomes
        unnecessary -- machines will self-heal on next DHCP renewal. If they
        cannot or will not fix it, this file is reserved for whatever
        workaround we devise next.

    .NOTES
        Written by Skyler Werner

        Original script enabled the WinRM HTTP-In firewall rule via WMI
        (MSFT_NetFirewallRule). That functionality was removed because the
        firewall is not the issue -- the WinRM service is running fine on
        affected machines. The error is a DNS/Kerberos problem.
#>

Write-Host "This script is not yet implemented." -ForegroundColor Yellow
Write-Host "See: Docs\DNS_Stale_Record_Investigation.md" -ForegroundColor DarkGray
Write-Host "See: Scripts\Utility\Discovery\Test-RemoteAccess.ps1" -ForegroundColor DarkGray
