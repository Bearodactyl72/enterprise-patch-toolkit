# DOTS formatting comment

# -----------------------------------------------------------------------------
# Environment configuration for the Remediation Script Library
# -----------------------------------------------------------------------------
# This file declares environment-specific values (domains, file shares,
# trusted runner hosts, org tags) so that the scripts themselves stay
# generic.
#
# Usage:
#   1. Copy this file to Config\Environment.psd1 (same directory).
#   2. Fill in values for your environment.
#   3. Leave Environment.psd1 out of source control (it is gitignored).
#
# Scripts load this via Import-RSLEnvironment in the Modules\RSL-Environment
# module. If Environment.psd1 is missing, the loader falls back to this
# example file so the repo is still runnable for a smoke test.
# -----------------------------------------------------------------------------

@{
    # Each entry is matched against $env:USERDNSDOMAIN at runtime to pick
    # the active profile. Covers the dual-network (primary/secondary)
    # pattern common in enterprise / segregated-network environments.
    # If $env:USERDNSDOMAIN matches nothing here, scripts treat the host
    # as a workgroup machine and skip domain prefix/suffix stripping.
    Networks = @(
        @{
            Name             = 'Primary'
            DomainFqdn       = 'corp.example.com'
            DomainShort      = 'CORP'
            PatchShareUnc    = '\\files.corp.example.com\WorkstationUtility\NetworkA'
            PatchLogShareUnc = '\\logs.corp.example.com\VMT\Metrics\PatchLogs'
        }
        @{
            Name             = 'Secondary'
            DomainFqdn       = 'corp.example.net'
            DomainShort      = 'CORPN'
            PatchShareUnc    = '\\files.corp.example.net\WorkstationUtility\NetworkB'
            PatchLogShareUnc = '\\logs.corp.example.net\VMT\Metrics\PatchLogs'
        }
    )

    # Drive letter Setup.ps1 maps to the active network's PatchShareUnc.
    MappedDriveLetter     = 'M'

    # Subpath under the mapped drive used as an "is the share healthy"
    # probe and as the base for CentralMainSwitchPath and CentralPSToolsPath.
    ShareAnchorPath       = 'Share\VMT'

    # Relative path under ShareAnchorPath to the authoritative copy of
    # Main-Switch.ps1. Setup.ps1 syncs the local copy against this when
    # the share is reachable.
    CentralMainSwitchPath = 'Scripts\Main-Switch\Main-Switch.ps1'

    # Relative path under ShareAnchorPath to PSTools (PsExec et al).
    # Copied to the operator's desktop by Setup.ps1 on trusted runners.
    CentralPSToolsPath    = 'Scripts\PSTools'

    # Computer-name regex patterns for hosts allowed to pull PSTools and
    # other centrally-hosted binaries onto the desktop. Prevents random
    # laptops from pulling sensitive tooling.
    TrustedRunnerHosts    = @(
        'ADMINBOX\d+'
        'PATCHRUNNER\d+'
    )

    # Organization tag used by CMDB / endpoint-agent integrations
    # (e.g. Tanium question filters) to identify machines belonging to
    # this org. Generic replacement for agency-specific component tags.
    OrgComponentTag       = 'Corp_Component:ACME'

    # Label for the source-control host, referenced only in Setup.ps1's
    # Unblock-File step after a fresh clone.
    SourceControlName     = 'Git'
}
