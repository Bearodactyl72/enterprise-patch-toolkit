# DOTS formatting comment

<#
    .SYNOPSIS
        Shared helpers used by every Patch GUI entry point.
    .DESCRIPTION
        Any helper that is pure logic (no XAML, no named-control
        dependencies) lives here so the canonical GUI, the gallery,
        and theme overrides all stay in sync automatically.

        Contract for callers:
          - Dot-source this file early, before invoking any helper.
          - Helpers rely on $PSScriptRoot of THIS file to locate
            Main-Switch.ps1, which means callers can live at any
            depth under Scripts/ without adjusting path math.
          - Helpers also rely on $script:PrefsDir / $script:PrefsPath,
            which this file sets when dot-sourced.

        Callers currently dot-sourcing this file:
          - Invoke-PatchGUI.ps1
          - Invoke-PatchGUI-Gallery.ps1
          - ThemeOverrides/CyberCommand.ps1

        Written by Skyler Werner
        Date: 2026/04/23
        Version 1.0.0
#>


# ============================================================================
#  Shared paths
# ============================================================================
# Dot-sourcing runs this block in the caller's script scope, so each
# caller ends up with these two script-scope variables populated without
# needing to duplicate the lines themselves.

$script:PrefsDir  = Join-Path $env:APPDATA 'Patching'
$script:PrefsPath = Join-Path $script:PrefsDir 'preferences.json'


# ============================================================================
#  Preferences I/O
# ============================================================================

function Get-PatchPreferences {
    if (-not (Test-Path $script:PrefsPath)) { return @{} }
    try {
        $raw = Get-Content -Raw -Path $script:PrefsPath -Encoding UTF8
        $obj = $raw | ConvertFrom-Json
        $h = @{}
        foreach ($p in $obj.PSObject.Properties) { $h[$p.Name] = $p.Value }
        return $h
    } catch {
        return @{}
    }
}

function Set-PatchPreferences {
    param([hashtable]$Preferences)
    if (-not (Test-Path $script:PrefsDir)) {
        New-Item -ItemType Directory -Path $script:PrefsDir -Force | Out-Null
    }
    $Preferences | ConvertTo-Json | Set-Content -Path $script:PrefsPath -Encoding UTF8
}


# ============================================================================
#  Main-Switch.ps1 discovery
# ============================================================================
# This file lives at Scripts/Patching/GUI/ regardless of which caller
# dot-sourced it, so $PSScriptRoot inside these functions resolves to
# GUI/ and Main-Switch.ps1 is always ..\..\Main-Switch.ps1 from here.

function Get-MainSwitchNames {
    [CmdletBinding()]
    param()

    # Skip candidates whose base path is null so Join-Path does not throw
    # when the script is launched outside a profile-loaded session
    # ($scriptPath is set by the admin's profile in normal use).
    $candidates = @()
    if ($PSScriptRoot) { $candidates += (Join-Path $PSScriptRoot '..\..\Main-Switch.ps1') }
    if ($scriptPath)   { $candidates += (Join-Path $scriptPath 'Main-Switch.ps1') }

    $mainSwitchPath = $null
    foreach ($c in $candidates) {
        if (Test-Path $c) {
            $mainSwitchPath = (Resolve-Path $c).Path
            break
        }
    }

    if (-not $mainSwitchPath) {
        Write-Warning "Could not find Main-Switch.ps1"
        return @()
    }

    # Parse the switch case names via regex
    $lines = Get-Content -Path $mainSwitchPath -Encoding Default
    $names = @()
    foreach ($line in $lines) {
        if ($line -match '^\s*"([^"]+)"\s*\{') {
            $names += $Matches[1]
        }
    }

    return ($names | Sort-Object)
}

function Get-MainSwitchListPaths {
    <#
        .SYNOPSIS
            Returns a hashtable mapping each software name to its listPath.
        .DESCRIPTION
            Parses Main-Switch.ps1 line by line to find each case header
            ("Name" {) and the first $listPath assignment inside it.
            Expands $listPathRoot and $env: variables.

            Avoids dot-sourcing Main-Switch entirely -- dot-sourcing fails
            in the GUI runspace because Main-Switch calls Get-Command
            against per-software patch scripts that do not resolve in a
            fresh context. A silent catch would swallow that error and
            produce an empty map, so the Machine field never populated.
    #>
    [CmdletBinding()]
    param()

    $candidates = @()
    if ($PSScriptRoot) { $candidates += (Join-Path $PSScriptRoot '..\..\Main-Switch.ps1') }
    if ($scriptPath)   { $candidates += (Join-Path $scriptPath 'Main-Switch.ps1') }

    $mainSwitchPath = $null
    foreach ($c in $candidates) {
        if (Test-Path $c) {
            $mainSwitchPath = (Resolve-Path $c).Path
            break
        }
    }

    if (-not $mainSwitchPath) { return @{} }

    # Context variable referenced by listPath values in Main-Switch.
    # ExpandString below resolves $listPathRoot against this scope.
    $listPathRoot = "$env:USERPROFILE\Desktop\Lists"

    $map         = @{}
    $currentCase = $null

    foreach ($line in (Get-Content -Path $mainSwitchPath -Encoding Default)) {
        # Case header: indented "Name" followed by opening brace
        if ($line -match '^\s*"([^"]+)"\s*\{') {
            $currentCase = $Matches[1]
            continue
        }

        # First $listPath assignment inside an active case
        if ($currentCase -and
            $line -match '^\s*\$listPath\s*=\s*"([^"]+)"') {
            $rawPath = $Matches[1]
            try {
                $expanded = $ExecutionContext.InvokeCommand.ExpandString($rawPath)
                $map[$currentCase] = $expanded
            }
            catch { }
            $currentCase = $null   # done; ignore further assignments in this case
        }
    }

    return $map
}


# ============================================================================
#  DryRun: mock results generator
# ============================================================================

function New-MockPatchResults {
    [CmdletBinding()]
    param(
        [string]$SoftwareName,
        [ValidateSet('Patch','Version')]
        [string]$Mode = 'Patch',
        [int]$Count = 15,
        [string[]]$ComputerName
    )

    # If caller supplied an explicit hostname list, drive the row count
    # from it and render those names verbatim. Otherwise fall back to
    # synthesized PC001-style names (keeps behaviour for callers that
    # just want N rows of plausible-looking data).
    $useProvidedNames = ($null -ne $ComputerName) -and (@($ComputerName).Count -gt 0)
    if ($useProvidedNames) { $Count = @($ComputerName).Count }

    $isVersionMode = ($Mode -eq 'Version')
    $rng       = New-Object System.Random
    $date      = Get-Date -Format 'yyyy/MM/dd HH:mm'
    $user      = $env:USERNAME
    $oldVers   = @('119.0.6045.123', '120.0.6099.71', '121.0.6167.85', '122.0.6261.57')
    $newVer    = '126.0.6478.127'
    $prefixes  = @('PC', 'WS', 'PC', 'PC', 'DT', 'LT')

    # Auto-detection mix: mostly Online, with occasional Offline and rare
    # Isolated (machine reached via WinRM-port fallback after ICMP failure).
    $statuses = @('Online','Online','Online','Online','Online','Online','Online','Online','Offline','Isolated')

    $results = @()
    for ($i = 1; $i -le $Count; $i++) {
        $status  = $statuses[$rng.Next($statuses.Count)]
        $machine = if ($useProvidedNames) {
            $ComputerName[$i - 1]
        } else {
            $prefix = $prefixes[$rng.Next($prefixes.Count)]
            "$prefix$($i.ToString('D3'))"
        }
        $octA    = $rng.Next(10, 11)
        $octB    = $rng.Next(1, 5)
        $octC    = $rng.Next(1, 255)

        if ($status -eq 'Offline') {
            $results += [PSCustomObject]@{
                IPAddress    = $null
                ComputerName = $machine
                Status       = 'Offline'
                SoftwareName = $SoftwareName
                Version      = $null
                Compliant    = $null
                NewVersion   = $null
                ExitCode     = $null
                Comment      = 'Ping failed'
                AdminName    = $user
                Date         = $date
            }
        }
        else {
            $oldVer = $oldVers[$rng.Next($oldVers.Count)]

            # Isolated machines have null IPAddress (ping didn't return one).
            $ip = if ($status -eq 'Isolated') { $null } else { "$octA.$octB.$octC.$i" }

            if ($isVersionMode) {
                # Version mode: real Invoke-Version only produces version +
                # compliant state. Leave patch-specific fields null so the
                # hidden columns render empty when the user flips the slider.
                $results += [PSCustomObject]@{
                    IPAddress    = $ip
                    ComputerName = $machine
                    Status       = $status
                    SoftwareName = $SoftwareName
                    Version      = $oldVer
                    Compliant    = $false
                    NewVersion   = $null
                    ExitCode     = $null
                    Comment      = $null
                    AdminName    = $user
                    Date         = $date
                }
                continue
            }

            $exitCode = @(0, 0, 0, 0, 0, 3010, 1603, 1618)[$rng.Next(8)]
            $gotNew   = ($exitCode -eq 0 -or $exitCode -eq 3010)
            $comment  = switch ($exitCode) {
                0       { 'Success' }
                3010    { 'Reboot required' }
                1603    { 'Fatal error during installation' }
                1618    { 'Another install in progress' }
                default { '' }
            }

            $results += [PSCustomObject]@{
                IPAddress    = $ip
                ComputerName = $machine
                Status       = $status
                SoftwareName = $SoftwareName
                Version      = $oldVer
                Compliant    = $false
                # Match real Invoke-Patch: when the install doesn't
                # change the detected version (attempt failed), the
                # cmdlet rewrites NewVersion to "No Change". This also
                # bubbles failed rows to the top of each Version block
                # under the NewVersion-DESC sort.
                NewVersion   = if ($gotNew) { $newVer } else { 'No Change' }
                ExitCode     = $exitCode
                Comment      = $comment
                AdminName    = $user
                Date         = $date
            }
        }
    }

    return $results
}


# ============================================================================
#  Display-row flattener
# ============================================================================

function ConvertTo-DisplayRow {
    <#
        .SYNOPSIS
            Returns a clone of the input row with array-valued display
            properties flattened to comma-separated strings.
        .DESCRIPTION
            The WPF DataGrid renders [object[]] as "System.Object[]". A machine
            with multiple installs of the same software (e.g. user + system
            Chrome) produces a Version array, and any upstream that hasn't
            already applied Add-Delimiter will surface the array to the grid.
            This helper defensively flattens the common offenders so the GUI
            is robust regardless of which version of the backend is running.
    #>
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [psobject]$Row
    )
    process {
        if ($null -eq $Row) { return }
        $clone = $Row.PSObject.Copy()
        foreach ($prop in @('Version', 'NewVersion', 'ExitCode', 'Comment')) {
            if ($null -eq $clone.PSObject.Properties[$prop]) { continue }
            $val = $clone.$prop
            # Flatten arrays / collections to a comma-separated string.
            # Treat single strings and scalars as-is.
            if ($val -is [System.Collections.IEnumerable] -and $val -isnot [string]) {
                $clone.$prop = ($val | ForEach-Object { "$_" }) -join ', '
            }
        }
        $clone
    }
}
