# DOTS formatting comment

<#
    .SYNOPSIS
        Content-aware merge tool for Main-Switch.ps1 across multiple admins.
    .DESCRIPTION
        Parses Main-Switch.ps1 into Header, individual switch-case blocks, and Footer.
        Compares two versions (local and central) using three-way merge with a stored
        base file to detect which side changed. Auto-merges non-conflicting changes
        and prompts for conflict resolution when both sides modified the same case.

        Written by Skyler Werner
        Date: 2026/03/18
        Version 1.0.0
#>


# ============================================================================
#  Private: Resolve default paths from Paths.txt
# ============================================================================

function Get-DefaultScriptPath {
    [CmdletBinding()]
    param()

    $pathsFile = Join-Path $env:APPDATA 'Patching\Paths.txt'
    if (-not (Test-Path $pathsFile)) { return $null }

    $lines = Get-Content -Path $pathsFile -Encoding Default
    foreach ($line in $lines) {
        if ($line -match '^\s*ScriptPath\s*:\s*(.+)$') {
            $resolved = $Matches[1].Trim()
            if (Test-Path $resolved) { return $resolved }
        }
    }
    return $null
}


# ============================================================================
#  Private: Parser -- ConvertTo-ParsedMainSwitch
# ============================================================================

function ConvertTo-ParsedMainSwitch {
    <#
        .SYNOPSIS
            Parses a Main-Switch.ps1 file into Header, Cases (ordered dict), and Footer.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "File not found: $Path"
    }

    # Read all lines preserving encoding
    $allLines = Get-Content -Path $Path -Encoding Default

    # --- Find header end: the "switch ($targetSoftware) {" line ---
    $switchOpenIndex = -1
    for ($i = 0; $i -lt $allLines.Count; $i++) {
        if ($allLines[$i] -match '^\s*switch\s*\(\$targetSoftware\)\s*\{') {
            $switchOpenIndex = $i
            break
        }
    }
    if ($switchOpenIndex -lt 0) {
        throw "Could not find 'switch (`$targetSoftware) {' in $Path"
    }

    # --- Find footer start: the "} # End switch" line ---
    $switchCloseIndex = -1
    for ($i = $allLines.Count - 1; $i -ge 0; $i--) {
        if ($allLines[$i] -match '^\s*\}\s*#\s*End\s+switch') {
            $switchCloseIndex = $i
            break
        }
    }
    if ($switchCloseIndex -lt 0) {
        throw "Could not find '} # End switch' in $Path"
    }

    # Header = lines 0 through switchOpenIndex (inclusive)
    $header = ($allLines[0..$switchOpenIndex]) -join "`r`n"

    # Footer = lines switchCloseIndex through end
    $footer = ($allLines[$switchCloseIndex..($allLines.Count - 1)]) -join "`r`n"

    # Switch body = lines between header and footer
    $bodyStart = $switchOpenIndex + 1
    $bodyEnd   = $switchCloseIndex - 1

    if ($bodyEnd -lt $bodyStart) {
        # Empty switch body
        $cases = [ordered]@{}
        return [PSCustomObject]@{
            Header = $header
            Cases  = $cases
            Footer = $footer
        }
    }

    $bodyLines = $allLines[$bodyStart..$bodyEnd]

    # --- Parse body into case blocks ---
    # Find all case-start line indices within the body
    $casePattern = '^\s*"([^"]+)"\s*\{'
    $caseStarts = [System.Collections.ArrayList]::new()

    for ($i = 0; $i -lt $bodyLines.Count; $i++) {
        if ($bodyLines[$i] -match $casePattern) {
            [void]$caseStarts.Add(@{
                Index = $i
                Key   = $Matches[1]
            })
        }
    }

    $cases = [ordered]@{}

    if ($caseStarts.Count -eq 0) {
        # No cases found -- store entire body as a special entry
        $cases['__BODY__'] = ($bodyLines) -join "`r`n"
    }
    else {
        # Text before the first case (region markers, blank lines)
        if ($caseStarts[0].Index -gt 0) {
            $prefixLines = $bodyLines[0..($caseStarts[0].Index - 1)]
        }
        else {
            $prefixLines = @()
        }

        for ($c = 0; $c -lt $caseStarts.Count; $c++) {
            $caseKey   = $caseStarts[$c].Key
            $startIdx  = $caseStarts[$c].Index

            # End index is the line before the next case's prefix starts,
            # or end of body for the last case
            if ($c -lt $caseStarts.Count - 1) {
                $endIdx = $caseStarts[$c + 1].Index - 1
            }
            else {
                $endIdx = $bodyLines.Count - 1
            }

            # The case block includes its prefix (comments/regions before it)
            # For the first case, prefix is the text before it
            # For subsequent cases, prefix is text after previous case's closing }
            $caseText = ($prefixLines + $bodyLines[$startIdx..$endIdx]) -join "`r`n"

            # Handle duplicate keys (unlikely but safe)
            $finalKey = $caseKey
            $suffix = 1
            while ($cases.Contains($finalKey)) {
                $suffix++
                $finalKey = "${caseKey}_DUP${suffix}"
                Write-Warning "Duplicate case key '$caseKey' found -- renamed to '$finalKey'"
            }

            $cases[$finalKey] = $caseText

            # Calculate prefix for the NEXT case: lines between this case's
            # content end and the next case's start line
            if ($c -lt $caseStarts.Count - 1) {
                $nextStartIdx = $caseStarts[$c + 1].Index
                # Find where the current case's content actually ends
                # We need to find lines between endIdx+1 and nextStartIdx-1
                # But endIdx is already nextStartIdx-1, so prefix is empty
                # Actually, let's re-examine: endIdx = nextStartIdx - 1
                # So the NEXT case's prefix is the gap lines between cases
                # Since we set endIdx = nextStart - 1, those gap lines are
                # included in the current case's block. The next case has no prefix.
                $prefixLines = @()
            }
        }
    }

    return [PSCustomObject]@{
        Header = $header
        Cases  = $cases
        Footer = $footer
    }
}


# ============================================================================
#  Private: Reconstructor -- ConvertFrom-ParsedMainSwitch
# ============================================================================

function ConvertFrom-ParsedMainSwitch {
    <#
        .SYNOPSIS
            Reassembles a parsed Main-Switch object back into file content.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Parsed
    )

    $parts = [System.Collections.ArrayList]::new()
    [void]$parts.Add($Parsed.Header)

    foreach ($key in $Parsed.Cases.Keys) {
        [void]$parts.Add($Parsed.Cases[$key])
    }

    [void]$parts.Add($Parsed.Footer)

    return (($parts -join "`r`n") + "`r`n")
}


# ============================================================================
#  Private: Normalize text for comparison
# ============================================================================

function Get-NormalizedText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Text
    )

    # Normalize line endings and trim trailing whitespace per line
    $lines = $Text -split "`r?`n"
    $trimmed = foreach ($line in $lines) { $line.TrimEnd() }
    return ($trimmed -join "`n")
}


# ============================================================================
#  Private: Three-way comparison engine
# ============================================================================

function Compare-ParsedFiles {
    <#
        .SYNOPSIS
            Compares local and central parsed files using optional base for 3-way merge.
        .OUTPUTS
            PSCustomObject with HeaderStatus, FooterStatus, CaseResults (ordered dict)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Local,

        [Parameter(Mandatory)]
        [PSCustomObject]$Central,

        [Parameter()]
        [PSCustomObject]$Base
    )

    $hasBase = $null -ne $Base

    # --- Helper: determine status for a section ---
    # Returns: Identical, LocalChanged, CentralChanged, Conflict
    $getStatus = {
        param($localText, $centralText, $baseText, $hasBase)

        $localNorm   = Get-NormalizedText -Text $localText
        $centralNorm = Get-NormalizedText -Text $centralText

        if ($localNorm -eq $centralNorm) { return 'Identical' }

        if ($hasBase -and $null -ne $baseText) {
            $baseNorm = Get-NormalizedText -Text $baseText
            $localChanged   = $localNorm   -ne $baseNorm
            $centralChanged = $centralNorm -ne $baseNorm
            if ($localChanged -and -not $centralChanged) { return 'LocalChanged' }
            if (-not $localChanged -and $centralChanged) { return 'CentralChanged' }
            return 'Conflict'
        }

        # No base -- any difference is a conflict
        return 'Conflict'
    }

    # Header comparison
    $headerStatus = & $getStatus $Local.Header $Central.Header $(if ($hasBase) { $Base.Header } else { $null }) $hasBase

    # Footer comparison
    $footerStatus = & $getStatus $Local.Footer $Central.Footer $(if ($hasBase) { $Base.Footer } else { $null }) $hasBase

    # Case comparison
    $caseResults = [ordered]@{}
    $allKeys = [System.Collections.ArrayList]::new()

    foreach ($key in $Local.Cases.Keys) {
        if (-not $allKeys.Contains($key)) { [void]$allKeys.Add($key) }
    }
    foreach ($key in $Central.Cases.Keys) {
        if (-not $allKeys.Contains($key)) { [void]$allKeys.Add($key) }
    }

    foreach ($key in $allKeys) {
        $inLocal   = $Local.Cases.Contains($key)
        $inCentral = $Central.Cases.Contains($key)
        $inBase    = $hasBase -and $Base.Cases.Contains($key)

        if ($inLocal -and $inCentral) {
            # Both have it
            $baseText = if ($inBase) { $Base.Cases[$key] } else { $null }
            $status = & $getStatus $Local.Cases[$key] $Central.Cases[$key] $baseText $hasBase
        }
        elseif ($inLocal -and -not $inCentral) {
            if ($inBase) {
                # Was in base, removed from central -- central deleted it
                $status = 'CentralDeleted'
            }
            else {
                $status = 'LocalOnly'
            }
        }
        elseif (-not $inLocal -and $inCentral) {
            if ($inBase) {
                # Was in base, removed from local -- local deleted it
                $status = 'LocalDeleted'
            }
            else {
                $status = 'CentralOnly'
            }
        }
        else {
            continue # Should not happen
        }

        $caseResults[$key] = [PSCustomObject]@{
            Key         = $key
            Status      = $status
            LocalText   = if ($inLocal)   { $Local.Cases[$key] }   else { $null }
            CentralText = if ($inCentral) { $Central.Cases[$key] } else { $null }
        }
    }

    return [PSCustomObject]@{
        HeaderStatus = $headerStatus
        FooterStatus = $footerStatus
        CaseResults  = $caseResults
    }
}


# ============================================================================
#  Private: Interactive conflict resolution
# ============================================================================

function Resolve-MergeConflicts {
    <#
        .SYNOPSIS
            Prompts admin to resolve each conflict interactively.
        .OUTPUTS
            Ordered dictionary of resolved case texts keyed by case name.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.Specialized.OrderedDictionary]$Conflicts
    )

    $resolved = [ordered]@{}

    foreach ($key in $Conflicts.Keys) {
        $conflict = $Conflicts[$key]

        Write-Host ''
        Write-Host ('=' * 70) -ForegroundColor Yellow
        Write-Host "  CONFLICT: `"$key`"" -ForegroundColor Yellow
        Write-Host ('=' * 70) -ForegroundColor Yellow
        Write-Host ''

        # Show side-by-side diff of the lines that differ
        $localLines   = ($conflict.LocalText   -split "`r?`n")
        $centralLines = ($conflict.CentralText -split "`r?`n")

        # Find differing lines
        $maxLines = [Math]::Max($localLines.Count, $centralLines.Count)
        $hasDiffs = $false

        for ($i = 0; $i -lt $maxLines; $i++) {
            $lLine = if ($i -lt $localLines.Count)   { $localLines[$i] }   else { '' }
            $cLine = if ($i -lt $centralLines.Count) { $centralLines[$i] } else { '' }

            if ($lLine.TrimEnd() -ne $cLine.TrimEnd()) {
                if (-not $hasDiffs) {
                    Write-Host '  Differences:' -ForegroundColor White
                    $hasDiffs = $true
                }
                Write-Host "    [LOCAL]   $lLine" -ForegroundColor Green
                Write-Host "    [CENTRAL] $cLine" -ForegroundColor Cyan
                Write-Host ''
            }
        }

        if (-not $hasDiffs) {
            Write-Host '  (Whitespace-only differences)' -ForegroundColor DarkGray
        }

        # Prompt for resolution
        $choice = ''
        while ($choice -notin @('L','C','S')) {
            Write-Host '  Choose: [L]ocal  [C]entral  [S]kip (keep local, resolve later)' -ForegroundColor Yellow -NoNewline
            $choice = (Read-Host ' ').Trim().ToUpper()
            if ($choice -eq '') { $choice = 'S' }
        }

        switch ($choice) {
            'L' {
                $resolved[$key] = $conflict.LocalText
                Write-Host "  -> Keeping LOCAL version of `"$key`"" -ForegroundColor Green
            }
            'C' {
                $resolved[$key] = $conflict.CentralText
                Write-Host "  -> Taking CENTRAL version of `"$key`"" -ForegroundColor Cyan
            }
            'S' {
                $resolved[$key] = $conflict.LocalText
                Write-Host "  -> Skipped -- keeping LOCAL version (resolve manually later)" -ForegroundColor DarkYellow
            }
        }
    }

    return $resolved
}


# ============================================================================
#  Private: Merge engine
# ============================================================================

function Merge-ParsedFiles {
    <#
        .SYNOPSIS
            Merges two parsed files based on comparison results.
        .DESCRIPTION
            Returns a merged parsed object and summary statistics.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Local,

        [Parameter(Mandatory)]
        [PSCustomObject]$Central,

        [Parameter(Mandatory)]
        [PSCustomObject]$Comparison,

        [Parameter()]
        [string]$MergeDirection = 'Pull'  # 'Pull' = central into local, 'Push' = local into central
    )

    # Start with a copy of the target side
    if ($MergeDirection -eq 'Pull') {
        $mergedHeader = $Local.Header
        $mergedFooter = $Local.Footer
        $mergedCases  = [ordered]@{}
        foreach ($key in $Local.Cases.Keys) {
            $mergedCases[$key] = $Local.Cases[$key]
        }
    }
    else {
        $mergedHeader = $Central.Header
        $mergedFooter = $Central.Footer
        $mergedCases  = [ordered]@{}
        foreach ($key in $Central.Cases.Keys) {
            $mergedCases[$key] = $Central.Cases[$key]
        }
    }

    # Stats
    $stats = @{
        Identical       = 0
        AutoMerged      = 0
        NewEntries      = 0
        Conflicts       = 0
        Skipped         = 0
        Deleted         = 0
    }

    # Collect conflicts for interactive resolution
    $conflicts = [ordered]@{}
    # Track new entries to append
    $newEntries = [ordered]@{}

    foreach ($key in $Comparison.CaseResults.Keys) {
        $result = $Comparison.CaseResults[$key]

        switch ($result.Status) {
            'Identical' {
                $stats.Identical++
            }

            'LocalChanged' {
                if ($MergeDirection -eq 'Pull') {
                    # Pulling: local changed, keep local (already there)
                    $stats.Identical++
                }
                else {
                    # Pushing: local changed, update central
                    $mergedCases[$key] = $result.LocalText
                    $stats.AutoMerged++
                }
            }

            'CentralChanged' {
                if ($MergeDirection -eq 'Pull') {
                    # Pulling: central changed, take central
                    $mergedCases[$key] = $result.CentralText
                    $stats.AutoMerged++
                }
                else {
                    # Pushing: central changed, keep central (already there)
                    $stats.Identical++
                }
            }

            'LocalOnly' {
                if ($MergeDirection -eq 'Pull') {
                    # Pulling: local has a new entry, keep it
                    $stats.Identical++
                }
                else {
                    # Pushing: local has a new entry, add to central
                    $newEntries[$key] = $result.LocalText
                    $stats.NewEntries++
                }
            }

            'CentralOnly' {
                if ($MergeDirection -eq 'Pull') {
                    # Pulling: central has a new entry, add to local
                    $newEntries[$key] = $result.CentralText
                    $stats.NewEntries++
                }
                else {
                    # Pushing: central has an entry we dont have, keep it
                    $stats.Identical++
                }
            }

            'CentralDeleted' {
                if ($MergeDirection -eq 'Pull') {
                    # Central removed it -- remove from local
                    $mergedCases.Remove($key)
                    $stats.Deleted++
                }
                else {
                    $stats.Identical++
                }
            }

            'LocalDeleted' {
                if ($MergeDirection -eq 'Pull') {
                    $stats.Identical++
                }
                else {
                    # Local removed it -- remove from central
                    $mergedCases.Remove($key)
                    $stats.Deleted++
                }
            }

            'Conflict' {
                $conflicts[$key] = $result
                $stats.Conflicts++
            }
        }
    }

    # --- Handle header/footer ---
    if ($MergeDirection -eq 'Pull') {
        if ($Comparison.HeaderStatus -eq 'CentralChanged') {
            $mergedHeader = $Central.Header
        }
        elseif ($Comparison.HeaderStatus -eq 'Conflict') {
            Write-Host ''
            Write-Host '  HEADER CONFLICT -- the header section differs on both sides.' -ForegroundColor Yellow
            Write-Host '  Review manually after merge completes.' -ForegroundColor Yellow
        }

        if ($Comparison.FooterStatus -eq 'CentralChanged') {
            $mergedFooter = $Central.Footer
        }
        elseif ($Comparison.FooterStatus -eq 'Conflict') {
            Write-Host ''
            Write-Host '  FOOTER CONFLICT -- the footer/parameters section differs on both sides.' -ForegroundColor Yellow
            Write-Host '  Review manually after merge completes.' -ForegroundColor Yellow
        }
    }
    else {
        if ($Comparison.HeaderStatus -eq 'LocalChanged') {
            $mergedHeader = $Local.Header
        }
        elseif ($Comparison.HeaderStatus -eq 'Conflict') {
            Write-Host ''
            Write-Host '  HEADER CONFLICT -- the header section differs on both sides.' -ForegroundColor Yellow
            Write-Host '  Review manually after merge completes.' -ForegroundColor Yellow
        }

        if ($Comparison.FooterStatus -eq 'LocalChanged') {
            $mergedFooter = $Local.Footer
        }
        elseif ($Comparison.FooterStatus -eq 'Conflict') {
            Write-Host ''
            Write-Host '  FOOTER CONFLICT -- the footer/parameters section differs on both sides.' -ForegroundColor Yellow
            Write-Host '  Review manually after merge completes.' -ForegroundColor Yellow
        }
    }

    # --- Resolve conflicts interactively ---
    if ($conflicts.Count -gt 0) {
        $resolved = Resolve-MergeConflicts -Conflicts $conflicts
        foreach ($key in $resolved.Keys) {
            $mergedCases[$key] = $resolved[$key]
        }
    }

    # --- Append new entries at end of cases ---
    if ($newEntries.Count -gt 0) {
        $sourceLabel = if ($MergeDirection -eq 'Pull') { 'central' } else { 'local' }

        foreach ($key in $newEntries.Keys) {
            $mergedCases[$key] = $newEntries[$key]
        }
    }

    $merged = [PSCustomObject]@{
        Header = $mergedHeader
        Cases  = $mergedCases
        Footer = $mergedFooter
    }

    return [PSCustomObject]@{
        Merged = $merged
        Stats  = $stats
    }
}


# ============================================================================
#  Private: Write file with ASCII validation
# ============================================================================

function Write-MainSwitchFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Content
    )

    # Validate ASCII-only
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Content)
    foreach ($b in $bytes) {
        if ($b -gt 127) {
            Write-Warning "Non-ASCII byte detected (0x$($b.ToString('X2'))). File may contain Unicode characters."
            Write-Warning "Proceeding with write, but verify the file in PowerShell ISE."
            break
        }
    }

    # Write with Default encoding (Windows-1252 on PS 5.1)
    Set-Content -Path $Path -Value $Content -Encoding Default -NoNewline
}


# ============================================================================
#  Public: Compare-MainSwitch (read-only diff)
# ============================================================================

function Compare-MainSwitch {
    <#
        .SYNOPSIS
            Compares local and central Main-Switch.ps1 without making changes.
        .DESCRIPTION
            Parses both files, compares them case-by-case, and displays a color-coded
            report showing what would change if you ran Receive-MainSwitch.
        .PARAMETER LocalPath
            Path to your local Main-Switch.ps1. Defaults to the path from Paths.txt.
        .PARAMETER CentralPath
            Path to the shared central Main-Switch.ps1. Defaults to M:\Share\VMT\Scripts\Main-Switch.ps1.
        .EXAMPLE
            Compare-MainSwitch
        .EXAMPLE
            Compare-MainSwitch -LocalPath "C:\MyScripts\Main-Switch.ps1" -CentralPath "\\Server\Share\Main-Switch.ps1"
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$LocalPath,

        [Parameter()]
        [string]$CentralPath = 'M:\Share\VMT\Scripts\Main-Switch\Main-Switch.ps1'
    )

    # Resolve local path default
    if (-not $LocalPath) {
        $scriptDir = Get-DefaultScriptPath
        if ($scriptDir) {
            $LocalPath = Join-Path $scriptDir 'Main-Switch.ps1'
        }
        else {
            throw 'Could not resolve local script path. Provide -LocalPath or run Setup.ps1 first.'
        }
    }

    if (-not (Test-Path -LiteralPath $LocalPath)) {
        throw "Local file not found: $LocalPath"
    }
    if (-not (Test-Path -LiteralPath $CentralPath)) {
        throw "Central file not found: $CentralPath"
    }

    $local   = ConvertTo-ParsedMainSwitch -Path $LocalPath
    $central = ConvertTo-ParsedMainSwitch -Path $CentralPath

    # Check for base file
    $basePath = "$LocalPath.base"
    $base = if (Test-Path -LiteralPath $basePath) {
        ConvertTo-ParsedMainSwitch -Path $basePath
    } else { $null }

    $comparison = Compare-ParsedFiles -Local $local -Central $central -Base $base

    # --- Display results ---
    Write-Host ''
    Write-Host ('-' * 60) -ForegroundColor White
    Write-Host '  Main-Switch.ps1 Comparison Report' -ForegroundColor White
    Write-Host ('-' * 60) -ForegroundColor White

    if (-not $base) {
        Write-Host '  NOTE: No base file found. This is likely the first comparison.' -ForegroundColor DarkGray
        Write-Host '  All differences will show as conflicts until first sync.' -ForegroundColor DarkGray
    }
    Write-Host ''

    # Header/Footer
    if ($comparison.HeaderStatus -ne 'Identical') {
        $color = switch ($comparison.HeaderStatus) {
            'LocalChanged'   { 'Green' }
            'CentralChanged' { 'Cyan' }
            'Conflict'       { 'Yellow' }
        }
        Write-Host "  HEADER: $($comparison.HeaderStatus)" -ForegroundColor $color
    }
    if ($comparison.FooterStatus -ne 'Identical') {
        $color = switch ($comparison.FooterStatus) {
            'LocalChanged'   { 'Green' }
            'CentralChanged' { 'Cyan' }
            'Conflict'       { 'Yellow' }
        }
        Write-Host "  FOOTER: $($comparison.FooterStatus)" -ForegroundColor $color
    }

    # Cases
    $counts = @{ Identical = 0; LocalChanged = 0; CentralChanged = 0; Conflict = 0; LocalOnly = 0; CentralOnly = 0; CentralDeleted = 0; LocalDeleted = 0 }

    foreach ($key in $comparison.CaseResults.Keys) {
        $result = $comparison.CaseResults[$key]
        $counts[$result.Status]++

        if ($result.Status -eq 'Identical') { continue }

        $color = switch ($result.Status) {
            'LocalChanged'   { 'Green' }
            'CentralChanged' { 'Cyan' }
            'Conflict'       { 'Yellow' }
            'LocalOnly'      { 'Green' }
            'CentralOnly'    { 'Cyan' }
            'CentralDeleted' { 'Magenta' }
            'LocalDeleted'   { 'Magenta' }
        }

        $label = switch ($result.Status) {
            'LocalChanged'   { 'LOCAL changed' }
            'CentralChanged' { 'CENTRAL changed' }
            'Conflict'       { 'CONFLICT (both changed)' }
            'LocalOnly'      { 'NEW (local only)' }
            'CentralOnly'    { 'NEW (central only)' }
            'CentralDeleted' { 'DELETED from central' }
            'LocalDeleted'   { 'DELETED from local' }
        }

        Write-Host "  $($key.PadRight(30)) $label" -ForegroundColor $color
    }

    # Summary
    Write-Host ''
    Write-Host ('-' * 60) -ForegroundColor White
    Write-Host "  Identical:        $($counts.Identical)" -ForegroundColor DarkGray
    Write-Host "  Local changed:    $($counts.LocalChanged)" -ForegroundColor Green
    Write-Host "  Central changed:  $($counts.CentralChanged)" -ForegroundColor Cyan
    Write-Host "  Conflicts:        $($counts.Conflict)" -ForegroundColor Yellow
    Write-Host "  New (local only): $($counts.LocalOnly)" -ForegroundColor Green
    Write-Host "  New (central):    $($counts.CentralOnly)" -ForegroundColor Cyan

    if ($counts.CentralDeleted -gt 0) {
        Write-Host "  Deleted (central):$($counts.CentralDeleted)" -ForegroundColor Magenta
    }
    if ($counts.LocalDeleted -gt 0) {
        Write-Host "  Deleted (local):  $($counts.LocalDeleted)" -ForegroundColor Magenta
    }

    Write-Host ('-' * 60) -ForegroundColor White
    Write-Host ''
}


# ============================================================================
#  Public: Receive-MainSwitch (pull)
# ============================================================================

function Receive-MainSwitch {
    <#
        .SYNOPSIS
            Pulls changes from the central Main-Switch.ps1 into your local copy.
        .DESCRIPTION
            Parses both files, compares them using three-way merge, auto-merges
            non-conflicting changes, and prompts for conflict resolution.
            Creates a timestamped backup before writing.
        .PARAMETER LocalPath
            Path to your local Main-Switch.ps1. Defaults to the path from Paths.txt.
        .PARAMETER CentralPath
            Path to the shared central Main-Switch.ps1. Defaults to M:\Share\VMT\Scripts\Main-Switch.ps1.
        .PARAMETER NoBackup
            Skip creating a backup of the local file before writing.
        .PARAMETER RefreshBase
            Overwrite the local .base file with the current central file and
            exit without merging. Use this to recover from a stale base after
            another admin ran Submit-MainSwitch -Initialize -Force. Your local
            file is not touched.
        .EXAMPLE
            Receive-MainSwitch
        .EXAMPLE
            Receive-MainSwitch -CentralPath "\\OtherServer\Share\Main-Switch.ps1"
        .EXAMPLE
            Receive-MainSwitch -RefreshBase
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [string]$LocalPath,

        [Parameter()]
        [string]$CentralPath = 'M:\Share\VMT\Scripts\Main-Switch\Main-Switch.ps1',

        [Parameter()]
        [switch]$NoBackup,

        [Parameter()]
        [switch]$RefreshBase
    )

    # Resolve local path default
    if (-not $LocalPath) {
        $scriptDir = Get-DefaultScriptPath
        if ($scriptDir) {
            $LocalPath = Join-Path $scriptDir 'Main-Switch.ps1'
        }
        else {
            throw 'Could not resolve local script path. Provide -LocalPath or run Setup.ps1 first.'
        }
    }

    if (-not (Test-Path -LiteralPath $LocalPath)) {
        throw "Local file not found: $LocalPath"
    }

    # --- RefreshBase mode: overwrite local base with central, no merge ---
    if ($RefreshBase) {
        if (-not (Test-Path -LiteralPath $CentralPath)) {
            throw "Central file not found: $CentralPath"
        }

        $basePath = "$LocalPath.base"

        Write-Host ''
        Write-Host ('-' * 60) -ForegroundColor White
        Write-Host '  Receive-MainSwitch -RefreshBase: Syncing base to central' -ForegroundColor White
        Write-Host ('-' * 60) -ForegroundColor White
        Write-Host ''
        Write-Host "  Central: $CentralPath" -ForegroundColor Cyan
        Write-Host "  Base:    $basePath" -ForegroundColor Cyan
        Write-Host ''

        if ($PSCmdlet.ShouldProcess($basePath, 'Overwrite base file with central')) {
            Copy-Item -LiteralPath $CentralPath -Destination $basePath -Force
            Write-Host "  Base file refreshed: $basePath" -ForegroundColor Green
            Write-Host '  Local file unchanged. Submit-MainSwitch should no longer be blocked.' -ForegroundColor DarkGray
        }

        Write-Host ''
        return
    }

    # Try reading central with error handling for file locks
    try {
        if (-not (Test-Path -LiteralPath $CentralPath)) {
            throw "Central file not found: $CentralPath"
        }
        $central = ConvertTo-ParsedMainSwitch -Path $CentralPath
    }
    catch [System.IO.IOException] {
        throw "Central file is locked by another user. Try again later. Details: $($_.Exception.Message)"
    }

    $local = ConvertTo-ParsedMainSwitch -Path $LocalPath

    # Check for base file
    $basePath = "$LocalPath.base"
    $base = if (Test-Path -LiteralPath $basePath) {
        ConvertTo-ParsedMainSwitch -Path $basePath
    } else { $null }

    Write-Host ''
    Write-Host ('-' * 60) -ForegroundColor White
    Write-Host '  Receive-MainSwitch: Merging central changes into local' -ForegroundColor White
    Write-Host ('-' * 60) -ForegroundColor White

    if (-not $base) {
        Write-Host '  First sync -- no base file found. All differences will be flagged as conflicts.' -ForegroundColor DarkYellow
        Write-Host '  After this sync, a base file will be created for future three-way merges.' -ForegroundColor DarkGray
    }
    Write-Host ''

    $comparison = Compare-ParsedFiles -Local $local -Central $central -Base $base
    $mergeResult = Merge-ParsedFiles -Local $local -Central $central -Comparison $comparison -MergeDirection 'Pull'

    $stats = $mergeResult.Stats
    $totalChanges = $stats.AutoMerged + $stats.NewEntries + $stats.Conflicts + $stats.Deleted

    if ($totalChanges -eq 0) {
        Write-Host '  Already up to date -- no changes needed.' -ForegroundColor Green
        Write-Host ''

        # Still create base if it does not exist
        if (-not (Test-Path -LiteralPath $basePath)) {
            if ($PSCmdlet.ShouldProcess($basePath, 'Create base file for future syncs')) {
                Copy-Item -LiteralPath $CentralPath -Destination $basePath -Force
                Write-Host "  Base file created: $basePath" -ForegroundColor DarkGray
            }
        }
        elseif ($null -ne $base) {
            # Self-heal: local matches central but base may be stale (for example,
            # another admin ran Submit-MainSwitch -Initialize -Force and overwrote
            # central). Without this, Submit-MainSwitch would stay blocked because
            # its publish check compares base vs central as whole files.
            $centralNorm = Get-NormalizedText -Text (ConvertFrom-ParsedMainSwitch -Parsed $central)
            $baseNorm    = Get-NormalizedText -Text (ConvertFrom-ParsedMainSwitch -Parsed $base)
            if ($centralNorm -ne $baseNorm) {
                if ($PSCmdlet.ShouldProcess($basePath, 'Refresh stale base file to match central')) {
                    Copy-Item -LiteralPath $CentralPath -Destination $basePath -Force
                    Write-Host "  Base file refreshed (was stale vs central): $basePath" -ForegroundColor DarkGray
                }
            }
        }
        return
    }

    # Display summary
    Write-Host ''
    Write-Host ('-' * 60) -ForegroundColor White
    Write-Host '  Merge Summary:' -ForegroundColor White
    Write-Host "    Auto-merged:  $($stats.AutoMerged) case(s)" -ForegroundColor Cyan
    Write-Host "    New entries:  $($stats.NewEntries) case(s)" -ForegroundColor Cyan
    Write-Host "    Conflicts:    $($stats.Conflicts) case(s)" -ForegroundColor Yellow
    Write-Host "    Deleted:      $($stats.Deleted) case(s)" -ForegroundColor Magenta
    Write-Host ('-' * 60) -ForegroundColor White
    Write-Host ''

    # Write merged file
    $mergedContent = ConvertFrom-ParsedMainSwitch -Parsed $mergeResult.Merged

    if ($PSCmdlet.ShouldProcess($LocalPath, 'Write merged Main-Switch.ps1')) {

        # Backup
        if (-not $NoBackup) {
            $timestamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
            $backupDir  = Join-Path (Split-Path $LocalPath -Parent) 'Patching\SwitchBackups'
            $backupName = "Main-Switch_$timestamp.bak"
            $backupPath = Join-Path $backupDir $backupName
            New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
            Copy-Item -LiteralPath $LocalPath -Destination $backupPath -Force
            Write-Host "  Backup created: $backupPath" -ForegroundColor DarkGray
        }

        Write-MainSwitchFile -Path $LocalPath -Content $mergedContent

        # Update base file to match central
        Copy-Item -LiteralPath $CentralPath -Destination $basePath -Force

        Write-Host "  Local file updated: $LocalPath" -ForegroundColor Green
        Write-Host "  Base file updated:  $basePath" -ForegroundColor DarkGray
    }

    Write-Host ''
}


# ============================================================================
#  Public: Submit-MainSwitch (push)
# ============================================================================

function Submit-MainSwitch {
    <#
        .SYNOPSIS
            Pushes your local Main-Switch.ps1 changes to the central shared copy.
        .DESCRIPTION
            Checks that you have synced first, then merges your local changes into
            the central copy. Creates a timestamped backup of central before writing.
        .PARAMETER LocalPath
            Path to your local Main-Switch.ps1. Defaults to the path from Paths.txt.
        .PARAMETER CentralPath
            Path to the shared central Main-Switch.ps1. Defaults to M:\Share\VMT\Scripts\Main-Switch.ps1.
        .PARAMETER NoBackup
            Skip creating a backup of the central file before writing.
        .PARAMETER Force
            Bypass the check that requires Receive-MainSwitch to be run first.
        .PARAMETER Initialize
            Initialize a new central location by copying your local file there.
            Creates parent directories if needed and sets the base file so future
            pull/push operations work normally. Fails if the central file already
            exists (use -Force to overwrite). Aliases: -Init, -New
        .EXAMPLE
            Submit-MainSwitch
        .EXAMPLE
            Submit-MainSwitch -Force
        .EXAMPLE
            Submit-MainSwitch -Initialize -CentralPath '\\NewServer\Share\Main-Switch\Main-Switch.ps1'
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [string]$LocalPath,

        [Parameter()]
        [string]$CentralPath = 'M:\Share\VMT\Scripts\Main-Switch\Main-Switch.ps1',

        [Parameter()]
        [switch]$NoBackup,

        [Parameter()]
        [switch]$Force,

        [Parameter()]
        [Alias('Init', 'New')]
        [switch]$Initialize
    )

    # Resolve local path default
    if (-not $LocalPath) {
        $scriptDir = Get-DefaultScriptPath
        if ($scriptDir) {
            $LocalPath = Join-Path $scriptDir 'Main-Switch.ps1'
        }
        else {
            throw 'Could not resolve local script path. Provide -LocalPath or run Setup.ps1 first.'
        }
    }

    if (-not (Test-Path -LiteralPath $LocalPath)) {
        throw "Local file not found: $LocalPath"
    }

    # --- Initialize mode: create a new central location from local ---
    if ($Initialize) {
        if ((Test-Path -LiteralPath $CentralPath) -and -not $Force) {
            Write-Host ''
            Write-Host '  INITIALIZE BLOCKED' -ForegroundColor Red
            Write-Host "  Central file already exists: $CentralPath" -ForegroundColor Yellow
            Write-Host '  Use -Initialize -Force to overwrite, or regular Submit-MainSwitch to push changes.' -ForegroundColor Yellow
            Write-Host ''
            return
        }

        # Validate local file parses correctly before initializing
        $null = ConvertTo-ParsedMainSwitch -Path $LocalPath

        Write-Host ''
        Write-Host ('-' * 60) -ForegroundColor White
        Write-Host '  Submit-MainSwitch -Initialize: Creating new central location' -ForegroundColor White
        Write-Host ('-' * 60) -ForegroundColor White
        Write-Host ''
        Write-Host "  Source:      $LocalPath" -ForegroundColor Cyan
        Write-Host "  Destination: $CentralPath" -ForegroundColor Cyan
        Write-Host ''

        if ($PSCmdlet.ShouldProcess($CentralPath, 'Initialize new central Main-Switch.ps1 from local')) {
            $centralDir = Split-Path $CentralPath -Parent
            if (-not (Test-Path -LiteralPath $centralDir)) {
                New-Item -ItemType Directory -Force -Path $centralDir | Out-Null
                Write-Host "  Created directory: $centralDir" -ForegroundColor DarkGray
            }

            Copy-Item -LiteralPath $LocalPath -Destination $CentralPath -Force
            Write-Host "  Central file created: $CentralPath" -ForegroundColor Green

            # Set base file so future pull/push syncs work
            $basePath = "$LocalPath.base"
            Copy-Item -LiteralPath $CentralPath -Destination $basePath -Force
            Write-Host "  Base file updated:    $basePath" -ForegroundColor DarkGray
        }

        Write-Host ''
        return
    }

    # Try reading central with error handling for file locks
    try {
        if (-not (Test-Path -LiteralPath $CentralPath)) {
            throw "Central file not found: $CentralPath"
        }
        $central = ConvertTo-ParsedMainSwitch -Path $CentralPath
    }
    catch [System.IO.IOException] {
        throw "Central file is locked by another user. Try again later. Details: $($_.Exception.Message)"
    }

    $local = ConvertTo-ParsedMainSwitch -Path $LocalPath

    # Check for base file
    $basePath = "$LocalPath.base"
    $base = if (Test-Path -LiteralPath $basePath) {
        ConvertTo-ParsedMainSwitch -Path $basePath
    } else { $null }

    # --- Publish gate: check if central has diverged ---
    if (-not $Force -and $null -ne $base) {
        $centralNorm = Get-NormalizedText -Text (ConvertFrom-ParsedMainSwitch -Parsed $central)
        $baseNorm    = Get-NormalizedText -Text (ConvertFrom-ParsedMainSwitch -Parsed $base)

        if ($centralNorm -ne $baseNorm) {
            Write-Host ''
            Write-Host '  PUBLISH BLOCKED' -ForegroundColor Red
            Write-Host '  The central file has changed since your last sync.' -ForegroundColor Yellow
            Write-Host '  Run Receive-MainSwitch first to pull those changes, then try again.' -ForegroundColor Yellow
            Write-Host '  Use -Force to bypass this check (not recommended).' -ForegroundColor DarkGray
            Write-Host ''
            return
        }
    }

    if (-not $Force -and $null -eq $base) {
        Write-Host ''
        Write-Host '  WARNING: No base file found. Run Receive-MainSwitch first to establish a baseline.' -ForegroundColor Yellow
        Write-Host '  Use -Force to bypass this check.' -ForegroundColor DarkGray
        Write-Host ''
        return
    }

    Write-Host ''
    Write-Host ('-' * 60) -ForegroundColor White
    Write-Host '  Submit-MainSwitch: Merging local changes into central' -ForegroundColor White
    Write-Host ('-' * 60) -ForegroundColor White
    Write-Host ''

    $comparison  = Compare-ParsedFiles -Local $local -Central $central -Base $base
    $mergeResult = Merge-ParsedFiles -Local $local -Central $central -Comparison $comparison -MergeDirection 'Push'

    $stats = $mergeResult.Stats
    $totalChanges = $stats.AutoMerged + $stats.NewEntries + $stats.Conflicts + $stats.Deleted

    if ($totalChanges -eq 0) {
        Write-Host '  Nothing to publish -- central is already up to date.' -ForegroundColor Green
        Write-Host ''
        return
    }

    # Display summary
    Write-Host ''
    Write-Host ('-' * 60) -ForegroundColor White
    Write-Host '  Merge Summary:' -ForegroundColor White
    Write-Host "    Auto-merged:  $($stats.AutoMerged) case(s)" -ForegroundColor Green
    Write-Host "    New entries:  $($stats.NewEntries) case(s)" -ForegroundColor Green
    Write-Host "    Conflicts:    $($stats.Conflicts) case(s)" -ForegroundColor Yellow
    Write-Host "    Deleted:      $($stats.Deleted) case(s)" -ForegroundColor Magenta
    Write-Host ('-' * 60) -ForegroundColor White
    Write-Host ''

    # Write merged file
    $mergedContent = ConvertFrom-ParsedMainSwitch -Parsed $mergeResult.Merged

    if ($PSCmdlet.ShouldProcess($CentralPath, 'Write merged Main-Switch.ps1 to central')) {

        # Backup central
        if (-not $NoBackup) {
            $timestamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
            $backupDir  = Join-Path (Split-Path $CentralPath -Parent) 'SwitchBackups'
            $backupName = "Main-Switch_$timestamp.bak"
            $backupPath = Join-Path $backupDir $backupName
            try {
                New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
                Copy-Item -LiteralPath $CentralPath -Destination $backupPath -Force
                Write-Host "  Central backup created: $backupPath" -ForegroundColor DarkGray
            }
            catch {
                Write-Warning "Could not create backup: $($_.Exception.Message)"
                Write-Warning "Aborting publish. Use -NoBackup to skip backup (not recommended)."
                return
            }
        }

        try {
            Write-MainSwitchFile -Path $CentralPath -Content $mergedContent
        }
        catch [System.IO.IOException] {
            throw "Could not write to central file (locked?). Details: $($_.Exception.Message)"
        }

        # Update local base to match new central state
        Copy-Item -LiteralPath $CentralPath -Destination $basePath -Force

        Write-Host "  Central file updated: $CentralPath" -ForegroundColor Green
        Write-Host "  Base file updated:    $basePath" -ForegroundColor DarkGray
    }

    Write-Host ''
}


# ============================================================================
#  Aliases
# ============================================================================

New-Alias -Name 'Pull-MainSwitch' -Value 'Receive-MainSwitch' -Force
New-Alias -Name 'Push-MainSwitch'  -Value 'Submit-MainSwitch'  -Force
