# DOTS formatting comment

<#
    .SYNOPSIS
        Comprehensive test suite for the Merge-MainSwitch module.
    .DESCRIPTION
        Tests all merge scenarios: round-trip parsing, one-sided changes, conflicts,
        new entries, deletions, multi-line value changes, header/footer changes,
        the publish gate, and full end-to-end workflows.

        Run from any location -- paths are resolved automatically.

        All case names, version strings, and content references are discovered
        dynamically from the current Main-Switch.ps1, so adding, removing, or
        renaming software entries will not break these tests.

        Tests that involve conflict resolution (interactive Read-Host prompts)
        are tested at the comparison/detection level rather than end-to-end,
        since Read-Host cannot be automated in a script.

    .NOTES
        Written by Skyler Werner
        Date: 2026/03/18
#>

# --- Setup ---
$repoRoot = Split-Path $PSScriptRoot -Parent
Import-Module "$repoRoot\Modules\Merge-MainSwitch" -Force

# Get private functions via module scope
$mod = Get-Module Merge-MainSwitch
$parse       = & $mod { ${function:ConvertTo-ParsedMainSwitch} }
$reconstruct = & $mod { ${function:ConvertFrom-ParsedMainSwitch} }
$compareFn   = & $mod { ${function:Compare-ParsedFiles} }
$mergeFn     = & $mod { ${function:Merge-ParsedFiles} }
$normalize   = & $mod { ${function:Get-NormalizedText} }

$originalPath = "$repoRoot\Scripts\Main-Switch.ps1"
$testDir = Join-Path $env:TEMP 'MergeMainSwitch_Tests'

# Counters
$script:passed = 0
$script:failed = 0
$script:testNum = 0

function Assert-True {
    param([bool]$Condition, [string]$Message)
    $script:testNum++
    if ($Condition) {
        $script:passed++
        Write-Host "    PASS: $Message" -ForegroundColor Green
    } else {
        $script:failed++
        Write-Host "    FAIL: $Message" -ForegroundColor Red
    }
}

function Assert-Equal {
    param($Expected, $Actual, [string]$Message)
    $script:testNum++
    if ($Expected -eq $Actual) {
        $script:passed++
        Write-Host "    PASS: $Message" -ForegroundColor Green
    } else {
        $script:failed++
        Write-Host "    FAIL: $Message" -ForegroundColor Red
        Write-Host "          Expected: [$Expected]" -ForegroundColor Yellow
        Write-Host "          Actual:   [$Actual]" -ForegroundColor Magenta
    }
}

function New-TestEnvironment {
    # Clean and recreate test directory with base, local, and central copies
    if (Test-Path $testDir) { Remove-Item $testDir -Recurse -Force }
    New-Item -ItemType Directory -Path $testDir | Out-Null

    $paths = @{
        Base    = Join-Path $testDir 'Main-Switch.base.ps1'
        Local   = Join-Path $testDir 'Main-Switch.ps1'
        Central = Join-Path $testDir 'Central-Main-Switch.ps1'
    }

    Copy-Item -Path $originalPath -Destination $paths.Base
    Copy-Item -Path $originalPath -Destination $paths.Local
    Copy-Item -Path $originalPath -Destination $paths.Central

    return $paths
}

function Edit-FileContent {
    param(
        [string]$Path,
        [string]$Find,
        [string]$Replace
    )
    $content = Get-Content -Path $Path -Encoding Default -Raw
    $content = $content.Replace($Find, $Replace)
    Set-Content -Path $Path -Value $content -Encoding Default -NoNewline
}

function Add-CaseToFile {
    param(
        [string]$Path,
        [string]$CaseText
    )
    $content = Get-Content -Path $Path -Encoding Default -Raw
    $content = $content.Replace('} # End switch', "$CaseText`r`n`r`n} # End switch")
    Set-Content -Path $Path -Value $content -Encoding Default -NoNewline
}

function Remove-CaseFromFile {
    param(
        [string]$Path,
        [string]$CaseKey
    )
    # Parse, remove the case, reconstruct
    $parsed = & $parse -Path $Path
    $parsed.Cases.Remove($CaseKey)
    $content = & $reconstruct -Parsed $parsed
    Set-Content -Path $Path -Value $content -Encoding Default -NoNewline
}


# --- Dynamic discovery (makes tests independent of specific Main-Switch content) ---
$baseParsed    = & $parse -Path $originalPath
$allCaseKeys   = @($baseParsed.Cases.Keys)
$baseCaseCount = $baseParsed.Cases.Count

# Helper: extract a quoted variable from a case block
function Get-CaseVar {
    param([string]$Block, [string]$VarName)
    if ($Block -match ('\$' + $VarName + '\s+=\s+"([^"]+)"')) {
        return $Matches[1]
    }
    return $null
}

# Cases with $compliantVer -- candidates for version-bump tests
$versionCases = @($allCaseKeys | Where-Object {
    $null -ne (Get-CaseVar $baseParsed.Cases[$_] 'compliantVer')
} | Sort-Object)

# Cases with $software -- candidates for software-name tests
$softwareCases = @($allCaseKeys | Where-Object {
    $null -ne (Get-CaseVar $baseParsed.Cases[$_] 'software')
} | Sort-Object)

# Cases with multi-line installLine (@())
$multiLineCases = @($allCaseKeys | Where-Object {
    $baseParsed.Cases[$_] -match '\$installLine\s+=\s+@\('
} | Sort-Object)

# Pick test candidates from sorted lists for deterministic, collision-free selection
$caseA    = $versionCases[0]   # version-bump tests (primary)
$caseB    = $versionCases[1]   # version-bump tests (secondary)
$caseAVer = Get-CaseVar $baseParsed.Cases[$caseA] 'compliantVer'
$caseBVer = Get-CaseVar $baseParsed.Cases[$caseB] 'compliantVer'
$escapedCaseAVer = [regex]::Escape($caseAVer)
$escapedCaseBVer = [regex]::Escape($caseBVer)

# Software-name case (for rename/conflict tests, different from caseA/caseB)
$usedKeys = @($caseA, $caseB)
$caseSW   = $softwareCases | Where-Object { $_ -notin $usedKeys } | Select-Object -First 1
$caseSWSoftware = Get-CaseVar $baseParsed.Cases[$caseSW] 'software'
$usedKeys += $caseSW

# Second software-name case (for conflict -- both sides change same case)
$caseCF   = $softwareCases | Where-Object { $_ -notin $usedKeys } | Select-Object -First 1
$caseCFSoftware = Get-CaseVar $baseParsed.Cases[$caseCF] 'software'
$usedKeys += $caseCF

# Third version case (for unrelated-change-preserved test)
$caseMisc    = $versionCases | Where-Object { $_ -notin $usedKeys } | Select-Object -First 1
$caseMiscVer = Get-CaseVar $baseParsed.Cases[$caseMisc] 'compliantVer'
$usedKeys += $caseMisc

# Deletion candidates (any cases not already reserved)
$delCandidates = @($allCaseKeys | Where-Object { $_ -notin $usedKeys } | Sort-Object)
$caseDel1 = $delCandidates[0]
$caseDel2 = $delCandidates[1]

# Multi-line case (if available)
$multiLineCase = if ($multiLineCases.Count -gt 0) { $multiLineCases[0] } else { $null }
$multiLineSnippet = $null
if ($multiLineCase) {
    # Grab a unique line from the multi-line installLine to use as edit target
    $mlBlock = $baseParsed.Cases[$multiLineCase]
    $mlLines = @($mlBlock -split "`r?`n" | Where-Object { $_ -match 'cmd /c|\.ps1|\.exe' })
    if ($mlLines.Count -gt 0) { $multiLineSnippet = $mlLines[-1].Trim() }
}

# Header edit target -- find a quoted path assignment in the header
$headerEditTarget = $null
$headerEditReplace = $null
if ($baseParsed.Header -match '(\$\w+\s+=\s+"[A-Z]:\\[^"]+")') {
    $headerEditTarget  = $Matches[1]
    $headerEditReplace = $headerEditTarget -replace '"[A-Z]:\\', '"X:\Replaced\'
}

# Footer edit target -- find a variable assignment in the footer
$footerEditTarget = $null
if ($baseParsed.Footer -match '(\s+\w+\s+=\s+\$\w+)\s*$') {
    $footerEditTarget = $Matches[1].TrimStart()
}


# ============================================================================
Write-Host ''
Write-Host ('=' * 70) -ForegroundColor White
Write-Host '  Merge-MainSwitch Test Suite' -ForegroundColor White
Write-Host ('=' * 70) -ForegroundColor White
Write-Host "  Cases discovered: $baseCaseCount"
Write-Host "  Test cases: A=$caseA  B=$caseB  SW=$caseSW  CF=$caseCF  Misc=$caseMisc"
Write-Host "  Deletion:   Del1=$caseDel1  Del2=$caseDel2"
if ($multiLineCase) { Write-Host "  Multi-line: $multiLineCase" }


# ============================================================================
#  TEST 1: Round-trip parsing fidelity
# ============================================================================
Write-Host "`n--- TEST 1: Round-trip parsing fidelity ---" -ForegroundColor Cyan

$parsed = & $parse -Path $originalPath
$reconstructed = & $reconstruct -Parsed $parsed
$original = Get-Content -Path $originalPath -Encoding Default -Raw

$origLines = $original -split "`r?`n"
$reconLines = $reconstructed -split "`r?`n"

Assert-Equal $origLines.Count $reconLines.Count 'Line count matches'

$allMatch = $true
for ($i = 0; $i -lt $origLines.Count; $i++) {
    if ($origLines[$i] -ne $reconLines[$i]) {
        $allMatch = $false
        Write-Host "      First diff at line $($i+1)" -ForegroundColor Red
        break
    }
}
Assert-True $allMatch 'All lines match byte-for-byte'


# ============================================================================
#  TEST 2: Parser extracts correct structure
# ============================================================================
Write-Host "`n--- TEST 2: Parser structure ---" -ForegroundColor Cyan

Assert-True ($parsed.Header.Length -gt 0) 'Header is non-empty'
Assert-True ($parsed.Footer.Length -gt 0) 'Footer is non-empty'
Assert-True ($parsed.Cases.Count -ge 50) "Case count is $($parsed.Cases.Count) (expected 50+)"
Assert-True ($parsed.Header -match 'switch \(\$targetSoftware\)') 'Header ends with switch statement'
Assert-True ($parsed.Footer -match '\} # End switch') 'Footer starts with End switch'
Assert-True ($parsed.Footer -match 'Return \$parameters') 'Footer contains Return statement'

# Verify discovered test cases exist (they must -- we picked them from the file)
Assert-True ($parsed.Cases.Contains($caseA)) "Case `"$caseA`" exists"
Assert-True ($parsed.Cases.Contains($caseB)) "Case `"$caseB`" exists"
Assert-True ($parsed.Cases.Contains($caseSW)) "Case `"$caseSW`" exists"
Assert-True ($parsed.Cases.Contains($caseCF)) "Case `"$caseCF`" exists"
Assert-True ($parsed.Cases.Contains($caseMisc)) "Case `"$caseMisc`" exists"


# ============================================================================
#  TEST 3: Parser handles multi-line values correctly
# ============================================================================
Write-Host "`n--- TEST 3: Multi-line value preservation ---" -ForegroundColor Cyan

if ($multiLineCase) {
    $mlBlock = $parsed.Cases[$multiLineCase]
    # Multi-line installLine uses @() syntax
    Assert-True ($mlBlock -match '\$installLine\s+=\s+@\(') "$multiLineCase has multi-line installLine (@() syntax)"
    # Count the lines inside @() -- should have at least 2
    $inArray = $false
    $arrayLineCount = 0
    foreach ($line in ($mlBlock -split "`r?`n")) {
        if ($line -match '\$installLine\s+=\s+@\(') { $inArray = $true; continue }
        if ($inArray -and $line -match '^\s*\)') { break }
        if ($inArray -and $line.Trim().Length -gt 0) { $arrayLineCount++ }
    }
    Assert-True ($arrayLineCount -ge 2) "$multiLineCase installLine has $arrayLineCount entries (expected 2+)"

    # Verify a second multi-line case if available
    if ($multiLineCases.Count -ge 2) {
        $ml2 = $multiLineCases[1]
        $ml2Block = $parsed.Cases[$ml2]
        Assert-True ($ml2Block -match '\$installLine\s+=\s+@\(') "$ml2 also has multi-line installLine"
        Assert-True ($ml2Block -match '\.ps1|\.exe') "$ml2 installLine references a script or executable"
    } else {
        # Only one multi-line case; verify it has executable references
        Assert-True ($mlBlock -match '\.ps1|\.exe') "$multiLineCase installLine references a script or executable"
        Assert-True $true '(Only 1 multi-line case found -- skipping second check)'
    }
} else {
    # No multi-line cases -- pass with a note
    Assert-True $true '(No multi-line installLine cases found -- skipping)'
    Assert-True $true '(No multi-line installLine cases found -- skipping)'
    Assert-True $true '(No multi-line installLine cases found -- skipping)'
    Assert-True $true '(No multi-line installLine cases found -- skipping)'
    Assert-True $true '(No multi-line installLine cases found -- skipping)'
}


# ============================================================================
#  TEST 4: Parser preserves region markers
# ============================================================================
Write-Host "`n--- TEST 4: Region marker preservation ---" -ForegroundColor Cyan

# Find all region markers across all case blocks
$regionBlocks = @($parsed.Cases.GetEnumerator() | Where-Object { $_.Value -match '#region' })
Assert-True ($regionBlocks.Count -gt 0) "Region markers found in $($regionBlocks.Count) case block(s)"

# Verify round-trip preserves them -- check reconstructed output
Assert-True ($reconstructed -match '#region') 'Region markers survive round-trip reconstruction'


# ============================================================================
#  TEST 5: Parser preserves preceding comments
# ============================================================================
Write-Host "`n--- TEST 5: Comment preservation ---" -ForegroundColor Cyan

# Comments between cases attach to the PRECEDING case's trailing text.
# Find any case block that contains a comment line (# followed by text, not a region)
$commentBlocks = @($parsed.Cases.GetEnumerator() | Where-Object {
    $_.Value -match '(?m)^\s*#\s+(?!region|endregion)[A-Z]'
})
Assert-True ($commentBlocks.Count -gt 0) "Comment lines found in $($commentBlocks.Count) case block(s)"

# Verify comments survive round-trip
$origComments = @($original -split "`r?`n" | Where-Object { $_ -match '^\s*#\s+Ready' })
$reconComments = @($reconstructed -split "`r?`n" | Where-Object { $_ -match '^\s*#\s+Ready' })
Assert-True ($origComments.Count -eq $reconComments.Count) 'Comment lines preserved in round-trip'


# ============================================================================
#  TEST 6: Identical files -- no changes detected
# ============================================================================
Write-Host "`n--- TEST 6: Identical files ---" -ForegroundColor Cyan

$paths = New-TestEnvironment
$base    = & $parse -Path $paths.Base
$local   = & $parse -Path $paths.Local
$central = & $parse -Path $paths.Central
$result  = & $compareFn -Local $local -Central $central -Base $base

Assert-Equal 'Identical' $result.HeaderStatus 'Header is Identical'
Assert-Equal 'Identical' $result.FooterStatus 'Footer is Identical'

$nonIdentical = @($result.CaseResults.Keys | Where-Object { $result.CaseResults[$_].Status -ne 'Identical' })
Assert-Equal 0 $nonIdentical.Count "All cases are Identical (found $($nonIdentical.Count) non-identical)"


# ============================================================================
#  TEST 7: Central changed one case (simple version bump)
# ============================================================================
Write-Host "`n--- TEST 7: Central changed one case ($caseA) ---" -ForegroundColor Cyan

$paths = New-TestEnvironment
Edit-FileContent -Path $paths.Central `
    -Find  "`$compliantVer  = `"$caseAVer`"" `
    -Replace '$compliantVer  = "999.0.7777.100"'

$base    = & $parse -Path $paths.Base
$local   = & $parse -Path $paths.Local
$central = & $parse -Path $paths.Central
$result  = & $compareFn -Local $local -Central $central -Base $base

Assert-Equal 'CentralChanged' $result.CaseResults[$caseA].Status "$caseA detected as CentralChanged"

# If another case shares the same version string, it should also be detected
$relatedCases = @($allCaseKeys | Where-Object {
    $_ -ne $caseA -and (Get-CaseVar $baseParsed.Cases[$_] 'compliantVer') -eq $caseAVer
})
if ($relatedCases.Count -gt 0) {
    $rel = $relatedCases[0]
    Assert-Equal 'CentralChanged' $result.CaseResults[$rel].Status "$rel also detected (shares version string)"
} else {
    Assert-True $true '(No related cases share version string -- skip cascade check)'
}
Assert-Equal 'Identical' $result.CaseResults[$caseB].Status "$caseB still Identical"


# ============================================================================
#  TEST 8: Local changed one case
# ============================================================================
Write-Host "`n--- TEST 8: Local changed one case ($caseSW) ---" -ForegroundColor Cyan

$paths = New-TestEnvironment
Edit-FileContent -Path $paths.Local `
    -Find  "`$software      = `"$caseSWSoftware`"" `
    -Replace "`$software      = `"$caseSWSoftware MODIFIED`""

$base    = & $parse -Path $paths.Base
$local   = & $parse -Path $paths.Local
$central = & $parse -Path $paths.Central
$result  = & $compareFn -Local $local -Central $central -Base $base

Assert-Equal 'LocalChanged' $result.CaseResults[$caseSW].Status "$caseSW detected as LocalChanged"

# If another case shares the same $software string, it should also be detected
$relatedSW = @($allCaseKeys | Where-Object {
    $_ -ne $caseSW -and (Get-CaseVar $baseParsed.Cases[$_] 'software') -eq $caseSWSoftware
})
if ($relatedSW.Count -gt 0) {
    $rel = $relatedSW[0]
    Assert-Equal 'LocalChanged' $result.CaseResults[$rel].Status "$rel also detected (shares software string)"
} else {
    Assert-True $true '(No related cases share software string -- skip cascade check)'
}


# ============================================================================
#  TEST 9: Both changed same case -- CONFLICT
# ============================================================================
Write-Host "`n--- TEST 9: Conflict detection ($caseCF) ---" -ForegroundColor Cyan

$paths = New-TestEnvironment
# Local changes software name one way
Edit-FileContent -Path $paths.Local `
    -Find  "`$software      = `"$caseCFSoftware`"" `
    -Replace "`$software      = `"$caseCFSoftware LocalEdit`""
# Central changes it a different way
Edit-FileContent -Path $paths.Central `
    -Find  "`$software      = `"$caseCFSoftware`"" `
    -Replace "`$software      = `"$caseCFSoftware CentralEdit`""

$base    = & $parse -Path $paths.Base
$local   = & $parse -Path $paths.Local
$central = & $parse -Path $paths.Central
$result  = & $compareFn -Local $local -Central $central -Base $base

Assert-Equal 'Conflict' $result.CaseResults[$caseCF].Status "$caseCF detected as Conflict"


# ============================================================================
#  TEST 10: Both changed DIFFERENT cases -- no conflict
# ============================================================================
Write-Host "`n--- TEST 10: Both changed different cases (no conflict) ---" -ForegroundColor Cyan

$paths = New-TestEnvironment
# Local changes caseA version
Edit-FileContent -Path $paths.Local `
    -Find  "`$compliantVer  = `"$caseAVer`"" `
    -Replace '$compliantVer  = "999.0.3900.50"'
# Central changes caseB version
Edit-FileContent -Path $paths.Central `
    -Find  "`$compliantVer  = `"$caseBVer`"" `
    -Replace '$compliantVer  = "999.0.8000.100"'

$base    = & $parse -Path $paths.Base
$local   = & $parse -Path $paths.Local
$central = & $parse -Path $paths.Central
$result  = & $compareFn -Local $local -Central $central -Base $base

Assert-Equal 'LocalChanged' $result.CaseResults[$caseA].Status "$caseA is LocalChanged"
Assert-Equal 'CentralChanged' $result.CaseResults[$caseB].Status "$caseB is CentralChanged"
Assert-Equal 'Identical' $result.CaseResults[$caseMisc].Status "$caseMisc untouched"


# ============================================================================
#  TEST 11: New case in central only
# ============================================================================
Write-Host "`n--- TEST 11: New case in central only ---" -ForegroundColor Cyan

$paths = New-TestEnvironment
$newCase = @'

"TestApp" {
    $listPath      = "$listPathRoot\Test\TestApp.txt"
    $software      = "Test Application"
    $processName   = "testapp"
    $compliantVer  = "1.0.0"
    $patchPath     = "$patchRoot\Test\TestApp_1.0.0"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files\TestApp\app.exe"
    $installLine   = "& cmd /c 'C:\Temp\TestApp_1.0.0\setup.exe /S'"
}
'@
Add-CaseToFile -Path $paths.Central -CaseText $newCase

$base    = & $parse -Path $paths.Base
$local   = & $parse -Path $paths.Local
$central = & $parse -Path $paths.Central
$result  = & $compareFn -Local $local -Central $central -Base $base

Assert-Equal 'CentralOnly' $result.CaseResults['TestApp'].Status 'TestApp detected as CentralOnly'


# ============================================================================
#  TEST 12: New case in local only
# ============================================================================
Write-Host "`n--- TEST 12: New case in local only ---" -ForegroundColor Cyan

$paths = New-TestEnvironment
$newCase = @'

"MyLocalApp" {
    $listPath      = "$listPathRoot\Test\MyLocalApp.txt"
    $software      = "My Local Application"
    $processName   = "mylocalapp"
    $compliantVer  = "2.0.0"
    $patchPath     = "$patchRoot\Test\MyLocalApp_2.0.0"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files\MyLocalApp\app.exe"
    $installLine   = "& cmd /c 'C:\Temp\MyLocalApp_2.0.0\setup.exe /S'"
}
'@
Add-CaseToFile -Path $paths.Local -CaseText $newCase

$base    = & $parse -Path $paths.Base
$local   = & $parse -Path $paths.Local
$central = & $parse -Path $paths.Central
$result  = & $compareFn -Local $local -Central $central -Base $base

Assert-Equal 'LocalOnly' $result.CaseResults['MyLocalApp'].Status 'MyLocalApp detected as LocalOnly'


# ============================================================================
#  TEST 13: Case deleted from central
# ============================================================================
Write-Host "`n--- TEST 13: Case deleted from central ($caseDel1) ---" -ForegroundColor Cyan

$paths = New-TestEnvironment
Remove-CaseFromFile -Path $paths.Central -CaseKey $caseDel1

$base    = & $parse -Path $paths.Base
$local   = & $parse -Path $paths.Local
$central = & $parse -Path $paths.Central
$result  = & $compareFn -Local $local -Central $central -Base $base

Assert-Equal 'CentralDeleted' $result.CaseResults[$caseDel1].Status "$caseDel1 detected as CentralDeleted"


# ============================================================================
#  TEST 14: Case deleted from local
# ============================================================================
Write-Host "`n--- TEST 14: Case deleted from local ($caseDel2) ---" -ForegroundColor Cyan

$paths = New-TestEnvironment
Remove-CaseFromFile -Path $paths.Local -CaseKey $caseDel2

$base    = & $parse -Path $paths.Base
$local   = & $parse -Path $paths.Local
$central = & $parse -Path $paths.Central
$result  = & $compareFn -Local $local -Central $central -Base $base

Assert-Equal 'LocalDeleted' $result.CaseResults[$caseDel2].Status "$caseDel2 detected as LocalDeleted"


# ============================================================================
#  TEST 15: Multi-line installLine change detected
# ============================================================================
Write-Host "`n--- TEST 15: Multi-line value change ---" -ForegroundColor Cyan

if ($multiLineCase -and $multiLineSnippet) {
    $paths = New-TestEnvironment
    Edit-FileContent -Path $paths.Central `
        -Find  $multiLineSnippet `
        -Replace ($multiLineSnippet -replace '\\', '\\REPLACED\\')

    $base    = & $parse -Path $paths.Base
    $local   = & $parse -Path $paths.Local
    $central = & $parse -Path $paths.Central
    $result  = & $compareFn -Local $local -Central $central -Base $base

    Assert-Equal 'CentralChanged' $result.CaseResults[$multiLineCase].Status "$multiLineCase multi-line change detected"
} else {
    Assert-True $true '(No multi-line case with editable snippet -- skipping)'
}


# ============================================================================
#  TEST 16: Header change (one side only)
# ============================================================================
Write-Host "`n--- TEST 16: Header change ---" -ForegroundColor Cyan

if ($headerEditTarget) {
    $paths = New-TestEnvironment
    Edit-FileContent -Path $paths.Central `
        -Find  $headerEditTarget `
        -Replace $headerEditReplace

    $base    = & $parse -Path $paths.Base
    $local   = & $parse -Path $paths.Local
    $central = & $parse -Path $paths.Central
    $result  = & $compareFn -Local $local -Central $central -Base $base

    Assert-Equal 'CentralChanged' $result.HeaderStatus 'Header detected as CentralChanged'
    Assert-Equal 'Identical' $result.FooterStatus 'Footer still Identical'
} else {
    Assert-True $true '(No header edit target found -- skipping)'
    Assert-True $true '(No header edit target found -- skipping)'
}


# ============================================================================
#  TEST 17: Footer change (one side only)
# ============================================================================
Write-Host "`n--- TEST 17: Footer change ---" -ForegroundColor Cyan

if ($footerEditTarget) {
    $paths = New-TestEnvironment
    Edit-FileContent -Path $paths.Central `
        -Find  $footerEditTarget `
        -Replace "$footerEditTarget`r`n    NewTestField  = `$newTestField"

    $base    = & $parse -Path $paths.Base
    $local   = & $parse -Path $paths.Local
    $central = & $parse -Path $paths.Central
    $result  = & $compareFn -Local $local -Central $central -Base $base

    Assert-Equal 'CentralChanged' $result.FooterStatus 'Footer detected as CentralChanged'
} else {
    Assert-True $true '(No footer edit target found -- skipping)'
}


# ============================================================================
#  TEST 18: Auto-merge (pull) applies central changes to local
# ============================================================================
Write-Host "`n--- TEST 18: Auto-merge pull ---" -ForegroundColor Cyan

$paths = New-TestEnvironment
# Central bumps caseB version
Edit-FileContent -Path $paths.Central `
    -Find  "`$compliantVer  = `"$caseBVer`"" `
    -Replace '$compliantVer  = "999.0.8000.100"'
# Local bumps caseA version (different case -- no conflict)
Edit-FileContent -Path $paths.Local `
    -Find  "`$compliantVer  = `"$caseAVer`"" `
    -Replace '$compliantVer  = "999.0.3900.50"'

$base    = & $parse -Path $paths.Base
$local   = & $parse -Path $paths.Local
$central = & $parse -Path $paths.Central
$comp    = & $compareFn -Local $local -Central $central -Base $base
$merge   = & $mergeFn -Local $local -Central $central -Comparison $comp -MergeDirection 'Pull'

$mergedContent = & $reconstruct -Parsed $merge.Merged

# Merged should have BOTH changes
Assert-True ($mergedContent -match '999\.0\.3900\.50') 'Merged contains local version bump'
Assert-True ($mergedContent -match '999\.0\.8000\.100') 'Merged contains central version bump'
Assert-True ($merge.Stats.AutoMerged -gt 0) "AutoMerged count > 0 (got $($merge.Stats.AutoMerged))"
Assert-Equal 0 $merge.Stats.Conflicts 'No conflicts'


# ============================================================================
#  TEST 19: Auto-merge (pull) adds new central entry
# ============================================================================
Write-Host "`n--- TEST 19: Auto-merge pulls new entry ---" -ForegroundColor Cyan

$paths = New-TestEnvironment
$newCase = @'

"PulledApp" {
    $listPath      = "$listPathRoot\Test\PulledApp.txt"
    $software      = "Pulled Application"
    $processName   = "pulledapp"
    $compliantVer  = "3.0.0"
    $patchPath     = "$patchRoot\Test\PulledApp_3.0.0"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files\PulledApp\app.exe"
    $installLine   = "& cmd /c 'C:\Temp\PulledApp_3.0.0\setup.exe /S'"
}
'@
Add-CaseToFile -Path $paths.Central -CaseText $newCase

$base    = & $parse -Path $paths.Base
$local   = & $parse -Path $paths.Local
$central = & $parse -Path $paths.Central
$comp    = & $compareFn -Local $local -Central $central -Base $base
$merge   = & $mergeFn -Local $local -Central $central -Comparison $comp -MergeDirection 'Pull'

Assert-True ($merge.Merged.Cases.Contains('PulledApp')) 'Merged contains new PulledApp case'
Assert-Equal 1 $merge.Stats.NewEntries 'NewEntries count is 1'

$mergedContent = & $reconstruct -Parsed $merge.Merged
Assert-True ($mergedContent -match 'Pulled Application') 'Merged content contains PulledApp software name'


# ============================================================================
#  TEST 20: Auto-merge (push) applies local changes to central
# ============================================================================
Write-Host "`n--- TEST 20: Auto-merge push ---" -ForegroundColor Cyan

$paths = New-TestEnvironment
Edit-FileContent -Path $paths.Local `
    -Find  "`$compliantVer  = `"$caseAVer`"" `
    -Replace '$compliantVer  = "999.0.3900.50"'

$base    = & $parse -Path $paths.Base
$local   = & $parse -Path $paths.Local
$central = & $parse -Path $paths.Central
$comp    = & $compareFn -Local $local -Central $central -Base $base
$merge   = & $mergeFn -Local $local -Central $central -Comparison $comp -MergeDirection 'Push'

$mergedContent = & $reconstruct -Parsed $merge.Merged
Assert-True ($mergedContent -match '999\.0\.3900\.50') 'Push merged local version into central'
Assert-True ($merge.Stats.AutoMerged -gt 0) "Push AutoMerged count > 0 (got $($merge.Stats.AutoMerged))"


# ============================================================================
#  TEST 21: Auto-merge (push) adds new local entry to central
# ============================================================================
Write-Host "`n--- TEST 21: Auto-merge push adds new entry ---" -ForegroundColor Cyan

$paths = New-TestEnvironment
$newCase = @'

"PushedApp" {
    $listPath      = "$listPathRoot\Test\PushedApp.txt"
    $software      = "Pushed Application"
    $processName   = "pushedapp"
    $compliantVer  = "4.0.0"
    $patchPath     = "$patchRoot\Test\PushedApp_4.0.0"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files\PushedApp\app.exe"
    $installLine   = "& cmd /c 'C:\Temp\PushedApp_4.0.0\setup.exe /S'"
}
'@
Add-CaseToFile -Path $paths.Local -CaseText $newCase

$base    = & $parse -Path $paths.Base
$local   = & $parse -Path $paths.Local
$central = & $parse -Path $paths.Central
$comp    = & $compareFn -Local $local -Central $central -Base $base
$merge   = & $mergeFn -Local $local -Central $central -Comparison $comp -MergeDirection 'Push'

Assert-True ($merge.Merged.Cases.Contains('PushedApp')) 'Push merged contains new PushedApp case'
Assert-Equal 1 $merge.Stats.NewEntries 'Push NewEntries count is 1'


# ============================================================================
#  TEST 22: Merge preserves local changes when pulling unrelated central changes
# ============================================================================
Write-Host "`n--- TEST 22: Pull preserves unrelated local changes ---" -ForegroundColor Cyan

$paths = New-TestEnvironment
# Local changes caseMisc version
Edit-FileContent -Path $paths.Local `
    -Find  "`$compliantVer  = `"$caseMiscVer`"" `
    -Replace '$compliantVer  = "999.0.2222.0"'
# Central changes caseSW software name (completely different case)
Edit-FileContent -Path $paths.Central `
    -Find  "`$software      = `"$caseSWSoftware`"" `
    -Replace "`$software      = `"$caseSWSoftware Updated`""

$base    = & $parse -Path $paths.Base
$local   = & $parse -Path $paths.Local
$central = & $parse -Path $paths.Central
$comp    = & $compareFn -Local $local -Central $central -Base $base
$merge   = & $mergeFn -Local $local -Central $central -Comparison $comp -MergeDirection 'Pull'

$mergedContent = & $reconstruct -Parsed $merge.Merged
Assert-True ($mergedContent -match '999\.0\.2222\.0') "Local $caseMisc version preserved after pull"
Assert-True ($mergedContent -match [regex]::Escape("$caseSWSoftware Updated")) "Central $caseSW change pulled in"


# ============================================================================
#  TEST 23: Two-way comparison (no base) treats diffs as Conflict
# ============================================================================
Write-Host "`n--- TEST 23: Two-way comparison (no base) ---" -ForegroundColor Cyan

$paths = New-TestEnvironment
Edit-FileContent -Path $paths.Central `
    -Find  "`$compliantVer  = `"$caseAVer`"" `
    -Replace '$compliantVer  = "999.0.0.0"'

$local   = & $parse -Path $paths.Local
$central = & $parse -Path $paths.Central
# No base passed
$result  = & $compareFn -Local $local -Central $central

Assert-Equal 'Conflict' $result.CaseResults[$caseA].Status "Without base, diff is Conflict"
Assert-Equal 'Identical' $result.CaseResults[$caseB].Status "Unchanged case still Identical without base"


# ============================================================================
#  TEST 24: Central deleted case handled correctly in merge
# ============================================================================
Write-Host "`n--- TEST 24: Central deletion applied in pull merge ($caseDel1) ---" -ForegroundColor Cyan

$paths = New-TestEnvironment
Remove-CaseFromFile -Path $paths.Central -CaseKey $caseDel1

$base    = & $parse -Path $paths.Base
$local   = & $parse -Path $paths.Local
$central = & $parse -Path $paths.Central
$comp    = & $compareFn -Local $local -Central $central -Base $base
$merge   = & $mergeFn -Local $local -Central $central -Comparison $comp -MergeDirection 'Pull'

Assert-True (-not $merge.Merged.Cases.Contains($caseDel1)) "$caseDel1 removed from merged result"
Assert-Equal 1 $merge.Stats.Deleted "Deleted count is 1 (got $($merge.Stats.Deleted))"


# ============================================================================
#  TEST 25: Local deletion applied in push merge
# ============================================================================
Write-Host "`n--- TEST 25: Local deletion applied in push merge ($caseDel2) ---" -ForegroundColor Cyan

$paths = New-TestEnvironment
Remove-CaseFromFile -Path $paths.Local -CaseKey $caseDel2

$base    = & $parse -Path $paths.Base
$local   = & $parse -Path $paths.Local
$central = & $parse -Path $paths.Central
$comp    = & $compareFn -Local $local -Central $central -Base $base
$merge   = & $mergeFn -Local $local -Central $central -Comparison $comp -MergeDirection 'Push'

Assert-True (-not $merge.Merged.Cases.Contains($caseDel2)) "$caseDel2 removed from push merged result"
Assert-Equal 1 $merge.Stats.Deleted "Push deleted count is 1 (got $($merge.Stats.Deleted))"


# ============================================================================
#  TEST 26: Mixed scenario -- multiple simultaneous changes
# ============================================================================
Write-Host "`n--- TEST 26: Mixed scenario ---" -ForegroundColor Cyan

$paths = New-TestEnvironment

# Local: bump caseA, add new case
Edit-FileContent -Path $paths.Local `
    -Find  "`$compliantVer  = `"$caseAVer`"" `
    -Replace '$compliantVer  = "999.0.3900.50"'
$newLocalCase = @'

"AdminAApp" {
    $listPath      = "$listPathRoot\Test\AdminAApp.txt"
    $software      = "Admin A Application"
    $processName   = "adminaapp"
    $compliantVer  = "1.0.0"
    $patchPath     = "$patchRoot\Test\AdminAApp_1.0.0"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files\AdminAApp\app.exe"
    $installLine   = "& cmd /c 'C:\Temp\AdminAApp_1.0.0\setup.exe /S'"
}
'@
Add-CaseToFile -Path $paths.Local -CaseText $newLocalCase

# Central: bump caseB, add different new case
Edit-FileContent -Path $paths.Central `
    -Find  "`$compliantVer  = `"$caseBVer`"" `
    -Replace '$compliantVer  = "999.0.8000.100"'
$newCentralCase = @'

"AdminBApp" {
    $listPath      = "$listPathRoot\Test\AdminBApp.txt"
    $software      = "Admin B Application"
    $processName   = "adminbapp"
    $compliantVer  = "2.0.0"
    $patchPath     = "$patchRoot\Test\AdminBApp_2.0.0"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files\AdminBApp\app.exe"
    $installLine   = "& cmd /c 'C:\Temp\AdminBApp_2.0.0\setup.exe /S'"
}
'@
Add-CaseToFile -Path $paths.Central -CaseText $newCentralCase

$base    = & $parse -Path $paths.Base
$local   = & $parse -Path $paths.Local
$central = & $parse -Path $paths.Central
$comp    = & $compareFn -Local $local -Central $central -Base $base
$merge   = & $mergeFn -Local $local -Central $central -Comparison $comp -MergeDirection 'Pull'

$mergedContent = & $reconstruct -Parsed $merge.Merged

Assert-True ($mergedContent -match '999\.0\.3900\.50') "Mixed: Local $caseA bump preserved"
Assert-True ($mergedContent -match '999\.0\.8000\.100') "Mixed: Central $caseB bump pulled"
Assert-True ($merge.Merged.Cases.Contains('AdminAApp')) 'Mixed: Local new case preserved'
Assert-True ($merge.Merged.Cases.Contains('AdminBApp')) 'Mixed: Central new case pulled'
Assert-Equal 0 $merge.Stats.Conflicts 'Mixed: No conflicts (changes in different cases)'


# ============================================================================
#  TEST 27: Publish gate -- Receive-MainSwitch via public function (WhatIf)
# ============================================================================
Write-Host "`n--- TEST 27: Receive-MainSwitch -WhatIf ---" -ForegroundColor Cyan

$paths = New-TestEnvironment
# Create base file at expected location
$localBase = "$($paths.Local).base"
Copy-Item -Path $paths.Base -Destination $localBase

Edit-FileContent -Path $paths.Central `
    -Find  "`$compliantVer  = `"$caseAVer`"" `
    -Replace '$compliantVer  = "999.0.3900.50"'

# Run with -WhatIf so it does not actually write
Receive-MainSwitch -LocalPath $paths.Local -CentralPath $paths.Central -WhatIf 4>$null

# Verify the local file was NOT changed (WhatIf)
$localContent = Get-Content -Path $paths.Local -Encoding Default -Raw
Assert-True ($localContent -match $escapedCaseAVer) 'WhatIf: Local file unchanged'

# Clean up base
if (Test-Path $localBase) { Remove-Item $localBase }


# ============================================================================
#  TEST 28: Submit-MainSwitch gate blocks when central has diverged
# ============================================================================
Write-Host "`n--- TEST 28: Publish gate blocks diverged central ---" -ForegroundColor Cyan

$paths = New-TestEnvironment
$localBase = "$($paths.Local).base"
Copy-Item -Path $paths.Base -Destination $localBase

# Central diverges from base
Edit-FileContent -Path $paths.Central `
    -Find  "`$compliantVer  = `"$caseAVer`"" `
    -Replace '$compliantVer  = "999.0.0.0"'

# Capture output
$output = Submit-MainSwitch -LocalPath $paths.Local -CentralPath $paths.Central 4>&1 6>&1 | Out-String

Assert-True ($output -match 'PUBLISH BLOCKED|changed since your last sync') 'Publish gate blocked the push'

# Clean up base
if (Test-Path $localBase) { Remove-Item $localBase }


# ============================================================================
#  TEST 29: Submit-MainSwitch gate blocks when no base file exists
# ============================================================================
Write-Host "`n--- TEST 29: Publish gate blocks when no base ---" -ForegroundColor Cyan

$paths = New-TestEnvironment
# No base file created

$output = Submit-MainSwitch -LocalPath $paths.Local -CentralPath $paths.Central 4>&1 6>&1 | Out-String

Assert-True ($output -match 'No base file|Receive-MainSwitch first') 'Publish gate blocked without base'


# ============================================================================
#  TEST 30: Merged output is valid -- re-parseable
# ============================================================================
Write-Host "`n--- TEST 30: Merged output is re-parseable ---" -ForegroundColor Cyan

$paths = New-TestEnvironment

# Make several changes and merge
Edit-FileContent -Path $paths.Local `
    -Find  "`$compliantVer  = `"$caseAVer`"" `
    -Replace '$compliantVer  = "999.0.3900.50"'
$newCase = @'

"RoundTripApp" {
    $listPath      = "$listPathRoot\Test\RoundTripApp.txt"
    $software      = "Round Trip Application"
    $processName   = "roundtripapp"
    $compliantVer  = "5.0.0"
    $patchPath     = "$patchRoot\Test\RoundTripApp_5.0.0"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files\RoundTripApp\app.exe"
    $installLine   = @(
        "Set-Location C:\Temp\RoundTripApp_5.0.0;"
        "PowerShell.exe -ExecutionPolicy Bypass -File 'C:\Temp\RoundTripApp_5.0.0\install.ps1'"
    )
}
'@
Add-CaseToFile -Path $paths.Central -CaseText $newCase

$base    = & $parse -Path $paths.Base
$local   = & $parse -Path $paths.Local
$central = & $parse -Path $paths.Central
$comp    = & $compareFn -Local $local -Central $central -Base $base
$merge   = & $mergeFn -Local $local -Central $central -Comparison $comp -MergeDirection 'Pull'

$mergedContent = & $reconstruct -Parsed $merge.Merged
$mergedPath = Join-Path $testDir 'Main-Switch-merged.ps1'
Set-Content -Path $mergedPath -Value $mergedContent -Encoding Default -NoNewline

# Re-parse the merged file
$reParsed = & $parse -Path $mergedPath
Assert-True ($reParsed.Cases.Count -ge $baseCaseCount) "Re-parsed case count: $($reParsed.Cases.Count) (was $baseCaseCount)"
Assert-True ($reParsed.Cases.Contains('RoundTripApp')) 'Re-parsed contains new RoundTripApp'
Assert-True ($reParsed.Cases.Contains($caseA)) "Re-parsed still contains $caseA"

# Re-reconstruct and compare
$reReconstructed = & $reconstruct -Parsed $reParsed
$match = (& $normalize -Text $mergedContent) -eq (& $normalize -Text $reReconstructed)
Assert-True $match 'Merged -> parsed -> reconstructed matches (double round-trip)'


# ============================================================================
#  Cleanup and Summary
# ============================================================================
if (Test-Path $testDir) { Remove-Item $testDir -Recurse -Force }

Write-Host ''
Write-Host ('=' * 70) -ForegroundColor White
Write-Host "  Results: $($script:passed) passed, $($script:failed) failed out of $($script:testNum) assertions" -ForegroundColor $(if ($script:failed -eq 0) { 'Green' } else { 'Red' })
Write-Host ('=' * 70) -ForegroundColor White
Write-Host ''

if ($script:failed -gt 0) { exit 1 }
