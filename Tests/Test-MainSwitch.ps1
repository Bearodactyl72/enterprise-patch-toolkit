# DOTS formatting comment

<#
    .SYNOPSIS
        Validates that Main-Switch.ps1 parses cleanly and every switch entry
        returns a well-formed parameters hashtable.
    .DESCRIPTION
        Three categories of tests:

          Category 1 : Syntax & Parse Validation
            Test  1 : File parses without errors (catches loose statements in switch)
            Test  2 : File returns a hashtable when given a valid case name

          Category 2 : Per-Entry Smoke Tests (dynamic, one per case)
            Test  3+: Every switch entry returns a hashtable with required keys
                      and correct types

          Category 3 : Cross-Entry Validation
            Test N+1: No duplicate case names
            Test N+2: Uninstall-only entries (null $patchPath) have null or
                      empty $patchName

        Case names are discovered dynamically by regex, so adding or removing
        switch entries will not break existing tests.

        Estimated runtime: < 30 seconds.

    .NOTES
        Written by Skyler Werner
        Date: 2026/03/30
    .EXAMPLE
        .\Test-MainSwitch.ps1
#>

# --- Setup ---
$repoRoot       = Split-Path $PSScriptRoot -Parent
$mainSwitchPath = "$repoRoot\Scripts\Main-Switch.ps1"

# Counters
$script:passed  = 0
$script:failed  = 0
$script:skipped = 0
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

function Skip-Test {
    param([string]$Message)
    $script:testNum++
    $script:skipped++
    Write-Host "    SKIP: $Message" -ForegroundColor Yellow
}


# ============================================================
#  Dynamic Discovery -- extract case names from the file
# ============================================================

$fileContent = Get-Content $mainSwitchPath -Raw
$fileLines   = Get-Content $mainSwitchPath

# Match all case labels:  "CaseName" {
$casePattern = '(?m)^"(\w+)"\s*\{'
$caseMatches = [regex]::Matches($fileContent, $casePattern)
$allCaseNames = @($caseMatches | ForEach-Object { $_.Groups[1].Value })

Write-Host ""
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "  Test-MainSwitch.ps1" -ForegroundColor Cyan
Write-Host "  Discovered $($allCaseNames.Count) switch entries" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""


# ============================================================
#  Build the scriptblock once -- used for all invocations
# ============================================================

$mainSwitchScript = $null
$parseError = $null

try {
    $mainSwitchScript = (Get-Command $mainSwitchPath).ScriptBlock
} catch {
    $parseError = $_.Exception.Message
}


# ============================================================
#  Category 1: Syntax & Parse Validation
# ============================================================

Write-Host "--- Category 1: Syntax and Parse Validation ---" -ForegroundColor White

# Test 1: File parses without errors
Assert-True ($null -ne $mainSwitchScript) "Main-Switch.ps1 parses without errors"
if ($null -eq $mainSwitchScript) {
    Write-Host "          Parse error: $parseError" -ForegroundColor Red
    Write-Host ""
    Write-Host "FATAL: Cannot continue -- file failed to parse." -ForegroundColor Red
    Write-Host "  Passed: $script:passed  Failed: $script:failed  Skipped: $script:skipped" -ForegroundColor Yellow
    exit 1
}

# Test 2: Returns a hashtable for a known case
$testCase = $allCaseNames[0]
$result = $null
try {
    $targetSoftware = $testCase
    $scriptPath     = "$repoRoot\Scripts"
    $result = & $mainSwitchScript
} catch { }

Assert-True ($result -is [hashtable]) "Returns hashtable for case '$testCase'"

Write-Host ""


# ============================================================
#  Category 2: Per-Entry Smoke Tests
# ============================================================

Write-Host "--- Category 2: Per-Entry Smoke Tests ---" -ForegroundColor White

# Required keys that every entry must populate (non-null)
$requiredKeys = @('Software', 'ListPath', 'CompliantVer', 'PatchScript', 'SoftwarePaths', 'InstallTimeout')

# All keys in the parameters hashtable
$allKeys = @('Tag', 'Software', 'ListPath', 'CompliantVer', 'VersionType',
             'PatchPath', 'PatchName', 'ProcessName', 'PatchScript',
             'SoftwarePaths', 'RegistryKey', 'InstallLine', 'InstallTimeout', 'KB')

# Collect results for cross-entry validation
$allResults     = @{}
$timeoutResults = @{}

foreach ($caseName in $allCaseNames) {
    $result = $null
    $runError = $null

    try {
        $targetSoftware = $caseName
        $scriptPath     = "$repoRoot\Scripts"
        $result = & $mainSwitchScript
    } catch {
        $runError = $_.Exception.Message
    }

    # Test: Entry executes without throwing
    if ($null -ne $runError) {
        Assert-True $false "$caseName -- executes without error"
        Write-Host "          Error: $runError" -ForegroundColor Red
        continue
    }

    # Test: Returns a hashtable
    if ($result -isnot [hashtable]) {
        Assert-True $false "$caseName -- returns hashtable"
        continue
    }

    # Test: Hashtable has InstallTimeout key
    $timeout = $result['InstallTimeout']
    if ($null -eq $timeout) {
        Assert-True $false "$caseName -- has InstallTimeout"
        continue
    }

    # Test: Required keys are non-null
    $missingKeys = @()
    foreach ($key in $requiredKeys) {
        if ($null -eq $result[$key]) {
            $missingKeys += $key
        }
    }

    if ($missingKeys.Count -gt 0) {
        Assert-True $false "$caseName -- missing required keys: $($missingKeys -join ', ')"
    } else {
        # Single PASS line for the whole entry
        Assert-True $true "$caseName -- OK (timeout=$timeout)"
    }

    $allResults[$caseName]     = $result
    $timeoutResults[$caseName] = $timeout
}

Write-Host ""


# ============================================================
#  Category 3: Cross-Entry Validation
# ============================================================

Write-Host "--- Category 3: Cross-Entry Validation ---" -ForegroundColor White

# Test: No duplicate case names
$uniqueNames = @($allCaseNames | Select-Object -Unique)
$dupes = @($allCaseNames | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
Assert-True ($dupes.Count -eq 0) "No duplicate case names"
if ($dupes.Count -gt 0) {
    Write-Host "          Duplicates: $($dupes -join ', ')" -ForegroundColor Red
}

# Test: Uninstall-only entries (null patchPath) should not have a patchName
$uninstallOnly = @($allResults.GetEnumerator() | Where-Object {
    $null -eq $_.Value['PatchPath'] -and
    $null -ne $_.Value['PatchName'] -and
    $_.Value['PatchName'] -ne ''
})
Assert-True ($uninstallOnly.Count -eq 0) "Uninstall-only entries have no patchName"
if ($uninstallOnly.Count -gt 0) {
    foreach ($entry in $uninstallOnly) {
        Write-Host "          $($entry.Key) has patchName '$($entry.Value['PatchName'])' but null patchPath" -ForegroundColor Red
    }
}

# Test: PatchScript values are ScriptBlock type (not string)
$badScripts = @($allResults.GetEnumerator() | Where-Object {
    $null -ne $_.Value['PatchScript'] -and
    $_.Value['PatchScript'] -isnot [scriptblock]
})
Assert-True ($badScripts.Count -eq 0) "All PatchScript values are ScriptBlock type"
if ($badScripts.Count -gt 0) {
    foreach ($entry in $badScripts) {
        $actualType = $entry.Value['PatchScript'].GetType().Name
        Write-Host "          $($entry.Key) PatchScript is [$actualType], expected [ScriptBlock]" -ForegroundColor Red
    }
}

# Test: InstallTimeout is always an integer
$nonIntTimeouts = @($timeoutResults.GetEnumerator() | Where-Object {
    $_.Value -isnot [int]
})
Assert-True ($nonIntTimeouts.Count -eq 0) "All InstallTimeout values are integers"
if ($nonIntTimeouts.Count -gt 0) {
    foreach ($entry in $nonIntTimeouts) {
        $actualType = $entry.Value.GetType().Name
        Write-Host "          $($entry.Key) timeout is [$actualType], expected [int]" -ForegroundColor Red
    }
}


# ============================================================
#  Summary
# ============================================================

Write-Host ""
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "  Results" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "  Total:   $script:testNum" -ForegroundColor White
Write-Host "  Passed:  $script:passed" -ForegroundColor Green
Write-Host "  Failed:  $script:failed" -ForegroundColor $(if ($script:failed -gt 0) { 'Red' } else { 'Green' })
Write-Host "  Skipped: $script:skipped" -ForegroundColor Yellow
Write-Host ""

# Timeout distribution summary
Write-Host "--- InstallTimeout Distribution ---" -ForegroundColor White
$timeoutResults.Values | Group-Object | Sort-Object Name | ForEach-Object {
    Write-Host "    $($_.Name) min : $($_.Count) entries" -ForegroundColor Gray
}
Write-Host ""

if ($script:failed -gt 0) {
    Write-Host "RESULT: FAIL" -ForegroundColor Red
    exit 1
} else {
    Write-Host "RESULT: PASS" -ForegroundColor Green
    exit 0
}
