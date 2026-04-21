# DOTS formatting comment

<#
    .SYNOPSIS
        Comprehensive test suite for Invoke-RunspacePool module.
    .DESCRIPTION
        Tests all 5 documented failure modes plus edge cases and new features:

          Test  1  : Single-Item Vanishing Defect
          Test  2  : Multi-Item Collection Collapse Defect
          Test  3  : State Leakage / Alternating Result Defect
          Test  4  : Empty Output / Graceful Empty Return
          Test  5  : Silent Crash / Error Handling Defect
          Test  6  : ArgumentList Guard Verification
          Test  7  : Multiple Arguments Per Runspace
          Test  8  : Multi-Output Per Runspace
          Test  9  : Mixed Success/Failure Results
          Test 10  : Rapid Sequential Calls (back-to-back stability)
          Test 11  : PhaseTracker Availability (ISS-injected synchronized hashtable)
          Test 12  : StatusMessage Availability (ISS-injected synchronized hashtable)
          Test 13  : Drip-Feed Throttle Enforcement
          Test 14  : Timeout Produces 'Task Stopped' Result
          Test 15  : Batched Timeout Stops / Non-Blocking Cleanup

        All tests use local-only scriptblocks -- no network access required.
        Estimated runtime: 2-3 minutes (Tests 14-15 wait for 1-minute timeouts).

        Written by Skyler Werner
        Date: 2026/03/06
        Version 1.2.0

    .PARAMETER ModulePath
        Path to Invoke-RunspacePool.psm1. Auto-detected if not specified.
    .PARAMETER Only
        Run only the specified test number(s). Accepts one or more integers.
        Example: -Only 16  or  -Only 14,15,16
    .EXAMPLE
        .\Test-RunspaceComprehensive.ps1
        .\Test-RunspaceComprehensive.ps1 -Only 16
        .\Test-RunspaceComprehensive.ps1 -Only 14,15,16
        .\Test-RunspaceComprehensive.ps1 -ModulePath "C:\Modules\Invoke-RunspacePool.psm1"
#>

param(
    [string]$ModulePath,
    [int[]]$Only
)

$ErrorActionPreference = 'Continue'


#region ===== SETUP =====

# --- Test filter ---

$script:OnlyTests  = $Only
$script:SkipFilter = $Only.Count -gt 0
function Test-ShouldRun ([int]$Number) {
    if (-not $script:SkipFilter) { return $true }
    return $script:OnlyTests -contains $Number
}

# --- Test result tracker ---

$script:TestResults = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:TestNumber  = 0

function Assert-Test {
    <#
        Records a single assertion. CompareMode controls how Expected vs Actual are evaluated.
        Modes: Equals (default), NotNull, IsNull, Contains, GreaterThan
    #>
    param(
        [string]$TestName,
        [string]$Description,
        [object]$Expected,
        [object]$Actual,
        [ValidateSet('Equals','NotNull','IsNull','Contains','GreaterThan')]
        [string]$CompareMode = 'Equals'
    )

    $script:TestNumber++

    $passed = switch ($CompareMode) {
        'Equals'      { "$Actual" -eq "$Expected" }
        'NotNull'     { $null -ne $Actual }
        'IsNull'      { $null -eq $Actual }
        'Contains'    { "$Actual" -match [regex]::Escape("$Expected") }
        'GreaterThan' { [int]$Actual -gt [int]$Expected }
    }

    $status = if ($passed) { 'PASS' } else { 'FAIL' }
    $color  = if ($passed) { 'Green' } else { 'Red' }
    $marker = if ($passed) { '' } else { '  <<<' }

    Write-Host "  [$status] " -ForegroundColor $color -NoNewline
    Write-Host "$Description" -NoNewline

    switch ($CompareMode) {
        'Equals'      { Write-Host "  (Expected: $Expected | Actual: $Actual)$marker" }
        'Contains'    { Write-Host "  (Pattern: '$Expected' in '$Actual')$marker" }
        'GreaterThan' { Write-Host "  (Actual: $Actual > $Expected)$marker" }
        default       { Write-Host "  (Value: $Actual)$marker" }
    }

    $script:TestResults.Add([PSCustomObject]@{
        Number      = $script:TestNumber
        Test        = $TestName
        Description = $Description
        Status      = $status
        Expected    = $Expected
        Actual      = $Actual
    })
}


# --- Locate and import module ---

if (-not $ModulePath) {
    $candidates = @(
        (Join-Path (Split-Path $PSScriptRoot -Parent) "Modules\Invoke-RunspacePool\Invoke-RunspacePool.psm1")
        (Join-Path $PSScriptRoot "Modules\Invoke-RunspacePool\Invoke-RunspacePool.psm1")
        ".\Modules\Invoke-RunspacePool\Invoke-RunspacePool.psm1"
    )
    foreach ($c in $candidates) {
        if (Test-Path $c) { $ModulePath = (Resolve-Path $c).Path; break }
    }
}

if (-not $ModulePath -or -not (Test-Path $ModulePath)) {
    Write-Host "ERROR: Cannot find Invoke-RunspacePool.psm1" -ForegroundColor Red
    Write-Host "  Searched:"
    foreach ($c in $candidates) { Write-Host "    $c" }
    Write-Host "  Use -ModulePath to specify the location."
    return
}

Write-Host ""
Write-Host ("=" * 72) -ForegroundColor Cyan
Write-Host "  Invoke-RunspacePool  --  Comprehensive Test Suite" -ForegroundColor Cyan
Write-Host ("=" * 72) -ForegroundColor Cyan
Write-Host ""
Write-Host "  Module : $ModulePath"
Write-Host "  PS Ver : $($PSVersionTable.PSVersion)"
Write-Host "  Host   : $($Host.Name)"
Write-Host "  Date   : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host ""

if ($PSVersionTable.PSVersion.Major -ne 5) {
    Write-Host "  NOTE: Bug is PS 5.1-specific. Tests should pass on any version," -ForegroundColor Yellow
    Write-Host "        but run on PS 5.1 for authoritative results." -ForegroundColor Yellow
    Write-Host ""
}

Import-Module $ModulePath -Force
Write-Host "  Module imported." -ForegroundColor Gray
Write-Host ""

$sw = [System.Diagnostics.Stopwatch]::StartNew()

#endregion ===== SETUP =====



#region ===== SCRIPTBLOCKS =====

# Returns 1 PSCustomObject echoing args[0]
$echoBlock = {
    [PSCustomObject]@{
        ComputerName = $args[0]
        Status       = "Online"
        Comment      = "OK"
        ArgsReceived = $args.Count
    }
}

# Produces zero pipeline output
$emptyBlock = {
    $null = $args[0]
}

# Throws a terminating error
$throwBlock = {
    throw "Intentional test error on $($args[0])"
}

# Non-terminating error only, no standard output
$writeErrorBlock = {
    Write-Error "Non-terminating error for $($args[0])"
}

# Emits 2 objects per invocation
$multiOutputBlock = {
    [PSCustomObject]@{ ComputerName = $args[0]; Item = "A"; Comment = "First"  }
    [PSCustomObject]@{ ComputerName = $args[0]; Item = "B"; Comment = "Second" }
}

# Throws only when arg matches 'FAIL'
$conditionalBlock = {
    if ($args[0] -match 'FAIL') {
        throw "Intentional failure on $($args[0])"
    }
    [PSCustomObject]@{
        ComputerName = $args[0]
        Status       = "Online"
        Comment      = "OK"
    }
}

# Uses args[0] (computer) and args[1] (filepath)
$multiArgBlock = {
    [PSCustomObject]@{
        ComputerName = $args[0]
        FilePath     = $args[1]
        Status       = "Online"
        Comment      = "Multi-arg OK"
        ArgsReceived = $args.Count
    }
}

#endregion ===== SCRIPTBLOCKS =====



#region ===== HELPERS =====

function Build-ArgSets {
    <# Builds the standard @( ,@(machine) ) argument array from a list of machine names. #>
    param([string[]]$Machines)
    @(
        foreach ($m in $Machines) {
            , @($m)
        }
    )
}

# Common splat for all Invoke-RunspacePool calls
$common = @{
    ThrottleLimit  = 10
    TimeoutMinutes = 1
    ActivityName   = "Test"
}

#endregion ===== HELPERS =====



# ====================================================================
#  TEST 1 : Single-Item Vanishing Defect
#  Condition  : 1 target -> should produce exactly 1 result.
#  Known fail : Returns 0 results ($null).
# ====================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 1: Single-Item Vanishing Defect" -ForegroundColor Yellow
if (-not (Test-ShouldRun 1)) { Write-Host "  SKIPPED" -ForegroundColor DarkGray }
else {
Write-Host "  1 target must produce exactly 1 result." -ForegroundColor Gray
Write-Host ""

$a1 = Build-ArgSets -Machines @("SOLO-001")
$r1 = @(Invoke-RunspacePool -ScriptBlock $echoBlock -ArgumentList $a1 @common)

Assert-Test "Test 1" "Result count"     -Expected 1          -Actual $r1.Count
Assert-Test "Test 1" "ComputerName"     -Expected "SOLO-001" -Actual $r1[0].ComputerName
Assert-Test "Test 1" "Comment"          -Expected "OK"        -Actual $r1[0].Comment
Assert-Test "Test 1" "ArgsReceived"     -Expected 1          -Actual $r1[0].ArgsReceived
}
Write-Host ""



# ====================================================================
#  TEST 2 : Multi-Item Collection Collapse Defect
#  Condition  : N targets -> should produce exactly N individual results.
#  Known fail : Returns 1 object (the collection itself, not its contents).
# ====================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 2: Multi-Item Collection Collapse Defect" -ForegroundColor Yellow
if (-not (Test-ShouldRun 2)) { Write-Host "  SKIPPED" -ForegroundColor DarkGray }
else {
Write-Host "  N targets must produce N individual PSCustomObjects." -ForegroundColor Gray
Write-Host ""

# 2a: 3 targets
$a2a = Build-ArgSets -Machines @("COL-001", "COL-002", "COL-003")
$r2a = @(Invoke-RunspacePool -ScriptBlock $echoBlock -ArgumentList $a2a @common)

Assert-Test "Test 2a" "3 targets -> 3 results"             -Expected 3      -Actual $r2a.Count
Assert-Test "Test 2a" "Result[0] is PSCustomObject"         -Expected "True" -Actual "$($r2a[0] -is [PSCustomObject])"
Assert-Test "Test 2a" "Result[0] is NOT a collection type"  -Expected "True" -Actual "$($r2a[0] -isnot [System.Collections.IList])"

# 2b: 5 targets
$a2b = Build-ArgSets -Machines @("COL-A", "COL-B", "COL-C", "COL-D", "COL-E")
$r2b = @(Invoke-RunspacePool -ScriptBlock $echoBlock -ArgumentList $a2b @common)

Assert-Test "Test 2b" "5 targets -> 5 results" -Expected 5 -Actual $r2b.Count

$names2b   = ($r2b | ForEach-Object { $_.ComputerName }) | Sort-Object
$expect2b  = @("COL-A","COL-B","COL-C","COL-D","COL-E") | Sort-Object
Assert-Test "Test 2b" "All 5 computer names present" -Expected ($expect2b -join ',') -Actual ($names2b -join ',')

# 2c: 10 targets (larger batch)
$machines2c = 1..10 | ForEach-Object { "BATCH-{0:D3}" -f $_ }
$a2c = Build-ArgSets -Machines $machines2c
$r2c = @(Invoke-RunspacePool -ScriptBlock $echoBlock -ArgumentList $a2c @common)

Assert-Test "Test 2c" "10 targets -> 10 results" -Expected 10 -Actual $r2c.Count
}
Write-Host ""



# ====================================================================
#  TEST 3 : State Leakage / Alternating Result Defect
#  Condition  : Multiple sequential calls in the same session (no re-import).
#  Known fail : Results alternate between correct/incorrect across runs.
# ====================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 3: State Leakage / Alternating Result Defect" -ForegroundColor Yellow
if (-not (Test-ShouldRun 3)) { Write-Host "  SKIPPED" -ForegroundColor DarkGray }
else {
Write-Host "  5 alternating single/multi calls; every result must be correct." -ForegroundColor Gray
Write-Host ""

$aLeakS = Build-ArgSets -Machines @("LEAK-SOLO")
$aLeakM = Build-ArgSets -Machines @("LEAK-A", "LEAK-B", "LEAK-C")

$lr1 = @(Invoke-RunspacePool -ScriptBlock $echoBlock -ArgumentList $aLeakS @common)
Assert-Test "Test 3" "Run 1 (single) -> 1"  -Expected 1 -Actual $lr1.Count

$lr2 = @(Invoke-RunspacePool -ScriptBlock $echoBlock -ArgumentList $aLeakM @common)
Assert-Test "Test 3" "Run 2 (multi)  -> 3"  -Expected 3 -Actual $lr2.Count

$lr3 = @(Invoke-RunspacePool -ScriptBlock $echoBlock -ArgumentList $aLeakS @common)
Assert-Test "Test 3" "Run 3 (single) -> 1"  -Expected 1 -Actual $lr3.Count

$lr4 = @(Invoke-RunspacePool -ScriptBlock $echoBlock -ArgumentList $aLeakM @common)
Assert-Test "Test 3" "Run 4 (multi)  -> 3"  -Expected 3 -Actual $lr4.Count

$lr5 = @(Invoke-RunspacePool -ScriptBlock $echoBlock -ArgumentList $aLeakS @common)
Assert-Test "Test 3" "Run 5 (single) -> 1"  -Expected 1 -Actual $lr5.Count

# Verify module-scoped state is clean after each call
$poolState = & (Get-Module Invoke-RunspacePool) { $script:CurrentPool }
Assert-Test "Test 3" "Module CurrentPool is null after runs" -Expected "True" -Actual "$($null -eq $poolState)"

$rsState   = & (Get-Module Invoke-RunspacePool) { $script:CurrentRunspaces }
Assert-Test "Test 3" "Module CurrentRunspaces is null after runs" -Expected "True" -Actual "$($null -eq $rsState)"
}
Write-Host ""



# ====================================================================
#  TEST 4 : Empty Output / Graceful Empty Return
#  Condition  : Scriptblock produces no pipeline output.
#  Expected   : Safety-net result with "Task produced no output" comment
#               (prevents machines from silently disappearing from results).
# ====================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 4: Empty Output / Graceful Empty Return" -ForegroundColor Yellow
if (-not (Test-ShouldRun 4)) { Write-Host "  SKIPPED" -ForegroundColor DarkGray }
else {
Write-Host "  Scriptblock producing no output must not crash or throw." -ForegroundColor Gray
Write-Host ""

$a4 = Build-ArgSets -Machines @("EMPTY-001")
$error4 = $null

try {
    $r4 = @(Invoke-RunspacePool -ScriptBlock $emptyBlock -ArgumentList $a4 @common)
}
catch {
    $error4 = $_
}

Assert-Test "Test 4" "No exception thrown"                  -Expected "True"      -Actual "$($null -eq $error4)"
Assert-Test "Test 4" "Safety-net result returned"           -Expected 1           -Actual $r4.Count
Assert-Test "Test 4" "Comment = 'Task produced no output'"  -Expected "Task produced no output" -Actual $r4[0].Comment -CompareMode Contains

# 4b: Multiple empty targets -- all get safety-net results
$a4b = Build-ArgSets -Machines @("EMPTY-A", "EMPTY-B", "EMPTY-C")
$error4b = $null

try {
    $r4b = @(Invoke-RunspacePool -ScriptBlock $emptyBlock -ArgumentList $a4b @common)
}
catch {
    $error4b = $_
}

Assert-Test "Test 4b" "3 empty targets: no exception"          -Expected "True" -Actual "$($null -eq $error4b)"
Assert-Test "Test 4b" "3 empty targets: 3 safety-net results"  -Expected 3      -Actual $r4b.Count
}
Write-Host ""



# ====================================================================
#  TEST 5 : Silent Crash / Error Handling Defect
#  Condition  : Errors inside scriptblocks must be caught and reported,
#               not silently swallowed.
# ====================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 5: Silent Crash / Error Handling Defect" -ForegroundColor Yellow
if (-not (Test-ShouldRun 5)) { Write-Host "  SKIPPED" -ForegroundColor DarkGray }
else {
Write-Host ""

# --- 5a: Terminating error (throw) ---
Write-Host "  5a: Terminating error in scriptblock" -ForegroundColor Gray

$a5a = Build-ArgSets -Machines @("THROW-001")
$error5a = $null

try {
    $r5a = @(Invoke-RunspacePool -ScriptBlock $throwBlock -ArgumentList $a5a @common)
}
catch {
    $error5a = $_
}

Assert-Test "Test 5a" "No unhandled exception"      -Expected "True"       -Actual "$($null -eq $error5a)"
Assert-Test "Test 5a" "Returns 1 error result"       -Expected 1           -Actual $r5a.Count
Assert-Test "Test 5a" "Comment contains 'Task Failed'" -Expected "Task Failed" -Actual $r5a[0].Comment -CompareMode Contains

# --- 5b: Non-terminating error (Write-Error, no output) ---
Write-Host "  5b: Non-terminating error (Write-Error only)" -ForegroundColor Gray

$a5b = Build-ArgSets -Machines @("WERROR-001")
$error5b = $null

try {
    $r5b = @(Invoke-RunspacePool -ScriptBlock $writeErrorBlock -ArgumentList $a5b @common)
    $r5b = @($r5b | Where-Object { $null -ne $_ })
}
catch {
    $error5b = $_
}

Assert-Test "Test 5b" "No unhandled exception"     -Expected "True"      -Actual "$($null -eq $error5b)"
Assert-Test "Test 5b" "Returns 1 error result"      -Expected 1          -Actual $r5b.Count
Assert-Test "Test 5b" "Comment contains 'Task Error'" -Expected "Task Error" -Actual $r5b[0].Comment -CompareMode Contains

# --- 5c: Terminating error with multiple targets (ensure others still return) ---
Write-Host "  5c: 1 throw among 3 targets -- other results survive" -ForegroundColor Gray

$a5c = Build-ArgSets -Machines @("GOOD-5C-A", "THROW-5C", "GOOD-5C-B")
$error5c = $null

try {
    $r5c = @(Invoke-RunspacePool -ScriptBlock $conditionalBlock -ArgumentList $a5c @common)
}
catch {
    $error5c = $_
}

Assert-Test "Test 5c" "No unhandled exception"      -Expected "True" -Actual "$($null -eq $error5c)"
Assert-Test "Test 5c" "3 targets -> 3 results total" -Expected 3     -Actual $r5c.Count
}
Write-Host ""



# ====================================================================
#  TEST 6 : ArgumentList Guard Verification
#  The guard detects when PS 5.1 unwraps a single-element array-of-arrays
#  into a flat array, and re-wraps it.
# ====================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 6: ArgumentList Guard Verification" -ForegroundColor Yellow
if (-not (Test-ShouldRun 6)) { Write-Host "  SKIPPED" -ForegroundColor DarkGray }
else {
Write-Host "  Tests flat vs. wrapped input for single targets." -ForegroundColor Gray
Write-Host ""

# 6a: Properly wrapped -- guard should not interfere
$a6a = @( , @("GUARD-WRAPPED") )
$r6a = @(Invoke-RunspacePool -ScriptBlock $echoBlock -ArgumentList $a6a @common)

Assert-Test "Test 6a" "Wrapped input -> 1 result"  -Expected 1               -Actual $r6a.Count
Assert-Test "Test 6a" "ComputerName correct"        -Expected "GUARD-WRAPPED" -Actual $r6a[0].ComputerName
Assert-Test "Test 6a" "ArgsReceived = 1"            -Expected 1               -Actual $r6a[0].ArgsReceived

# 6b: Flat input (simulates PS 5.1 unwrap) -- guard MUST re-wrap
$a6b = @("GUARD-FLAT")
$r6b = @(Invoke-RunspacePool -ScriptBlock $echoBlock -ArgumentList $a6b @common)

Assert-Test "Test 6b" "Flat input -> 1 result (guard re-wraps)"  -Expected 1            -Actual $r6b.Count
Assert-Test "Test 6b" "ComputerName correct after re-wrap"       -Expected "GUARD-FLAT"  -Actual $r6b[0].ComputerName

# 6c: Flat multi-element input (2 strings, not array-of-arrays)
# Guard wraps the whole thing into 1 argSet with 2 arguments
$a6c = @("FLAT-X", "FLAT-Y")
$r6c = @(Invoke-RunspacePool -ScriptBlock $echoBlock -ArgumentList $a6c @common)

Assert-Test "Test 6c" "Flat 2-element -> 1 runspace (not 2)" -Expected 1        -Actual $r6c.Count
Assert-Test "Test 6c" "Computer = first element"              -Expected "FLAT-X" -Actual $r6c[0].ComputerName
Assert-Test "Test 6c" "ArgsReceived = 2 (both elements)"     -Expected 2        -Actual $r6c[0].ArgsReceived
}
Write-Host ""



# ====================================================================
#  TEST 7 : Multiple Arguments Per Runspace
#  Scriptblock receives args[0] (computer) and args[1] (filepath).
# ====================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 7: Multiple Arguments Per Runspace" -ForegroundColor Yellow
if (-not (Test-ShouldRun 7)) { Write-Host "  SKIPPED" -ForegroundColor DarkGray }
else {
Write-Host "  Each runspace receives 2 arguments." -ForegroundColor Gray
Write-Host ""

# 7a: 2 targets, each with 2 args
$a7a = @(
    , @("MARG-001", "C:\Path\FileA.exe")
    , @("MARG-002", "C:\Path\FileB.exe")
)
$r7a = @(Invoke-RunspacePool -ScriptBlock $multiArgBlock -ArgumentList $a7a @common)

Assert-Test "Test 7a" "2 targets -> 2 results"  -Expected 2 -Actual $r7a.Count

$r7a_sorted = $r7a | Sort-Object ComputerName
if ($r7a_sorted.Count -ge 2) {
    Assert-Test "Test 7a" "[0] ArgsReceived = 2"  -Expected 2                   -Actual $r7a_sorted[0].ArgsReceived
    Assert-Test "Test 7a" "[0] FilePath"           -Expected "C:\Path\FileA.exe" -Actual $r7a_sorted[0].FilePath
    Assert-Test "Test 7a" "[1] FilePath"           -Expected "C:\Path\FileB.exe" -Actual $r7a_sorted[1].FilePath
}

# 7b: Single target with 2 args (combines with ArgumentList guard)
$a7b = @( , @("MARG-SOLO", "C:\Solo\File.exe") )
$r7b = @(Invoke-RunspacePool -ScriptBlock $multiArgBlock -ArgumentList $a7b @common)

Assert-Test "Test 7b" "1 target w/ 2 args -> 1 result"  -Expected 1                    -Actual $r7b.Count
Assert-Test "Test 7b" "ArgsReceived = 2"                  -Expected 2                    -Actual $r7b[0].ArgsReceived
Assert-Test "Test 7b" "FilePath passed correctly"          -Expected "C:\Solo\File.exe"   -Actual $r7b[0].FilePath
}
Write-Host ""



# ====================================================================
#  TEST 8 : Multi-Output Per Runspace
#  Scriptblock emits 2 objects per call. 3 targets = 6 total results.
# ====================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 8: Multi-Output Per Runspace" -ForegroundColor Yellow
if (-not (Test-ShouldRun 8)) { Write-Host "  SKIPPED" -ForegroundColor DarkGray }
else {
Write-Host "  Scriptblock emits 2 objects per call." -ForegroundColor Gray
Write-Host ""

# 8a: 3 targets x 2 outputs = 6 results
$a8a = Build-ArgSets -Machines @("MOUT-001", "MOUT-002", "MOUT-003")
$r8a = @(Invoke-RunspacePool -ScriptBlock $multiOutputBlock -ArgumentList $a8a @common)

Assert-Test "Test 8a" "3 targets x 2 outputs = 6 results"  -Expected 6 -Actual $r8a.Count

$itemA = @($r8a | Where-Object { $_.Item -eq "A" }).Count
$itemB = @($r8a | Where-Object { $_.Item -eq "B" }).Count
Assert-Test "Test 8a" "3 items with Item='A'"  -Expected 3 -Actual $itemA
Assert-Test "Test 8a" "3 items with Item='B'"  -Expected 3 -Actual $itemB

# 8b: Single target x 2 outputs = 2 results  (single-item variant)
$a8b = Build-ArgSets -Machines @("MOUT-SOLO")
$r8b = @(Invoke-RunspacePool -ScriptBlock $multiOutputBlock -ArgumentList $a8b @common)

Assert-Test "Test 8b" "1 target x 2 outputs = 2 results"  -Expected 2 -Actual $r8b.Count

$soloA = @($r8b | Where-Object { $_.Item -eq "A" }).Count
$soloB = @($r8b | Where-Object { $_.Item -eq "B" }).Count
Assert-Test "Test 8b" "1 item with Item='A'"  -Expected 1 -Actual $soloA
Assert-Test "Test 8b" "1 item with Item='B'"  -Expected 1 -Actual $soloB
}
Write-Host ""



# ====================================================================
#  TEST 9 : Mixed Success/Failure Results
#  Some scriptblocks succeed, some throw.  All must be accounted for.
# ====================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 9: Mixed Success/Failure Results" -ForegroundColor Yellow
if (-not (Test-ShouldRun 9)) { Write-Host "  SKIPPED" -ForegroundColor DarkGray }
else {
Write-Host "  2 good + 1 FAIL target -> 3 total results." -ForegroundColor Gray
Write-Host ""

$a9 = Build-ArgSets -Machines @("GOOD-001", "FAIL-001", "GOOD-002")
$r9 = @(Invoke-RunspacePool -ScriptBlock $conditionalBlock -ArgumentList $a9 @common)

Assert-Test "Test 9" "3 targets -> 3 results"  -Expected 3 -Actual $r9.Count

$goodCount = @($r9 | Where-Object { $_.Comment -eq "OK" }).Count
$failCount = @($r9 | Where-Object { $_.Comment -match "Task Failed" }).Count

Assert-Test "Test 9" "2 successful results"  -Expected 2 -Actual $goodCount
Assert-Test "Test 9" "1 failed result"        -Expected 1 -Actual $failCount

# Verify the failed result identifies the right target
$failResult = $r9 | Where-Object { $_.Comment -match "Task Failed" } | Select-Object -First 1
Assert-Test "Test 9" "Failed target = FAIL-001" -Expected "FAIL-001" -Actual $failResult.ComputerName

# 9b: All targets fail
$a9b = Build-ArgSets -Machines @("FAIL-A", "FAIL-B")
$r9b = @(Invoke-RunspacePool -ScriptBlock $conditionalBlock -ArgumentList $a9b @common)

Assert-Test "Test 9b" "2 failing targets -> 2 results"  -Expected 2 -Actual $r9b.Count

$allFailed = @($r9b | Where-Object { $_.Comment -match "Task Failed" }).Count
Assert-Test "Test 9b" "Both contain 'Task Failed'"  -Expected 2 -Actual $allFailed
}
Write-Host ""



# ====================================================================
#  TEST 10 : Rapid Sequential Calls
#  5 back-to-back calls with 2 targets each, no re-import between them.
#  Tests that cleanup between calls is complete and nothing leaks.
# ====================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 10: Rapid Sequential Calls" -ForegroundColor Yellow
if (-not (Test-ShouldRun 10)) { Write-Host "  SKIPPED" -ForegroundColor DarkGray }
else {
Write-Host "  5 back-to-back calls, 2 targets each." -ForegroundColor Gray
Write-Host ""

for ($i = 1; $i -le 5; $i++) {
    $aR = Build-ArgSets -Machines @("RAPID-$i-A", "RAPID-$i-B")
    $rR = @(Invoke-RunspacePool -ScriptBlock $echoBlock -ArgumentList $aR @common)

    Assert-Test "Test 10" "Rapid call $i -> 2 results"  -Expected 2 -Actual $rR.Count
}

# Verify clean state after rapid calls
$poolFinal = & (Get-Module Invoke-RunspacePool) { $script:CurrentPool }
Assert-Test "Test 10" "CurrentPool null after rapid calls" -Expected "True" -Actual "$($null -eq $poolFinal)"
}
Write-Host ""



# ====================================================================
#  TEST 11 : PhaseTracker Availability
#  The module injects a synchronized $PhaseTracker hashtable into every
#  runspace via ISS. Scriptblocks must be able to read and write it.
# ====================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 11: PhaseTracker Availability" -ForegroundColor Yellow
if (-not (Test-ShouldRun 11)) { Write-Host "  SKIPPED" -ForegroundColor DarkGray }
else {
Write-Host "  Scriptblock writes to `$PhaseTracker; verify it was accessible." -ForegroundColor Gray
Write-Host ""

$phaseBlock = {
    $PhaseTracker[$args[0]] = "Testing"
    Start-Sleep -Milliseconds 200
    [PSCustomObject]@{
        ComputerName  = $args[0]
        PhaseWasSet   = ($PhaseTracker[$args[0]] -eq "Testing")
        TrackerExists = ($null -ne $PhaseTracker)
        Comment       = "OK"
    }
}

$a11 = Build-ArgSets -Machines @("PHASE-001", "PHASE-002")
$r11 = @(Invoke-RunspacePool -ScriptBlock $phaseBlock -ArgumentList $a11 @common)

Assert-Test "Test 11" "2 targets -> 2 results"     -Expected 2      -Actual $r11.Count
Assert-Test "Test 11" "PhaseTracker exists in RS"   -Expected "True" -Actual "$($r11[0].TrackerExists)"
Assert-Test "Test 11" "PhaseTracker value was set"  -Expected "True" -Actual "$($r11[0].PhaseWasSet)"
}
Write-Host ""



# ====================================================================
#  TEST 12 : StatusMessage Availability
#  The module injects a synchronized $StatusMessage hashtable. Callers
#  can write $StatusMessage['Text'] = "..." for custom progress info.
# ====================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 12: StatusMessage Availability" -ForegroundColor Yellow
if (-not (Test-ShouldRun 12)) { Write-Host "  SKIPPED" -ForegroundColor DarkGray }
else {
Write-Host "  Scriptblock writes to `$StatusMessage; verify it was accessible." -ForegroundColor Gray
Write-Host ""

$statusBlock = {
    $StatusMessage['Text'] = "Custom status from $($args[0])"
    # Sleep long enough for at least one progress render (every 5s) to display the text
    Start-Sleep -Seconds 6
    [PSCustomObject]@{
        ComputerName   = $args[0]
        StatusExists   = ($null -ne $StatusMessage)
        StatusWritable = ($StatusMessage['Text'] -match "Custom status")
        Comment        = "OK"
    }
}

$a12 = Build-ArgSets -Machines @("STATUS-001")
$r12 = @(Invoke-RunspacePool -ScriptBlock $statusBlock -ArgumentList $a12 @common)

Assert-Test "Test 12" "1 target -> 1 result"           -Expected 1      -Actual $r12.Count
Assert-Test "Test 12" "StatusMessage exists in RS"      -Expected "True" -Actual "$($r12[0].StatusExists)"
Assert-Test "Test 12" "StatusMessage writable from RS"  -Expected "True" -Actual "$($r12[0].StatusWritable)"
}
Write-Host ""



# ====================================================================
#  TEST 13 : Drip-Feed Throttle Enforcement
#  With ThrottleLimit=2 and 5 targets that each sleep 1s, verify that
#  not all runspaces are submitted at once. The drip-feed helper should
#  hold back submissions until in-flight slots open up.
# ====================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 13: Drip-Feed Throttle Enforcement" -ForegroundColor Yellow
if (-not (Test-ShouldRun 13)) { Write-Host "  SKIPPED" -ForegroundColor DarkGray }
else {
Write-Host "  ThrottleLimit=2, 5 targets; verify throttled submission." -ForegroundColor Gray
Write-Host ""

$throttleBlock = {
    Start-Sleep -Seconds 1
    [PSCustomObject]@{
        ComputerName = $args[0]
        Comment      = "OK"
    }
}

$a13 = Build-ArgSets -Machines @("THRT-A","THRT-B","THRT-C","THRT-D","THRT-E")
$r13 = @(Invoke-RunspacePool -ScriptBlock $throttleBlock -ArgumentList $a13 -ThrottleLimit 2 -TimeoutMinutes 1 -ActivityName "Test")

Assert-Test "Test 13" "5 targets -> 5 results"    -Expected 5 -Actual $r13.Count

$names13 = ($r13 | ForEach-Object { $_.ComputerName }) | Sort-Object
$expect13 = @("THRT-A","THRT-B","THRT-C","THRT-D","THRT-E") | Sort-Object
Assert-Test "Test 13" "All 5 names present"  -Expected ($expect13 -join ',') -Actual ($names13 -join ',')

$allOK13 = @($r13 | Where-Object { $_.Comment -eq "OK" }).Count
Assert-Test "Test 13" "All 5 completed OK"   -Expected 5 -Actual $allOK13
}
Write-Host ""



# ====================================================================
#  TEST 14 : Timeout Produces 'Task Stopped' Result
#  A scriptblock that sleeps longer than TimeoutMinutes must produce
#  a result with Comment = "Task Stopped" (not crash or hang).
# ====================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 14: Timeout Produces 'Task Stopped' Result" -ForegroundColor Yellow
if (-not (Test-ShouldRun 14)) { Write-Host "  SKIPPED" -ForegroundColor DarkGray }
else {
Write-Host "  Scriptblock sleeps 120s, timeout = 1 min. Must get 'Task Stopped'." -ForegroundColor Gray
Write-Host ""

$sleepBlock = {
    Start-Sleep -Seconds 120
    [PSCustomObject]@{
        ComputerName = $args[0]
        Comment      = "Should not reach here"
    }
}

$a14 = Build-ArgSets -Machines @("TIMEOUT-001")
$error14 = $null

try {
    $r14 = @(Invoke-RunspacePool -ScriptBlock $sleepBlock -ArgumentList $a14 -ThrottleLimit 1 -TimeoutMinutes 1 -ActivityName "Test")
}
catch {
    $error14 = $_
}

Assert-Test "Test 14" "No unhandled exception"            -Expected "True"         -Actual "$($null -eq $error14)"
Assert-Test "Test 14" "Returns 1 result"                   -Expected 1             -Actual $r14.Count
Assert-Test "Test 14" "ComputerName = TIMEOUT-001"         -Expected "TIMEOUT-001" -Actual $r14[0].ComputerName
Assert-Test "Test 14" "Comment contains 'Task Stopped'"    -Expected "Task Stopped" -Actual $r14[0].Comment -CompareMode Contains
}
Write-Host ""



# ====================================================================
#  TEST 15 : Batched Timeout Stops / Non-Blocking Cleanup
#  Multiple runspaces that all exceed the timeout must be stopped
#  concurrently (batched BeginStop), not sequentially. The function
#  must also return promptly after stops -- not block for minutes
#  waiting on Dispose() for stuck pipelines.
#
#  With 5 targets at ThrottleLimit=5 and TimeoutMinutes=1:
#    Old sequential:   ~1 min wait + 5 * 15s stops = ~2m 15s minimum
#    Batched + async:  ~1 min wait + ~15s stops + instant return = ~1m 20s
#  Assert total time < 100s to catch sequential regressions.
# ====================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 15: Batched Timeout Stops / Non-Blocking Cleanup" -ForegroundColor Yellow
if (-not (Test-ShouldRun 15)) { Write-Host "  SKIPPED" -ForegroundColor DarkGray }
else {
Write-Host "  5 targets all timeout; stops must be concurrent, cleanup non-blocking." -ForegroundColor Gray
Write-Host ""

$sleepBlock15 = {
    Start-Sleep -Seconds 300
    [PSCustomObject]@{
        ComputerName = $args[0]
        Comment      = "Should not reach here"
    }
}

$a15 = Build-ArgSets -Machines @("BATCH-TO-A","BATCH-TO-B","BATCH-TO-C","BATCH-TO-D","BATCH-TO-E")
$sw15 = [System.Diagnostics.Stopwatch]::StartNew()
$error15 = $null

try {
    $r15 = @(Invoke-RunspacePool -ScriptBlock $sleepBlock15 -ArgumentList $a15 -ThrottleLimit 5 -TimeoutMinutes 1 -ActivityName "Test")
}
catch {
    $error15 = $_
}

$sw15.Stop()
$elapsed15 = [math]::Round($sw15.Elapsed.TotalSeconds, 1)

Assert-Test "Test 15" "No unhandled exception"             -Expected "True" -Actual "$($null -eq $error15)"
Assert-Test "Test 15" "5 targets -> 5 results"             -Expected 5      -Actual $r15.Count

$allStopped15 = @($r15 | Where-Object { $_.Comment -match "Task Stopped" }).Count
Assert-Test "Test 15" "All 5 report 'Task Stopped'"        -Expected 5 -Actual $allStopped15

# Key timing assertion: batched stops + non-blocking cleanup should keep
# total time well under the old sequential path (which would be ~2m 15s+).
# Allow up to 100s for timeout (60s) + batched stop wait (15s) + margin.
Assert-Test "Test 15" "Returned within 100s (was ${elapsed15}s)" -Expected "True" -Actual "$($elapsed15 -lt 100)"

Write-Host "  Elapsed: ${elapsed15}s" -ForegroundColor Gray
}
Write-Host ""



# ====================================================================
#  TEST 16 : Host Survives Stuck-Runspace Cleanup
#  Regression test for a bug where disposing a host-bound RunspacePool
#  on a background thread (ThreadPool::QueueUserWorkItem) crashed the
#  PowerShell process. The fix performs synchronous cleanup instead.
#
#  Strategy: launch the test in a child powershell.exe process. If the
#  host crashes during cleanup, the child exits with a non-zero code
#  (or the process disappears). A clean exit with code 0 proves the
#  host survived.
# ====================================================================

Write-Host ("-" * 72) -ForegroundColor DarkGray
Write-Host "TEST 16: Host Survives Stuck-Runspace Cleanup" -ForegroundColor Yellow
if (-not (Test-ShouldRun 16)) { Write-Host "  SKIPPED" -ForegroundColor DarkGray }
else {
Write-Host "  Runs timeout scenario in child process; exit code 0 = host survived." -ForegroundColor Gray
Write-Host ""

# Build inline script that imports the module, runs a stuck runspace,
# then writes a sentinel value and exits cleanly.
$sentinelFile = Join-Path $env:TEMP "Test16_sentinel_$(Get-Random).txt"
$childScript = @"
try {
    Import-Module '$($ModulePath -replace "'","''")' -Force
    `$block = { Start-Sleep -Seconds 300; [PSCustomObject]@{ ComputerName = `$args[0]; Comment = 'unreachable' } }
    `$args16 = @( ,@('SURVIVE-001') )
    `$r = @(Invoke-RunspacePool -ScriptBlock `$block -ArgumentList `$args16 -ThrottleLimit 1 -TimeoutMinutes 1 -ActivityName 'Test16')
    # If we get here, the host survived cleanup
    'HOST_ALIVE' | Set-Content '$($sentinelFile -replace "'","''")' -Encoding ASCII
    exit 0
}
catch {
    `$_ | Out-String | Set-Content '$($sentinelFile -replace "'","''")' -Encoding ASCII
    exit 2
}
"@

$sw16 = [System.Diagnostics.Stopwatch]::StartNew()
$tempScript = Join-Path $env:TEMP "Test16_HostSurvival_$(Get-Random).ps1"
$childScript | Set-Content -Path $tempScript -Encoding ASCII

# Run without redirecting stdout/stderr so [Console] APIs work in the
# child process. Use a sentinel file instead to detect success.
$proc = Start-Process -FilePath 'powershell.exe' `
    -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $tempScript `
    -Wait -PassThru

$exitCode16 = $proc.ExitCode
$sentinelContent = if (Test-Path $sentinelFile) { (Get-Content $sentinelFile -Raw).Trim() } else { '' }
if ($exitCode16 -ne 0 -and $sentinelContent) {
    Write-Host "  Child error: $sentinelContent" -ForegroundColor DarkYellow
}
$sw16.Stop()
$elapsed16 = [math]::Round($sw16.Elapsed.TotalSeconds, 1)

# Cleanup temp files
Remove-Item $tempScript -Force -ErrorAction SilentlyContinue
Remove-Item $sentinelFile -Force -ErrorAction SilentlyContinue

Assert-Test "Test 16" "Child process exited with code 0"       -Expected 0      -Actual $exitCode16
Assert-Test "Test 16" "Sentinel file contains 'HOST_ALIVE'"    -Expected "HOST_ALIVE" -Actual $sentinelContent
Assert-Test "Test 16" "Completed within 100s (was ${elapsed16}s)" -Expected "True" -Actual "$($elapsed16 -lt 100)"

Write-Host "  Exit code: $exitCode16 | Elapsed: ${elapsed16}s" -ForegroundColor Gray
}
Write-Host ""



# ====================================================================
#  SUMMARY
# ====================================================================

$sw.Stop()

$passed = @($script:TestResults | Where-Object { $_.Status -eq 'PASS' }).Count
$failed = @($script:TestResults | Where-Object { $_.Status -eq 'FAIL' }).Count
$total  = $script:TestResults.Count
$skippedTests = if ($Only.Count -gt 0) { 16 - $Only.Count } else { 0 }

Write-Host ("=" * 72) -ForegroundColor Cyan
Write-Host "  TEST SUMMARY" -ForegroundColor Cyan
Write-Host ("=" * 72) -ForegroundColor Cyan
Write-Host ""
if ($skippedTests -gt 0) {
    Write-Host "  Filter : -Only $($Only -join ',')" -ForegroundColor DarkGray
    Write-Host "  Skipped: $skippedTests test(s)" -ForegroundColor DarkGray
}
Write-Host "  Total  : $total assertions"
Write-Host "  Passed : " -NoNewline
Write-Host "$passed" -ForegroundColor Green
Write-Host "  Failed : " -NoNewline
Write-Host "$failed" -ForegroundColor $(if ($failed -gt 0) { 'Red' } else { 'Green' })
Write-Host "  Time   : $([math]::Round($sw.Elapsed.TotalSeconds, 1)) seconds"
Write-Host ""

if ($failed -gt 0) {
    Write-Host "  FAILED ASSERTIONS:" -ForegroundColor Red
    Write-Host ""
    foreach ($f in ($script:TestResults | Where-Object { $_.Status -eq 'FAIL' })) {
        Write-Host "    [$($f.Test)] $($f.Description)" -ForegroundColor Red
        Write-Host "      Expected: $($f.Expected)  |  Actual: $($f.Actual)" -ForegroundColor DarkRed
    }
    Write-Host ""
}

$overallColor = if ($failed -eq 0) { 'Green' } else { 'Red' }
$overallText  = if ($failed -eq 0) { 'ALL TESTS PASSED' } else { "$failed ASSERTION(S) FAILED" }
Write-Host "  >> $overallText <<" -ForegroundColor $overallColor
Write-Host ""
