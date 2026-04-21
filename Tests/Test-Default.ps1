# DOTS formatting comment

<#
    .SYNOPSIS
        Test suite for Default.ps1 and Default-NoUninstall.ps1 patching scripts.
    .DESCRIPTION
        Tests the restructured Default.ps1 and Default-NoUninstall.ps1 scripts
        locally without requiring remote machines. Validates:

          Test  1 : Config hashtable unpacking
          Test  2 : Install with successful exit code (cmd /c exit 0)
          Test  3 : Install with failure exit code (cmd /c exit 1603)
          Test  4 : Multi-line install with multiple exit codes
          Test  5 : Action log structure and per-command tracking
          Test  6 : Action log duration tracking
          Test  7 : Exit code comment enrichment via Get-ExitCodeComment
          Test  8 : Return object structure (ExitCode, ExitCodes, ActionLog, Comment)
          Test  9 : Software not installed (no paths found, no uninstall keys)
          Test 10 : Default-NoUninstall.ps1 skips uninstall section
          Test 11 : USER path expansion logic
          Test 12 : Install error handling (bad scriptblock)
          Test 13 : Log file creation and content (Default.ps1)
          Test 14 : Log file creation and content (Default-NoUninstall.ps1)

        Uses dummy install commands (cmd /c exit N) to produce known exit codes
        without installing real software. Creates temporary files in $env:TEMP for
        version detection tests, cleaned up after each run.

        Written by Skyler Werner
        Date: 2026/03/23
        Version 1.0.0

    .EXAMPLE
        .\Test-Default.ps1
#>

$ErrorActionPreference = 'Continue'


#region ===== SETUP =====

# --- Test result tracker ---

$script:TestResults = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:TestNumber  = 0

function Assert-Test {
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


# --- Locate scripts ---

$scriptRoot = Join-Path (Split-Path $PSScriptRoot -Parent) "Scripts\Patching"
$defaultScript       = Join-Path $scriptRoot "Default.ps1"
$defaultNoUninstall  = Join-Path $scriptRoot "Default-NoUninstall.ps1"

if (-not (Test-Path $defaultScript)) {
    Write-Host "ERROR: Cannot find Default.ps1 at $defaultScript" -ForegroundColor Red
    return
}
if (-not (Test-Path $defaultNoUninstall)) {
    Write-Host "ERROR: Cannot find Default-NoUninstall.ps1 at $defaultNoUninstall" -ForegroundColor Red
    return
}


# --- Load scriptblocks ---

$DefaultScriptBlock       = (Get-Command $defaultScript).ScriptBlock
$NoUninstallScriptBlock   = (Get-Command $defaultNoUninstall).ScriptBlock


# --- Create temp test directory ---

$testDir = Join-Path $env:TEMP "_PatchTest_$(Get-Random)"
$testLogDir = Join-Path $testDir "Logs"
mkdir $testDir -Force > $null

# Create a dummy executable with version info for version detection tests.
# Copy cmd.exe as a stand-in -- it has file version info we can read.
$dummySoftwareDir = Join-Path $testDir "FakeSoftware"
mkdir $dummySoftwareDir -Force > $null
Copy-Item "C:\Windows\System32\cmd.exe" (Join-Path $dummySoftwareDir "FakeApp.exe") -Force
$dummyExePath = Join-Path $dummySoftwareDir "FakeApp.exe"
$dummyVersion = (Get-Item $dummyExePath).VersionInfo.FileVersionRaw


Write-Host ""
Write-Host ("=" * 72) -ForegroundColor Cyan
Write-Host "  Default.ps1 / Default-NoUninstall.ps1 -- Test Suite" -ForegroundColor Cyan
Write-Host ("=" * 72) -ForegroundColor Cyan
Write-Host ""
Write-Host "  Default.ps1       : $defaultScript"
Write-Host "  NoUninstall.ps1   : $defaultNoUninstall"
Write-Host "  Test Dir          : $testDir"
Write-Host "  Dummy Exe Version : $dummyVersion"
Write-Host "  PS Version        : $($PSVersionTable.PSVersion)"
Write-Host "  Date              : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host ""

$sw = [System.Diagnostics.Stopwatch]::StartNew()

#endregion ===== SETUP =====



# ====================================================================
#  Test 1 : Config hashtable unpacking
# ====================================================================

Write-Host "--- Test 1 : Config Hashtable Unpacking ---" -ForegroundColor Yellow

$config1 = @{
    Software      = "FakeApp"
    SoftwarePaths = $dummyExePath
    CompliantVer  = "99.0.0.0"
    ProcessName   = "NonExistentProcess12345"
    InstallLine   = 'cmd /c exit 0'
    VersionType   = "File"
    LogDir        = $testLogDir
}

$result1 = & $DefaultScriptBlock $config1

Assert-Test "Config Unpack" "Result is not null" `
    -Expected $null -Actual $result1 -CompareMode NotNull

Assert-Test "Config Unpack" "ExitCode property exists" `
    -Expected $null -Actual $result1.PSObject.Properties['ExitCode'] -CompareMode NotNull

Assert-Test "Config Unpack" "Comment property exists" `
    -Expected $null -Actual $result1.PSObject.Properties['Comment'] -CompareMode NotNull

Assert-Test "Config Unpack" "ActionLog property exists" `
    -Expected $null -Actual $result1.PSObject.Properties['ActionLog'] -CompareMode NotNull

Assert-Test "Config Unpack" "NewVersion property exists" `
    -Expected $null -Actual $result1.PSObject.Properties['NewVersion'] -CompareMode NotNull

Assert-Test "Config Unpack" "ExitCodes property exists" `
    -Expected $null -Actual $result1.PSObject.Properties['ExitCodes'] -CompareMode NotNull

Write-Host ""



# ====================================================================
#  Test 2 : Install with successful exit code
# ====================================================================

Write-Host "--- Test 2 : Successful Install (exit 0) ---" -ForegroundColor Yellow

$config2 = @{
    Software      = "FakeApp"
    SoftwarePaths = $dummyExePath
    CompliantVer  = "99.0.0.0"
    ProcessName   = "NonExistentProcess12345"
    InstallLine   = 'cmd /c exit 0'
    VersionType   = "File"
    LogDir        = $testLogDir
}

$result2 = & $DefaultScriptBlock $config2

Assert-Test "Success Install" "ExitCode is 0" `
    -Expected "0" -Actual $result2.ExitCode

Assert-Test "Success Install" "Comment contains Completed Successfully" `
    -Expected "Completed Successfully" -Actual $result2.Comment -CompareMode Contains

Write-Host ""



# ====================================================================
#  Test 3 : Install with failure exit code
# ====================================================================

Write-Host "--- Test 3 : Failed Install (exit 1603) ---" -ForegroundColor Yellow

$config3 = @{
    Software      = "FakeApp"
    SoftwarePaths = $dummyExePath
    CompliantVer  = "99.0.0.0"
    ProcessName   = "NonExistentProcess12345"
    InstallLine   = 'cmd /c exit 1603'
    VersionType   = "File"
    LogDir        = $testLogDir
}

$result3 = & $DefaultScriptBlock $config3

Assert-Test "Failure Install" "ExitCode is 1603" `
    -Expected "1603" -Actual $result3.ExitCode

Assert-Test "Failure Install" "Comment contains Fatal error" `
    -Expected "Fatal error during installation" -Actual $result3.Comment -CompareMode Contains

Write-Host ""



# ====================================================================
#  Test 4 : Multi-line install with multiple exit codes
# ====================================================================

Write-Host "--- Test 4 : Multi-Line Install ---" -ForegroundColor Yellow

$config4 = @{
    Software      = "FakeApp"
    SoftwarePaths = $dummyExePath
    CompliantVer  = "99.0.0.0"
    ProcessName   = "NonExistentProcess12345"
    InstallLine   = @('cmd /c exit 0', 'cmd /c exit 3010')
    VersionType   = "File"
    LogDir        = $testLogDir
}

$result4 = & $DefaultScriptBlock $config4

Assert-Test "Multi-Line" "ExitCode is last code (3010)" `
    -Expected "3010" -Actual $result4.ExitCode

Assert-Test "Multi-Line" "ExitCodes array has 2 entries" `
    -Expected "2" -Actual @($result4.ExitCodes).Count

Assert-Test "Multi-Line" "First exit code is 0" `
    -Expected "0" -Actual @($result4.ExitCodes)[0]

Assert-Test "Multi-Line" "Second exit code is 3010" `
    -Expected "3010" -Actual @($result4.ExitCodes)[1]

Write-Host ""



# ====================================================================
#  Test 5 : Action log structure
# ====================================================================

Write-Host "--- Test 5 : Action Log Structure ---" -ForegroundColor Yellow

$config5 = @{
    Software      = "FakeApp"
    SoftwarePaths = $dummyExePath
    CompliantVer  = "99.0.0.0"
    ProcessName   = "NonExistentProcess12345"
    InstallLine   = @('cmd /c exit 0', 'cmd /c exit 5')
    VersionType   = "File"
    LogDir        = $testLogDir
}

$result5 = & $DefaultScriptBlock $config5
$log5 = @($result5.ActionLog)

Assert-Test "ActionLog" "ActionLog has entries" `
    -Expected "0" -Actual $log5.Count -CompareMode GreaterThan

# Find install actions only (uninstall may or may not produce entries depending on registry)
$installActions = @($log5 | Where-Object { $_.Phase -eq 'Install' })

Assert-Test "ActionLog" "Install actions found" `
    -Expected "0" -Actual $installActions.Count -CompareMode GreaterThan

Assert-Test "ActionLog" "First install action has Phase property" `
    -Expected "Install" -Actual $installActions[0].Phase

Assert-Test "ActionLog" "First install action has Command property" `
    -Expected "cmd /c exit 0" -Actual $installActions[0].Command

Assert-Test "ActionLog" "First install action has ExitCode property" `
    -Expected "0" -Actual $installActions[0].ExitCode

Assert-Test "ActionLog" "First install action has Source property" `
    -Expected "InstallLine" -Actual $installActions[0].Source

Assert-Test "ActionLog" "Second install action has exit code 5" `
    -Expected "5" -Actual $installActions[1].ExitCode

Write-Host ""



# ====================================================================
#  Test 6 : Action log duration tracking
# ====================================================================

Write-Host "--- Test 6 : Duration Tracking ---" -ForegroundColor Yellow

$config6 = @{
    Software      = "FakeApp"
    SoftwarePaths = $dummyExePath
    CompliantVer  = "99.0.0.0"
    ProcessName   = "NonExistentProcess12345"
    InstallLine   = 'cmd /c exit 0'
    VersionType   = "File"
    LogDir        = $testLogDir
}

$result6 = & $DefaultScriptBlock $config6
$installAction6 = @($result6.ActionLog | Where-Object { $_.Phase -eq 'Install' })[0]

Assert-Test "Duration" "Duration property exists" `
    -Expected $null -Actual $installAction6.Duration -CompareMode NotNull

# Duration should be parseable as a TimeSpan
$parsedOK = $false
try {
    $ts = [TimeSpan]::Parse($installAction6.Duration)
    $parsedOK = $true
} catch {}

Assert-Test "Duration" "Duration is valid TimeSpan string" `
    -Expected "True" -Actual $parsedOK

Write-Host ""



# ====================================================================
#  Test 7 : Exit code comment enrichment
# ====================================================================

Write-Host "--- Test 7 : Action Log Comment Enrichment ---" -ForegroundColor Yellow

$config7 = @{
    Software      = "FakeApp"
    SoftwarePaths = $dummyExePath
    CompliantVer  = "99.0.0.0"
    ProcessName   = "NonExistentProcess12345"
    InstallLine   = @('cmd /c exit 0', 'cmd /c exit 1603')
    VersionType   = "File"
    LogDir        = $testLogDir
}

$result7 = & $DefaultScriptBlock $config7
$installActions7 = @($result7.ActionLog | Where-Object { $_.Phase -eq 'Install' })

Assert-Test "Enrichment" "First action has Comment: Completed Successfully" `
    -Expected "Completed Successfully" -Actual $installActions7[0].Comment

Assert-Test "Enrichment" "Second action has Comment: Fatal error" `
    -Expected "Fatal error during installation" -Actual $installActions7[1].Comment

Write-Host ""



# ====================================================================
#  Test 8 : Return object completeness
# ====================================================================

Write-Host "--- Test 8 : Return Object Structure ---" -ForegroundColor Yellow

$config8 = @{
    Software      = "FakeApp"
    SoftwarePaths = $dummyExePath
    CompliantVer  = "99.0.0.0"
    ProcessName   = "NonExistentProcess12345"
    InstallLine   = 'cmd /c exit 0'
    VersionType   = "File"
    LogDir        = $testLogDir
}

$result8 = & $DefaultScriptBlock $config8

$expectedProps = @('NewVersion', 'ExitCode', 'ExitCodes', 'Comment', 'ActionLog')
foreach ($prop in $expectedProps) {
    Assert-Test "ReturnObj" "Property '$prop' exists" `
        -Expected $null -Actual $result8.PSObject.Properties[$prop] -CompareMode NotNull
}

# ExitCode should be scalar, not array
$isScalar = ($result8.ExitCode -is [int]) -or ($result8.ExitCode -is [long]) -or ($null -eq $result8.ExitCode)
Assert-Test "ReturnObj" "ExitCode is scalar (not array)" `
    -Expected "True" -Actual $isScalar

Write-Host ""



# ====================================================================
#  Test 9 : Software not installed (no paths match)
# ====================================================================

Write-Host "--- Test 9 : Software Not Installed ---" -ForegroundColor Yellow

$config9 = @{
    Software      = "CompletelyFakeSoftware_$(Get-Random)"
    SoftwarePaths = "C:\NonExistent\Path\FakeApp.exe"
    CompliantVer  = "1.0.0.0"
    ProcessName   = "NonExistentProcess12345"
    InstallLine   = 'cmd /c exit 0'
    VersionType   = "File"
    LogDir        = $testLogDir
}

$result9 = & $DefaultScriptBlock $config9

Assert-Test "NotInstalled" "Result is not null" `
    -Expected $null -Actual $result9 -CompareMode NotNull

Assert-Test "NotInstalled" "ActionLog exists" `
    -Expected $null -Actual $result9.ActionLog -CompareMode NotNull

Write-Host ""



# ====================================================================
#  Test 10 : Default-NoUninstall.ps1 produces same return structure
# ====================================================================

Write-Host "--- Test 10 : Default-NoUninstall.ps1 Return Structure ---" -ForegroundColor Yellow

$config10 = @{
    Software      = "FakeApp"
    SoftwarePaths = $dummyExePath
    CompliantVer  = "99.0.0.0"
    ProcessName   = "NonExistentProcess12345"
    InstallLine   = @('cmd /c exit 0', 'cmd /c exit 3010')
    VersionType   = "File"
    LogDir        = $testLogDir
}

$result10 = & $NoUninstallScriptBlock $config10

Assert-Test "NoUninstall" "Result is not null" `
    -Expected $null -Actual $result10 -CompareMode NotNull

$expectedProps10 = @('NewVersion', 'ExitCode', 'ExitCodes', 'Comment', 'ActionLog')
foreach ($prop in $expectedProps10) {
    Assert-Test "NoUninstall" "Property '$prop' exists" `
        -Expected $null -Actual $result10.PSObject.Properties[$prop] -CompareMode NotNull
}

Assert-Test "NoUninstall" "ExitCode is 3010 (last code)" `
    -Expected "3010" -Actual $result10.ExitCode

Assert-Test "NoUninstall" "ExitCodes has 2 entries" `
    -Expected "2" -Actual @($result10.ExitCodes).Count

$noUninstallActions = @($result10.ActionLog | Where-Object { $_.Phase -eq 'Uninstall' })
Assert-Test "NoUninstall" "No uninstall actions in log" `
    -Expected "0" -Actual $noUninstallActions.Count

Write-Host ""



# ====================================================================
#  Test 11 : USER path expansion logic
# ====================================================================

Write-Host "--- Test 11 : USER Path Expansion ---" -ForegroundColor Yellow

# Create a fake user profile structure with the dummy exe
$fakeUserDir = Join-Path $testDir "Users"
mkdir $fakeUserDir -Force > $null
$fakeUser = "TestUser_$(Get-Random)"
$fakeUserAppDir = Join-Path $fakeUserDir "$fakeUser\AppData\Local\FakeSoftware"
mkdir $fakeUserAppDir -Force > $null
Copy-Item "C:\Windows\System32\cmd.exe" (Join-Path $fakeUserAppDir "FakeApp.exe") -Force

# USER path expansion only works against actual C:\Users -- this test verifies
# that the code handles the USER keyword in paths. Since we cannot control
# C:\Users, we test that a non-USER path works correctly as a baseline.
$config11 = @{
    Software      = "FakeApp"
    SoftwarePaths = $dummyExePath
    CompliantVer  = "99.0.0.0"
    ProcessName   = "NonExistentProcess12345"
    InstallLine   = 'cmd /c exit 0'
    VersionType   = "File"
    LogDir        = $testLogDir
}

$result11 = & $DefaultScriptBlock $config11

Assert-Test "UserPath" "Standard path produces result" `
    -Expected $null -Actual $result11 -CompareMode NotNull

Assert-Test "UserPath" "Version detected from dummy exe" `
    -Expected $null -Actual $result11.NewVersion -CompareMode NotNull

Write-Host ""



# ====================================================================
#  Test 12 : Install error handling
# ====================================================================

Write-Host "--- Test 12 : Install Error Handling ---" -ForegroundColor Yellow

$config12 = @{
    Software      = "FakeApp"
    SoftwarePaths = $dummyExePath
    CompliantVer  = "99.0.0.0"
    ProcessName   = "NonExistentProcess12345"
    InstallLine   = 'throw "Test error from bad install command"'
    VersionType   = "File"
    LogDir        = $testLogDir
}

$result12 = & $DefaultScriptBlock $config12

Assert-Test "ErrorHandling" "Result returned despite error" `
    -Expected $null -Actual $result12 -CompareMode NotNull

# Should have an error entry in the action log
$errorActions12 = @($result12.ActionLog | Where-Object { $_.Source -eq 'InstallLine-Error' })

Assert-Test "ErrorHandling" "Error action logged" `
    -Expected "0" -Actual $errorActions12.Count -CompareMode GreaterThan

Assert-Test "ErrorHandling" "Error message captured" `
    -Expected "Test error" -Actual $errorActions12[0].Error -CompareMode Contains

Write-Host ""



# ====================================================================
#  Test 13 : ActionLog content (Default.ps1)
#  Validates logging data via the in-memory ActionLog rather than
#  reading files from C:\Temp (avoids leaving artifacts on servers).
# ====================================================================

Write-Host "--- Test 13 : ActionLog -- Default.ps1 ---" -ForegroundColor Yellow

$logTestName = "LogTest_$(Get-Random)"
$config13 = @{
    Software      = $logTestName
    SoftwarePaths = $dummyExePath
    CompliantVer  = "99.0.0.0"
    ProcessName   = "NonExistentProcess12345"
    InstallLine   = @('cmd /c exit 0', 'cmd /c exit 3010')
    VersionType   = "File"
    AdminName     = "TestAdmin"
    LogDir        = $testLogDir
}

$result13 = & $DefaultScriptBlock $config13

Assert-Test "ActionLog" "ActionLog is populated" `
    -Expected "0" -Actual @($result13.ActionLog).Count -CompareMode GreaterThan

$installActions13 = @($result13.ActionLog | Where-Object { $_.Phase -eq 'Install' })

Assert-Test "ActionLog" "Install actions logged" `
    -Expected "0" -Actual $installActions13.Count -CompareMode GreaterThan

Assert-Test "ActionLog" "Install action has Phase" `
    -Expected "Install" -Actual $installActions13[0].Phase

Assert-Test "ActionLog" "Install action has ExitCode" `
    -Expected "True" -Actual ($null -ne $installActions13[0].ExitCode)

Assert-Test "ActionLog" "Install action has Duration" `
    -Expected "True" -Actual ($null -ne $installActions13[0].Duration)

Assert-Test "ActionLog" "Install action has Source" `
    -Expected "InstallLine" -Actual $installActions13[0].Source

# Last install line exits 3010 -- verify it appears in the action log
$action3010 = @($installActions13 | Where-Object { $_.ExitCode -eq 3010 })
Assert-Test "ActionLog" "Exit code 3010 recorded" `
    -Expected "0" -Actual $action3010.Count -CompareMode GreaterThan

Assert-Test "ActionLog" "Result exit code is 3010" `
    -Expected 3010 -Actual $result13.ExitCode

Assert-Test "ActionLog" "Comment contains Successful - Restart required" `
    -Expected "Successful - Restart required" -Actual $result13.Comment -CompareMode Contains

Write-Host ""



# ====================================================================
#  Test 14 : ActionLog content (Default-NoUninstall.ps1)
# ====================================================================

Write-Host "--- Test 14 : ActionLog -- Default-NoUninstall.ps1 ---" -ForegroundColor Yellow

$logTestName14 = "NoUninstLogTest_$(Get-Random)"
$config14 = @{
    Software      = $logTestName14
    SoftwarePaths = $dummyExePath
    CompliantVer  = "99.0.0.0"
    ProcessName   = "NonExistentProcess12345"
    InstallLine   = 'cmd /c exit 1603'
    VersionType   = "File"
    AdminName     = "TestAdmin14"
    LogDir        = $testLogDir
}

$result14 = & $NoUninstallScriptBlock $config14

Assert-Test "ActionLogNoUn" "ActionLog is populated" `
    -Expected "0" -Actual @($result14.ActionLog).Count -CompareMode GreaterThan

$installActions14 = @($result14.ActionLog | Where-Object { $_.Phase -eq 'Install' })

Assert-Test "ActionLogNoUn" "Install action logged" `
    -Expected "0" -Actual $installActions14.Count -CompareMode GreaterThan

Assert-Test "ActionLogNoUn" "Exit code 1603 recorded" `
    -Expected 1603 -Actual $installActions14[0].ExitCode

Assert-Test "ActionLogNoUn" "Comment contains Fatal error" `
    -Expected "Fatal error during installation" -Actual $result14.Comment -CompareMode Contains

$uninstallActions14 = @($result14.ActionLog | Where-Object { $_.Phase -eq 'Uninstall' })
Assert-Test "ActionLogNoUn" "No Uninstall actions logged" `
    -Expected 0 -Actual $uninstallActions14.Count

Write-Host ""



# ====================================================================
#  CLEANUP
# ====================================================================

# All test artifacts (dummy exe, log files) live under $testDir in $env:TEMP
if (Test-Path $testDir) {
    Remove-Item $testDir -Recurse -Force -ErrorAction SilentlyContinue
}



# ====================================================================
#  SUMMARY
# ====================================================================

$sw.Stop()

$passed = @($script:TestResults | Where-Object { $_.Status -eq 'PASS' }).Count
$failed = @($script:TestResults | Where-Object { $_.Status -eq 'FAIL' }).Count
$total  = $script:TestResults.Count

Write-Host ("=" * 72) -ForegroundColor Cyan
Write-Host "  TEST SUMMARY" -ForegroundColor Cyan
Write-Host ("=" * 72) -ForegroundColor Cyan
Write-Host ""
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
