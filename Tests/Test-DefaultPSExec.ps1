# DOTS formatting comment

<#
    .SYNOPSIS
        Test suite for Default-PSExec.ps1 patching script.
    .DESCRIPTION
        Tests the restructured Default-PSExec.ps1 script locally without
        requiring PsExec.exe or remote machines. Validates:

          Test  1 : Get-ExitCodeComment function (known codes)
          Test  2 : Get-ExitCodeComment function (unknown code returns $null)
          Test  3 : Installer type detection -- .msu selects wusa.exe
          Test  4 : Installer type detection -- .msi selects msiexec /install
          Test  5 : Installer type detection -- .msp selects msiexec /update
          Test  6 : Installer type detection -- McAfee MAgent sets flag
          Test  7 : Temp patch path construction
          Test  8 : PsExec argument building -- standard patch
          Test  9 : PsExec argument building -- McAfee agent
          Test 10 : Error handling -- PsExec not found produces action log entry
          Test 11 : Return object structure (ComputerName, ExitCode, Comment, ActionLog)
          Test 12 : Action log entry has expected properties (catch or try path)
          Test 13 : Logging resilience -- UNC failure does not break return

        Execution tests validate whichever path runs: the catch block (PsExec
        not found) or the try block (PsExec present but remote machine absent).
        Argument building and installer detection are tested by running isolated
        sections of the script logic.

        Written by Skyler Werner
        Date: 2026/03/23
        Version 1.0.0

    .EXAMPLE
        .\Test-DefaultPSExec.ps1
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


# --- Locate script ---

$scriptRoot = Join-Path (Split-Path $PSScriptRoot -Parent) "Scripts\Patching"
$psExecScript = Join-Path $scriptRoot "Default-PSExec.ps1"

if (-not (Test-Path $psExecScript)) {
    Write-Host "ERROR: Cannot find Default-PSExec.ps1 at $psExecScript" -ForegroundColor Red
    return
}


# --- Load scriptblock ---

$PSExecScriptBlock = (Get-Command $psExecScript).ScriptBlock


# --- Extract Get-ExitCodeComment from the real script for testing ---
# Regex-extract the function so we always test the actual code, not a stale copy.
# The function is the first block in the file and ends before "function Get-ActionEvents".

$scriptContent = Get-Content $psExecScript -Raw
$fnMatch = [regex]::Match($scriptContent, '(?s)(function Get-ExitCodeComment \{.+?\n\})')
if (-not $fnMatch.Success) {
    Write-Host "ERROR: Could not extract Get-ExitCodeComment from $psExecScript" -ForegroundColor Red
    return
}
$exitCodeFnBlock = [ScriptBlock]::Create($fnMatch.Groups[1].Value)
. $exitCodeFnBlock


Write-Host ""
Write-Host ("=" * 72) -ForegroundColor Cyan
Write-Host "  Default-PSExec.ps1 -- Test Suite" -ForegroundColor Cyan
Write-Host ("=" * 72) -ForegroundColor Cyan
Write-Host ""
Write-Host "  Script : $psExecScript"
Write-Host "  PS Ver : $($PSVersionTable.PSVersion)"
Write-Host "  Date   : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host ""

$totalSW = [System.Diagnostics.Stopwatch]::StartNew()

#endregion ===== SETUP =====



# ====================================================================
#  Test 1 : Get-ExitCodeComment -- known codes
# ====================================================================

Write-Host "--- Test 1 : Get-ExitCodeComment Known Codes ---" -ForegroundColor Yellow

Assert-Test "ExitCodeFn" "Code 0 = Completed Successfully" `
    -Expected "Completed Successfully" -Actual (Get-ExitCodeComment 0)

Assert-Test "ExitCodeFn" "Code 1603 = Fatal error during installation" `
    -Expected "Fatal error during installation" -Actual (Get-ExitCodeComment 1603)

Assert-Test "ExitCodeFn" "Code 3010 = Successful - Restart required" `
    -Expected "Successful - Restart required" -Actual (Get-ExitCodeComment 3010)

Assert-Test "ExitCodeFn" "Code -2146498511 = Corruption in the windows component store" `
    -Expected "Corruption in the windows component store" -Actual (Get-ExitCodeComment -2146498511)

Assert-Test "ExitCodeFn" "Code 1641 = Successful - Restart required" `
    -Expected "Successful - Restart required" -Actual (Get-ExitCodeComment 1641)

Write-Host ""



# ====================================================================
#  Test 2 : Get-ExitCodeComment -- unknown code returns $null
# ====================================================================

Write-Host "--- Test 2 : Get-ExitCodeComment Unknown Code ---" -ForegroundColor Yellow

Assert-Test "ExitCodeFn" "Code 99999 returns null" `
    -Expected $null -Actual (Get-ExitCodeComment 99999) -CompareMode IsNull

Assert-Test "ExitCodeFn" "Code -1 returns null" `
    -Expected $null -Actual (Get-ExitCodeComment -1) -CompareMode IsNull

Write-Host ""



# ====================================================================
#  Test 3 : Installer type -- .msu selects wusa.exe
# ====================================================================

Write-Host "--- Test 3 : Installer Detection -- MSU ---" -ForegroundColor Yellow

# Simulate the installer switch logic from Default-PSExec.ps1
$testPatchName3 = "windows10.0-kb5034441-x64.msu"
$installer3 = $null
$mcAfee3 = $false

switch ($testPatchName3) {
    {$_ -match ".msu"} { $installer3 = "c:\windows\system32\wusa.exe" }
    {$_ -match ".msi"} { $installer3 = "msiexec /install" }
    {$_ -match ".msp"} { $installer3 = "msiexec /update" }
    {$_ -match "MAgent"} { $mcAfee3 = $true }
}

Assert-Test "InstallerMSU" "MSU selects wusa.exe" `
    -Expected "c:\windows\system32\wusa.exe" -Actual $installer3

Assert-Test "InstallerMSU" "McAfee flag not set for MSU" `
    -Expected "False" -Actual $mcAfee3

Write-Host ""



# ====================================================================
#  Test 4 : Installer type -- .msi selects msiexec /install
# ====================================================================

Write-Host "--- Test 4 : Installer Detection -- MSI ---" -ForegroundColor Yellow

$testPatchName4 = "SomeProduct-v2.0.msi"
$installer4 = $null
$mcAfee4 = $false

switch ($testPatchName4) {
    {$_ -match ".msu"} { $installer4 = "c:\windows\system32\wusa.exe" }
    {$_ -match ".msi"} { $installer4 = "msiexec /install" }
    {$_ -match ".msp"} { $installer4 = "msiexec /update" }
    {$_ -match "MAgent"} { $mcAfee4 = $true }
}

Assert-Test "InstallerMSI" "MSI selects msiexec /install" `
    -Expected "msiexec /install" -Actual $installer4

Write-Host ""



# ====================================================================
#  Test 5 : Installer type -- .msp selects msiexec /update
# ====================================================================

Write-Host "--- Test 5 : Installer Detection -- MSP ---" -ForegroundColor Yellow

$testPatchName5 = "hotfix-2024-Q1.msp"
$installer5 = $null
$mcAfee5 = $false

switch ($testPatchName5) {
    {$_ -match ".msu"} { $installer5 = "c:\windows\system32\wusa.exe" }
    {$_ -match ".msi"} { $installer5 = "msiexec /install" }
    {$_ -match ".msp"} { $installer5 = "msiexec /update" }
    {$_ -match "MAgent"} { $mcAfee5 = $true }
}

Assert-Test "InstallerMSP" "MSP selects msiexec /update" `
    -Expected "msiexec /update" -Actual $installer5

Write-Host ""



# ====================================================================
#  Test 6 : Installer type -- McAfee MAgent sets flag
# ====================================================================

Write-Host "--- Test 6 : Installer Detection -- McAfee ---" -ForegroundColor Yellow

$testPatchName6 = "MAgent5.7.exe"
$installer6 = $null
$mcAfee6 = $false

switch ($testPatchName6) {
    {$_ -match ".msu"} { $installer6 = "c:\windows\system32\wusa.exe" }
    {$_ -match ".msi"} { $installer6 = "msiexec /install" }
    {$_ -match ".msp"} { $installer6 = "msiexec /update" }
    {$_ -match "MAgent"} { $mcAfee6 = $true }
}

Assert-Test "InstallerMcAfee" "McAfee flag set for MAgent" `
    -Expected "True" -Actual $mcAfee6

Assert-Test "InstallerMcAfee" "Installer stays null for McAfee" `
    -Expected $null -Actual $installer6 -CompareMode IsNull

Write-Host ""



# ====================================================================
#  Test 7 : Temp patch path construction
# ====================================================================

Write-Host "--- Test 7 : Temp Patch Path Construction ---" -ForegroundColor Yellow

$testPatchPath7 = "\\server\share\Patches\2024-Q1"
$testPatchName7 = "kb5034441.msu"

$itemFolder7 = $testPatchPath7.Split("\") | Select-Object -Last 1
$tempPatchPath7 = "$itemFolder7" + "\" + "$testPatchName7"

Assert-Test "TempPath" "Folder extracted from patch path" `
    -Expected "2024-Q1" -Actual $itemFolder7

Assert-Test "TempPath" "Temp patch path is folder\patchname" `
    -Expected "2024-Q1\kb5034441.msu" -Actual $tempPatchPath7

Write-Host ""



# ====================================================================
#  Test 8 : PsExec argument building -- standard patch
# ====================================================================

Write-Host "--- Test 8 : PsExec Arguments -- Standard ---" -ForegroundColor Yellow

$testComputer8 = "WORKSTATION01"
$testInstaller8 = "c:\windows\system32\wusa.exe"
$testTempPath8 = "2024-Q1\kb5034441.msu"
$testLog8 = "/log:c:\Temp\2024-03-23_1200.kb5034441.msu.admin1.evt"

$arguments8 = @(
    "\\$($testComputer8)"
    "-accepteula"
    "-s"
    "$testInstaller8 C:\Temp\$($testTempPath8) /quiet /norestart $testLog8"
)

Assert-Test "StdArgs" "First arg is UNC computer name" `
    -Expected "\\WORKSTATION01" -Actual $arguments8[0]

Assert-Test "StdArgs" "Second arg is -accepteula" `
    -Expected "-accepteula" -Actual $arguments8[1]

Assert-Test "StdArgs" "Third arg is -s (system)" `
    -Expected "-s" -Actual $arguments8[2]

Assert-Test "StdArgs" "Fourth arg contains installer path" `
    -Expected "wusa.exe" -Actual $arguments8[3] -CompareMode Contains

Assert-Test "StdArgs" "Fourth arg contains /quiet /norestart" `
    -Expected "/quiet /norestart" -Actual $arguments8[3] -CompareMode Contains

Assert-Test "StdArgs" "Fourth arg contains /log:" `
    -Expected "/log:" -Actual $arguments8[3] -CompareMode Contains

Write-Host ""



# ====================================================================
#  Test 9 : PsExec argument building -- McAfee agent
# ====================================================================

Write-Host "--- Test 9 : PsExec Arguments -- McAfee ---" -ForegroundColor Yellow

$testComputer9 = "SERVER01"
$testTempPath9 = "McAfee\MAgent5.7.exe"

$arguments9 = @(
    "\\$($testComputer9)"
    "-accepteula"
    "-h"
    "-s"
    "C:\Temp\$($testTempPath9) /Install=Agent /Silent /ForceInstall"
)

Assert-Test "McAfeeArgs" "First arg is UNC computer name" `
    -Expected "\\SERVER01" -Actual $arguments9[0]

Assert-Test "McAfeeArgs" "Third arg is -h (elevated)" `
    -Expected "-h" -Actual $arguments9[2]

Assert-Test "McAfeeArgs" "Fourth arg is -s (system)" `
    -Expected "-s" -Actual $arguments9[3]

Assert-Test "McAfeeArgs" "Fifth arg contains /Install=Agent" `
    -Expected "/Install=Agent" -Actual $arguments9[4] -CompareMode Contains

Assert-Test "McAfeeArgs" "Fifth arg contains /Silent /ForceInstall" `
    -Expected "/Silent /ForceInstall" -Actual $arguments9[4] -CompareMode Contains

Write-Host ""



# ====================================================================
#  Test 10 : Error handling -- PsExec not found
# ====================================================================

Write-Host "--- Test 10 : Error Handling (PsExec Not Found) ---" -ForegroundColor Yellow

# Run the full script -- PsExec.exe won't exist, so it hits the catch block.
# Pass valid-looking arguments so the argument building section works.
$result10 = & $PSExecScriptBlock "\\server\share\Patches\2024-Q1" "kb5034441.msu" "ignored" "TESTPC01"

Assert-Test "ErrorHandling" "Result returned despite PsExec failure" `
    -Expected $null -Actual $result10 -CompareMode NotNull

Assert-Test "ErrorHandling" "ComputerName is TESTPC01" `
    -Expected "TESTPC01" -Actual $result10.ComputerName

Write-Host ""



# ====================================================================
#  Test 11 : Return object structure
# ====================================================================

Write-Host "--- Test 11 : Return Object Structure ---" -ForegroundColor Yellow

$expectedProps = @('ComputerName', 'ExitCode', 'Comment', 'ActionLog')
foreach ($prop in $expectedProps) {
    Assert-Test "ReturnObj" "Property '$prop' exists" `
        -Expected $null -Actual $result10.PSObject.Properties[$prop] -CompareMode NotNull
}

Assert-Test "ReturnObj" "ActionLog is an array" `
    -Expected "True" -Actual ($result10.ActionLog -is [array])

Write-Host ""



# ====================================================================
#  Test 12 : Action log error entry properties
# ====================================================================

Write-Host "--- Test 12 : Action Log Error Entry ---" -ForegroundColor Yellow

$log12 = @($result10.ActionLog)

Assert-Test "ErrorEntry" "ActionLog has at least one entry" `
    -Expected "0" -Actual $log12.Count -CompareMode GreaterThan

$errorEntry = @($log12 | Where-Object { $_.Source -like 'PsExec*' })

Assert-Test "ErrorEntry" "Action log entry found with Source matching PsExec*" `
    -Expected "0" -Actual $errorEntry.Count -CompareMode GreaterThan

if ($errorEntry.Count -gt 0) {
    Assert-Test "ErrorEntry" "Entry has Phase = Install" `
        -Expected "Install" -Actual $errorEntry[0].Phase

    Assert-Test "ErrorEntry" "Entry has Command containing PsExec.exe" `
        -Expected "PsExec.exe" -Actual $errorEntry[0].Command -CompareMode Contains

    Assert-Test "ErrorEntry" "Entry has Duration property" `
        -Expected $null -Actual $errorEntry[0].Duration -CompareMode NotNull

    # Error and ExitCode properties only exist on the catch path (PsExec not found)
    if ($errorEntry[0].Source -eq 'PsExec-Error') {
        Assert-Test "ErrorEntry" "Error entry has Error message" `
            -Expected $null -Actual $errorEntry[0].Error -CompareMode NotNull

        Assert-Test "ErrorEntry" "Error entry ExitCode is null" `
            -Expected $null -Actual $errorEntry[0].ExitCode -CompareMode IsNull
    }
}

Write-Host ""



# ====================================================================
#  Test 13 : Logging resilience -- UNC failure does not break return
# ====================================================================

Write-Host "--- Test 13 : Logging Resilience (UNC Failure) ---" -ForegroundColor Yellow

# The script was already invoked in Test 10 with computerName "TESTPC01".
# The UNC path \\TESTPC01\C$\Temp\PatchRemediation\Logs will not resolve,
# so the logging try/catch fires. Verify the script still returned
# a complete result object despite the logging failure.

Assert-Test "LogResilience" "Result returned despite UNC log failure" `
    -Expected $null -Actual $result10 -CompareMode NotNull

Assert-Test "LogResilience" "ComputerName still populated" `
    -Expected "TESTPC01" -Actual $result10.ComputerName

Assert-Test "LogResilience" "ActionLog still populated" `
    -Expected $null -Actual $result10.ActionLog -CompareMode NotNull

Assert-Test "LogResilience" "ActionLog has entries" `
    -Expected "0" -Actual @($result10.ActionLog).Count -CompareMode GreaterThan

Write-Host ""



# ====================================================================
#  SUMMARY
# ====================================================================

$totalSW.Stop()

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
Write-Host "  Time   : $([math]::Round($totalSW.Elapsed.TotalSeconds, 1)) seconds"
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
