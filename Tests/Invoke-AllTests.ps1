# DOTS formatting comment

function Invoke-AllTests {

<#
    .SYNOPSIS
        Runs all Test-*.ps1 test suites sequentially.
    .DESCRIPTION
        Discovers every Test-*.ps1 file in the Tests folder and runs each in
        a child powershell.exe process that shares the current console window.
        This keeps console features (colors, cursor control, progress) working
        while protecting the caller from exit statements in test scripts.

        Each suite's output prints directly to the console with full fidelity.
        After all suites finish, a summary line shows exit codes per suite.

        Written by Skyler Werner
        Date: 2026/03/27
    .PARAMETER Filter
        Optional wildcard filter for test file names. Default is 'Test-*.ps1'.
        Example: -Filter 'Test-Merge*' to run only MergeMainSwitch tests.
    .EXAMPLE
        Invoke-AllTests
        Runs all test suites in the Tests folder.
    .EXAMPLE
        Invoke-AllTests -Filter 'Test-Default*'
        Runs only Test-Default.ps1 and Test-DefaultPSExec.ps1.
#>

    param(
        [string]$Filter = 'Test-*.ps1'
    )

    # Resolve Tests directory. When dot-sourced from the profile, $testsPath
    # is set by the profile before dot-sourcing. Otherwise fall back to the
    # script file's own directory.
    $testsDir = $null
    if ($script:testsPath -and (Test-Path $script:testsPath)) {
        $testsDir = $script:testsPath
    }
    if (-not $testsDir) {
        $testsDir = Split-Path -Parent $MyInvocation.MyCommand.ScriptBlock.File
    }
    if (-not $testsDir) {
        $testsDir = $PSScriptRoot
    }
    if (-not $testsDir) {
        Write-Host 'Could not determine Tests directory.' -ForegroundColor Red
        return
    }

    $testFiles = @(Get-ChildItem -Path $testsDir -Filter $Filter -File |
        Where-Object { $_.Name -ne 'Invoke-AllTests.ps1' } |
        Sort-Object Name)

    if ($testFiles.Count -eq 0) {
        Write-Host "No test files matching '$Filter' found in $testsDir" -ForegroundColor Yellow
        return
    }

    Write-Host ''
    Write-Host ('=' * 72) -ForegroundColor Cyan
    Write-Host '  TEST RUNNER -- Invoke-AllTests' -ForegroundColor Cyan
    Write-Host ('=' * 72) -ForegroundColor Cyan
    Write-Host "  Suites found: $($testFiles.Count)"
    foreach ($f in $testFiles) { Write-Host "    - $($f.Name)" }
    Write-Host ''

    # --- Run each suite in its own process ---

    $suiteResults = New-Object System.Collections.Generic.List[PSCustomObject]
    $overallSW = [System.Diagnostics.Stopwatch]::StartNew()

    foreach ($testFile in $testFiles) {

        Write-Host ''
        Write-Host ('#' * 72) -ForegroundColor Magenta
        Write-Host "  SUITE: $($testFile.Name)" -ForegroundColor Magenta
        Write-Host ('#' * 72) -ForegroundColor Magenta
        Write-Host ''

        $suiteSW = [System.Diagnostics.Stopwatch]::StartNew()

        # Run in a child process with -NoNewWindow so it shares the real
        # console. This keeps colors, [Console]::CursorVisible, and progress
        # display working. Output goes directly to the screen -- no capture.
        $proc = Start-Process powershell.exe -ArgumentList `
            "-NoProfile -ExecutionPolicy Bypass -File `"$($testFile.FullName)`"" `
            -NoNewWindow -Wait -PassThru

        $suiteSW.Stop()

        $suiteResults.Add([PSCustomObject]@{
            Suite    = $testFile.Name
            ExitCode = $proc.ExitCode
            Time     = "$([math]::Round($suiteSW.Elapsed.TotalSeconds, 1))s"
        })
    }

    $overallSW.Stop()

    # --- Summary ---

    $failedSuites = @($suiteResults | Where-Object { $_.ExitCode -ne 0 })

    Write-Host ''
    Write-Host ''
    Write-Host ('=' * 72) -ForegroundColor Cyan
    Write-Host '  GRAND SUMMARY' -ForegroundColor Cyan
    Write-Host ('=' * 72) -ForegroundColor Cyan
    Write-Host ''

    foreach ($r in $suiteResults) {
        $color = if ($r.ExitCode -eq 0) { 'Green' } else { 'Red' }
        $status = if ($r.ExitCode -eq 0) { 'OK' } else { "EXIT $($r.ExitCode)" }
        Write-Host "  $($r.Suite)  --  $status  ($($r.Time))" -ForegroundColor $color
    }

    Write-Host ''
    Write-Host "  Suites: $($suiteResults.Count)    Time: $([math]::Round($overallSW.Elapsed.TotalSeconds, 1))s"

    if ($failedSuites.Count -gt 0) {
        Write-Host "  $($failedSuites.Count) suite(s) exited non-zero." -ForegroundColor Red
    } else {
        Write-Host '  All suites exited clean.' -ForegroundColor Green
    }
    Write-Host ''
}
