# DOTS formatting comment

<#
    .SYNOPSIS
        Diagnostic test suite for PowerShell Format-Table corruption.
    .DESCRIPTION
        Isolates and tests each potential failure point that could corrupt
        PowerShell's formatting subsystem. Uses a "canary" pattern: verifies
        Format-Table renders correctly before and after each suspicious
        operation. When a canary fails, the guilty operation is identified.

        Run this on any server where Invoke-Patch / Invoke-Version breaks
        table output. Results are displayed inline and summarized at the end.

        Tests:
          1 : Environment baseline (console dimensions, language mode, policy)
          2 : Set-ExecutionPolicy output leakage
          3 : Module import output leakage
          4 : Console cursor manipulation ([Console]::SetCursorPosition)
          5 : Format-Table with 9+ columns at current buffer width
          6 : Bare Format-Table from inside a function (no Out-Host)
          7 : Format-Table piped to Out-Host from inside a function
          8 : Invoke-RunspacePool progress display simulation
          9 : Full Invoke-Patch mini-pipeline simulation
         10 : Bare .NET RunspacePool lifecycle (create/use/dispose with $Host)
         11 : Real Invoke-RunspacePool with local scriptblocks
         12 : Full Invoke-Patch pattern (RunspacePool + Format-Table in function)
         13 : Output layer diagnosis (Write-Host / Out-String / pipeline / formatter)

        Written by Skyler Werner
        Date: 2026/04/08
#>


#region ===== SETUP =====

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


# --- Canary: a simple Format-Table test that should always work ---
# Returns $true if Format-Table renders correctly, $false if corrupted.
function Test-FormatCanary {
    param([string]$Label = "Canary")

    try {
        $canaryObjects = @(
            [PSCustomObject]@{ Name = "Alpha"; Value = 1 }
            [PSCustomObject]@{ Name = "Bravo"; Value = 2 }
        )

        $rendered = ($canaryObjects | Format-Table -AutoSize | Out-String).Trim()

        # A healthy Format-Table produces a header line with "Name" and "Value"
        $hasHeader = $rendered -match 'Name\s+Value'
        # And data rows with our values
        $hasAlpha  = $rendered -match 'Alpha\s+1'
        $hasBravo  = $rendered -match 'Bravo\s+2'

        $healthy = $hasHeader -and $hasAlpha -and $hasBravo

        if (-not $healthy) {
            Write-Host "    [CANARY DEAD] $Label - Format-Table output:" -ForegroundColor Red
            Write-Host "    $rendered" -ForegroundColor DarkRed
        }

        return $healthy
    }
    catch {
        Write-Host "    [CANARY ERROR] $Label - $_" -ForegroundColor Red
        return $false
    }
}

# --- Test if the bare output pipeline works (not just Format-Table) ---
# Captures output from a scriptblock. If the pipeline is dead, $testVal is $null.
function Test-OutputCanary {
    param([string]$Label = "Output Canary")
    try {
        $marker = "CanaryOutput_$(Get-Random)"
        $testVal = & { $marker }
        $hasValue = $null -ne $testVal -and "$testVal" -match 'CanaryOutput_'
        if (-not $hasValue) {
            Write-Host "    [OUTPUT DEAD] $Label - pipeline captured nothing" -ForegroundColor Red
        }
        return $hasValue
    }
    catch {
        Write-Host "    [OUTPUT ERROR] $Label - $_" -ForegroundColor Red
        return $false
    }
}

# --- Test if Write-Host still works (host UI health) ---
# Write-Host bypasses Out-Default entirely. If this fails, the Host UI is corrupted.
function Test-WriteHostCanary {
    param([string]$Label = "WriteHost Canary")
    try {
        $stream6 = Write-Host "WriteHostCanary_Test" 6>&1
        $works = $null -ne $stream6 -and "$stream6" -match 'WriteHostCanary'
        if (-not $works) {
            # Can't use Write-Host to report if Write-Host is broken, try [Console]
            try { [Console]::Error.WriteLine("    [HOST DEAD] $Label - Write-Host produced no output") } catch {}
        }
        return $works
    }
    catch {
        try { [Console]::Error.WriteLine("    [HOST ERROR] $Label - $_") } catch {}
        return $false
    }
}

# --- Test if Out-Host works independently ---
function Test-OutHostCanary {
    param([string]$Label = "OutHost Canary")
    try {
        $marker = "OutHostCanary_$(Get-Random)"
        $rendered = $marker | Out-String
        $works = $null -ne $rendered -and "$rendered" -match 'OutHostCanary_'
        if (-not $works) {
            Write-Host "    [OUT-HOST DEAD] $Label - Out-String returned nothing" -ForegroundColor Red
        }
        return $works
    }
    catch {
        Write-Host "    [OUT-HOST ERROR] $Label - $_" -ForegroundColor Red
        return $false
    }
}

# --- Run all canaries and report which layers are alive/dead ---
function Test-AllCanaries {
    param([string]$Label)

    $fmt  = Test-FormatCanary    "$Label"
    $out  = Test-OutputCanary    "$Label"
    $wh   = Test-WriteHostCanary "$Label"
    $oh   = Test-OutHostCanary   "$Label"

    Assert-Test $Label "Format-Table canary" -Expected "True" -Actual $fmt
    Assert-Test $Label "Output pipeline canary" -Expected "True" -Actual $out
    Assert-Test $Label "Write-Host canary (host UI)" -Expected "True" -Actual $wh
    Assert-Test $Label "Out-String canary" -Expected "True" -Actual $oh

    if (-not $fmt -or -not $out -or -not $wh -or -not $oh) {
        Write-Host ""
        Write-Host "    --- Corruption Layer Report ---" -ForegroundColor Red
        Write-Host "    Write-Host (host UI):     $(if ($wh) { 'ALIVE' } else { 'DEAD' })" -ForegroundColor $(if ($wh) { 'Green' } else { 'Red' })
        Write-Host "    Out-String:                $(if ($oh) { 'ALIVE' } else { 'DEAD' })" -ForegroundColor $(if ($oh) { 'Green' } else { 'Red' })
        Write-Host "    Output pipeline:           $(if ($out) { 'ALIVE' } else { 'DEAD' })" -ForegroundColor $(if ($out) { 'Green' } else { 'Red' })
        Write-Host "    Format-Table:              $(if ($fmt) { 'ALIVE' } else { 'DEAD' })" -ForegroundColor $(if ($fmt) { 'Green' } else { 'Red' })
        Write-Host ""
        if (-not $wh) {
            Write-Host "    DIAGNOSIS: Host UI is dead. RunspacePool disposal likely corrupted `$Host." -ForegroundColor Red
        }
        elseif (-not $out) {
            Write-Host "    DIAGNOSIS: Output pipeline is dead. Formatter may be stuck (unclosed FormatStartData)." -ForegroundColor Red
        }
        elseif (-not $fmt) {
            Write-Host "    DIAGNOSIS: Format-Table broken but pipeline works. Format system corruption only." -ForegroundColor Red
        }
    }

    return ($fmt -and $out -and $wh -and $oh)
}


# --- Build sample objects matching Invoke-Patch result schema ---
function New-SampleResults {
    param([int]$Count = 3)
    $results = @()
    for ($i = 1; $i -le $Count; $i++) {
        $results += [PSCustomObject]@{
            IPAddress    = "10.0.0.$i"
            ComputerName = "TESTPC-$($i.ToString('D4'))"
            Status       = "Online"
            SoftwareName = "TestSoftware"
            Version      = "1.0.0.$i"
            Compliant    = if ($i % 2 -eq 0) { "Yes" } else { "No" }
            NewVersion   = "2.0.0.$i"
            ExitCode     = "0"
            Comment      = "Test result $i"
            AdminName    = "test.admin"
            Date         = "2026/04/08 12:00"
        }
    }
    return $results
}

#endregion ===== SETUP =====



# ================================================================
Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "  Format-Table Corruption Diagnostic Suite" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host ""
# ================================================================



#region ===== TEST 1: Environment Baseline =====

Write-Host "--- Test 1: Environment Baseline ---" -ForegroundColor Yellow

$bufferWidth  = try { [Console]::BufferWidth  } catch { "N/A" }
$bufferHeight = try { [Console]::BufferHeight } catch { "N/A" }
$windowWidth  = try { [Console]::WindowWidth  } catch { "N/A" }
$windowHeight = try { [Console]::WindowHeight } catch { "N/A" }
$rawUIWidth   = try { (Get-Host).UI.RawUI.WindowSize.Width  } catch { "N/A" }
$rawUIHeight  = try { (Get-Host).UI.RawUI.WindowSize.Height } catch { "N/A" }

Write-Host "  Console.BufferWidth:    $bufferWidth"
Write-Host "  Console.BufferHeight:   $bufferHeight"
Write-Host "  Console.WindowWidth:    $windowWidth"
Write-Host "  Console.WindowHeight:   $windowHeight"
Write-Host "  RawUI.WindowSize:       ${rawUIWidth}x${rawUIHeight}"
Write-Host "  PSVersion:              $($PSVersionTable.PSVersion)"
Write-Host "  Host.Name:              $($Host.Name)"
Write-Host "  LanguageMode:           $($ExecutionContext.SessionState.LanguageMode)"

$currentPolicy = try { Get-ExecutionPolicy -Scope Process } catch { "Error: $_" }
$machinePolicy = try { Get-ExecutionPolicy -Scope LocalMachine } catch { "Error: $_" }
$gpoPolicy     = try { Get-ExecutionPolicy -Scope MachinePolicy } catch { "Error: $_" }

Write-Host "  ExecutionPolicy (Process): $currentPolicy"
Write-Host "  ExecutionPolicy (Machine): $machinePolicy"
Write-Host "  ExecutionPolicy (GPO):     $gpoPolicy"

# Use RawUI as fallback when [Console] has no handle (non-interactive)
$effectiveWidth = if ($bufferWidth -ne "N/A") { $bufferWidth } else { $rawUIWidth }

Assert-Test "Env" "Console buffer width is at least 80" `
    -Expected 80 -Actual $effectiveWidth -CompareMode GreaterThan

Assert-Test "Env" "Host is ConsoleHost" `
    -Expected "ConsoleHost" -Actual $Host.Name

Assert-Test "Env" "Language mode is FullLanguage" `
    -Expected "FullLanguage" -Actual $ExecutionContext.SessionState.LanguageMode

$canary1 = Test-FormatCanary "Pre-test baseline"
Assert-Test "Env" "Format-Table works at baseline" `
    -Expected "True" -Actual $canary1

Write-Host ""

#endregion


#region ===== TEST 2: Set-ExecutionPolicy Output Leakage =====

Write-Host "--- Test 2: Set-ExecutionPolicy Output Leakage ---" -ForegroundColor Yellow

# Capture ALL streams from Set-ExecutionPolicy
$sepOutput = $null
$sepError  = $null
try {
    # Stream 1 (output) captured in $sepOutput
    # Stream 2 (error) captured in $sepError
    $sepOutput = Set-ExecutionPolicy Bypass -Scope Process -Force 2>&1 3>&1 4>&1 5>&1 6>&1
}
catch {
    $sepError = $_
}

$leakedOutput = if ($null -ne $sepOutput) {
    @($sepOutput).Count
} else { 0 }

Assert-Test "SetExecPolicy" "Set-ExecutionPolicy produces no output" `
    -Expected 0 -Actual $leakedOutput

if ($leakedOutput -gt 0) {
    Write-Host "    LEAKED OBJECTS:" -ForegroundColor Red
    foreach ($obj in $sepOutput) {
        Write-Host "      Type: $($obj.GetType().FullName)" -ForegroundColor DarkRed
        Write-Host "      Value: $obj" -ForegroundColor DarkRed
    }
}

if ($null -ne $sepError) {
    Write-Host "    CAUGHT ERROR: $sepError" -ForegroundColor DarkYellow
}

$canary2 = Test-FormatCanary "After Set-ExecutionPolicy"
Assert-Test "SetExecPolicy" "Format-Table still works after Set-ExecutionPolicy" `
    -Expected "True" -Actual $canary2

Write-Host ""

#endregion


#region ===== TEST 3: Module Import Output Leakage =====

Write-Host "--- Test 3: Module Import Output Leakage ---" -ForegroundColor Yellow

# Check if modules are already loaded (they should be from profile)
$loadedModules = @(Get-Module | Select-Object -ExpandProperty Name)
Write-Host "  Currently loaded modules: $($loadedModules -join ', ')"

# Test reimporting each module and check for leaked output
if (Test-Path "$env:APPDATA\Patching\Paths.txt") {
    $pathsContent = Get-Content "$env:APPDATA\Patching\Paths.txt"
    $modulePaths = @()
    foreach ($line in $pathsContent) {
        if ($line -match "ModulePath :") {
            $modulePaths += $line.Replace("ModulePath : ","")
        }
    }

    if ($modulePaths.Count -gt 0 -and (Test-Path $modulePaths[0])) {
        $moduleFolders = Get-ChildItem $modulePaths[0] -Directory
        $totalLeaked = 0

        foreach ($folder in $moduleFolders) {
            $importOutput = Import-Module $folder.FullName -Force 2>&1
            $leaked = if ($null -ne $importOutput) { @($importOutput).Count } else { 0 }

            if ($leaked -gt 0) {
                Write-Host "    $($folder.Name) leaked $leaked object(s):" -ForegroundColor Red
                foreach ($obj in $importOutput) {
                    Write-Host "      Type: $($obj.GetType().FullName)  Value: $obj" -ForegroundColor DarkRed
                }
            }
            $totalLeaked += $leaked
        }

        Assert-Test "ModuleImport" "No modules leak output on import" `
            -Expected 0 -Actual $totalLeaked
    }
    else {
        Write-Host "  (Skipped - module paths not found on this machine)" -ForegroundColor DarkGray
        Assert-Test "ModuleImport" "Module paths exist" `
            -Expected "True" -Actual "False (skipped)"
    }
}
else {
    Write-Host "  (Skipped - Paths.txt not found)" -ForegroundColor DarkGray
    Assert-Test "ModuleImport" "Paths.txt exists" `
        -Expected "True" -Actual "False (skipped)"
}

$canary3 = Test-FormatCanary "After module reimport"
Assert-Test "ModuleImport" "Format-Table still works after module reimport" `
    -Expected "True" -Actual $canary3

Write-Host ""

#endregion


#region ===== TEST 4: Console Cursor Manipulation =====

Write-Host "--- Test 4: Console Cursor Manipulation ---" -ForegroundColor Yellow

$isConsoleHost = $Host.Name -eq 'ConsoleHost'

if ($isConsoleHost) {
    # Test basic cursor operations
    $cursorOK = $true
    try {
        $savedTop  = [Console]::CursorTop
        $savedLeft = [Console]::CursorLeft

        # Write some content
        Write-Host "    Cursor test line 1"
        Write-Host "    Cursor test line 2"
        Write-Host "    Cursor test line 3"
        $afterWrite = [Console]::CursorTop

        # Move cursor back and overwrite
        [Console]::SetCursorPosition(0, $afterWrite - 2)
        $padWidth = [Console]::BufferWidth - 1
        [Console]::Write((" " * $padWidth))
        [Console]::SetCursorPosition(0, $afterWrite - 1)
        [Console]::Write((" " * $padWidth))

        # Park cursor at end
        [Console]::SetCursorPosition(0, $afterWrite)
        Write-Host "    (Lines 1-2 overwritten, line 3 remains)"
    }
    catch {
        $cursorOK = $false
        Write-Host "    CURSOR ERROR: $_" -ForegroundColor Red
    }

    Assert-Test "Cursor" "[Console]::SetCursorPosition works" `
        -Expected "True" -Actual $cursorOK

    # Now test the exact pattern from Invoke-RunspacePool cleanup
    $cleanupOK = $true
    try {
        $progressTop = [Console]::CursorTop
        $progressEnd = $progressTop

        # Simulate rendering a progress table
        Write-Host "    Simulated progress line 1"
        Write-Host "    Simulated progress line 2"
        $progressEnd = [Console]::CursorTop

        # Simulate the cleanup code from Invoke-RunspacePool lines 562-588
        $cursorNow = [Console]::CursorTop
        if ($cursorNow -lt $progressEnd) {
            $delta = $progressEnd - $cursorNow
            $progressTop = [math]::Max(0, $progressTop - $delta)
            $progressEnd = $cursorNow
        }
        $clearWidth = [Console]::BufferWidth - 1
        for ($clr = $progressTop; $clr -lt $progressEnd; $clr++) {
            try {
                [Console]::SetCursorPosition(0, $clr)
                [Console]::Write(" " * $clearWidth)
            } catch { break }
        }
        [Console]::SetCursorPosition(0, $progressTop)
        Write-Host "    (Progress area cleared successfully)"
    }
    catch {
        $cleanupOK = $false
        Write-Host "    CLEANUP ERROR: $_" -ForegroundColor Red
    }

    Assert-Test "Cursor" "RunspacePool cleanup pattern works" `
        -Expected "True" -Actual $cleanupOK
}
else {
    Write-Host "  (Skipped - not ConsoleHost)" -ForegroundColor DarkGray
    Assert-Test "Cursor" "Running in ConsoleHost" `
        -Expected "True" -Actual "False (skipped)"
}

$canary4 = Test-FormatCanary "After cursor manipulation"
Assert-Test "Cursor" "Format-Table still works after cursor manipulation" `
    -Expected "True" -Actual $canary4

Write-Host ""

#endregion


#region ===== TEST 5: Wide Format-Table =====

Write-Host "--- Test 5: Format-Table With 9 Columns (Invoke-Patch Schema) ---" -ForegroundColor Yellow

$sampleResults = New-SampleResults -Count 3

$displayProperties = @("IPAddress", "Computername", "Status", "SoftwareName",
    "Version", "Compliant", "NewVersion", "ExitCode", "Comment")

# Measure how wide the table would be
$rendered = $sampleResults | Select-Object $displayProperties |
    Format-Table -AutoSize | Out-String
$renderedLines = $rendered.Trim() -split "`n"

$maxLineWidth = 0
foreach ($line in $renderedLines) {
    if ($line.Length -gt $maxLineWidth) { $maxLineWidth = $line.Length }
}

$testBufferWidth = if ($bufferWidth -ne "N/A") { [int]$bufferWidth } else { [int]$rawUIWidth }
$fitsInBuffer = $maxLineWidth -le $testBufferWidth

Write-Host "  Table width: $maxLineWidth chars"
Write-Host "  Buffer width: $testBufferWidth chars"
Write-Host "  Fits in buffer: $(if ($fitsInBuffer) { 'Yes' } else { 'NO - TABLE TRUNCATED' })"

$tableHasAllColumns = $true
foreach ($prop in $displayProperties) {
    if ($rendered -notmatch $prop) {
        $tableHasAllColumns = $false
        Write-Host "    MISSING COLUMN: $prop" -ForegroundColor Red
    }
}

Assert-Test "WideTable" "Table width fits in console buffer" `
    -Expected "True" -Actual $fitsInBuffer

Assert-Test "WideTable" "All 9 columns are present in output" `
    -Expected "True" -Actual $tableHasAllColumns

# Verify no AdminName/Date leaked through Select-Object
$hasAdminName = $rendered -match 'AdminName'
$hasDate      = $rendered -match '\bDate\b'

Assert-Test "WideTable" "AdminName is NOT in display output" `
    -Expected "False" -Actual $hasAdminName

Assert-Test "WideTable" "Date is NOT in display output" `
    -Expected "False" -Actual $hasDate

$canary5 = Test-FormatCanary "After wide table test"
Assert-Test "WideTable" "Format-Table still works after wide table" `
    -Expected "True" -Actual $canary5

Write-Host ""

#endregion


#region ===== TEST 6: Bare Format-Table From Function (Original Bug) =====

Write-Host "--- Test 6: Bare Format-Table From Function (No Out-Host) ---" -ForegroundColor Yellow

# This replicates the original bug pattern: Format-Table as function output
function Test-BareFormatTable {
    [CmdletBinding()]
    param()

    begin {
        # Simulate the begin block -- these could leak
        Set-ExecutionPolicy Bypass -Scope Process -Force
    }

    process {
        $data = New-SampleResults -Count 3
        $props = @("IPAddress", "Computername", "Status", "SoftwareName",
            "Version", "Compliant", "NewVersion", "ExitCode", "Comment")

        # THE ORIGINAL BUG: Format-Table without Out-Host
        $data | Select-Object $props | Format-Table -AutoSize
    }
}

# Capture the function output to see what types come out
$bareOutput = Test-BareFormatTable

$outputTypes = @($bareOutput | ForEach-Object { $_.GetType().Name } | Sort-Object -Unique)
Write-Host "  Output object types: $($outputTypes -join ', ')"

$hasFormatObjects = ($outputTypes -match 'Format|Group').Count -gt 0
$hasNonFormatObjects = ($outputTypes | Where-Object { $_ -notmatch 'Format|Group' }).Count -gt 0

Assert-Test "BareFormat" "Function emits format objects" `
    -Expected "True" -Actual $hasFormatObjects

Assert-Test "BareFormat" "Function emits ONLY format objects (no leaks)" `
    -Expected "False" -Actual $hasNonFormatObjects

if ($hasNonFormatObjects) {
    Write-Host "    LEAKED NON-FORMAT OBJECTS:" -ForegroundColor Red
    foreach ($obj in $bareOutput) {
        $typeName = $obj.GetType().Name
        if ($typeName -notmatch 'Format|Group') {
            Write-Host "      Type: $($obj.GetType().FullName)  Value: $obj" -ForegroundColor DarkRed
        }
    }
}

$canary6 = Test-FormatCanary "After bare Format-Table function"
Assert-Test "BareFormat" "Format-Table still works after bare function call" `
    -Expected "True" -Actual $canary6

Write-Host ""

#endregion


#region ===== TEST 7: Format-Table With Out-Host From Function (Fixed) =====

Write-Host "--- Test 7: Format-Table With Out-Host From Function (Fixed) ---" -ForegroundColor Yellow

function Test-FixedFormatTable {
    [CmdletBinding()]
    param()

    begin {
        Set-ExecutionPolicy Bypass -Scope Process -Force *> $null
    }

    process {
        $data = New-SampleResults -Count 3
        $props = @("IPAddress", "Computername", "Status", "SoftwareName",
            "Version", "Compliant", "NewVersion", "ExitCode", "Comment")

        # THE FIX: Format-Table piped to Out-Host
        $data | Select-Object $props | Format-Table -AutoSize | Out-Host
    }
}

# Capture -- should be empty since Out-Host sends directly to console
$fixedOutput = Test-FixedFormatTable

$fixedLeaked = if ($null -ne $fixedOutput) { @($fixedOutput).Count } else { 0 }

Assert-Test "FixedFormat" "Fixed function leaks nothing to pipeline" `
    -Expected 0 -Actual $fixedLeaked

$canary7 = Test-FormatCanary "After fixed Format-Table function"
Assert-Test "FixedFormat" "Format-Table still works after fixed function call" `
    -Expected "True" -Actual $canary7

Write-Host ""

#endregion


#region ===== TEST 8: Invoke-RunspacePool Progress Simulation =====

Write-Host "--- Test 8: Progress Display Simulation (Cursor Manipulation) ---" -ForegroundColor Yellow

if ($isConsoleHost) {
    # Simulate what Invoke-RunspacePool does during monitoring
    $simOK = $true
    try {
        $progressTop = [Console]::CursorTop
        $progressEnd = $progressTop
        $padWidth    = [Console]::BufferWidth - 1

        # Simulate 3 render cycles
        for ($cycle = 1; $cycle -le 3; $cycle++) {
            [Console]::CursorVisible = $false
            [Console]::SetCursorPosition(0, $progressTop)

            # Header lines (same as Invoke-RunspacePool)
            Write-Host (' ' * $padWidth)
            Write-Host "Waiting on " -NoNewline
            Write-Host "TestSoftware " -ForegroundColor Magenta -NoNewline
            Write-Host ("tasks...    Timeout: 30 min    (Progress: $cycle/3)".PadRight($padWidth - 60))
            Write-Host ("Queued: 0  |  Active: $(3 - $cycle)  |  Completed: $cycle  |  Failed: 0".PadRight($padWidth)) -ForegroundColor DarkGray
            Write-Host (' ' * $padWidth)

            # Simulate a mini progress table using Format-Table | Out-String
            $progressData = @(
                [PSCustomObject]@{ Computer = "TESTPC-0001"; Elapsed = "12s"; Status = "Patching" }
                [PSCustomObject]@{ Computer = "TESTPC-0002"; Elapsed = "8s";  Status = "Copying" }
            )
            $tableLines = (($progressData | Format-Table -AutoSize | Out-String).TrimEnd()) -split "`n"
            foreach ($tl in $tableLines) {
                Write-Host ($tl.TrimEnd().PadRight($padWidth))
            }

            $progressEnd = [Console]::CursorTop
            [Console]::CursorVisible = $true

            Start-Sleep -Milliseconds 200
        }

        # Simulate cleanup (Invoke-RunspacePool lines 562-588)
        $cursorNow = [Console]::CursorTop
        $clearWidth = [Console]::BufferWidth - 1
        for ($clr = $progressTop; $clr -lt $progressEnd; $clr++) {
            try {
                [Console]::SetCursorPosition(0, $clr)
                [Console]::Write(" " * $clearWidth)
            } catch { break }
        }
        [Console]::SetCursorPosition(0, $progressTop)

        # Completion message
        Write-Host ""
        Write-Host "TestSoftware " -ForegroundColor Magenta -NoNewline
        Write-Host "tasks complete!  3 Completed  0 Failed"
        Write-Host ""
    }
    catch {
        $simOK = $false
        Write-Host "    SIMULATION ERROR: $_" -ForegroundColor Red
    }

    Assert-Test "ProgressSim" "Progress display simulation completed" `
        -Expected "True" -Actual $simOK

    # Now try rendering a Format-Table AFTER the progress cleanup
    # This is exactly what Invoke-Patch does
    $postProgressOK = $true
    try {
        $data = New-SampleResults -Count 3
        $props = @("IPAddress", "Computername", "Status", "SoftwareName",
            "Version", "Compliant", "NewVersion", "ExitCode", "Comment")

        $rendered = $data | Select-Object $props | Format-Table -AutoSize | Out-String
        $hasHeader = $rendered -match 'IPAddress'
        $hasData   = $rendered -match 'TESTPC'

        if (-not $hasHeader -or -not $hasData) {
            $postProgressOK = $false
            Write-Host "    TABLE AFTER PROGRESS IS BROKEN:" -ForegroundColor Red
            Write-Host "    $rendered" -ForegroundColor DarkRed
        }
    }
    catch {
        $postProgressOK = $false
        Write-Host "    POST-PROGRESS TABLE ERROR: $_" -ForegroundColor Red
    }

    Assert-Test "ProgressSim" "Format-Table renders correctly after progress cleanup" `
        -Expected "True" -Actual $postProgressOK
}
else {
    Write-Host "  (Skipped - not ConsoleHost)" -ForegroundColor DarkGray
    Assert-Test "ProgressSim" "Running in ConsoleHost" `
        -Expected "True" -Actual "False (skipped)"
}

$canary8 = Test-FormatCanary "After progress simulation"
Assert-Test "ProgressSim" "Format-Table still works after full simulation" `
    -Expected "True" -Actual $canary8

Write-Host ""

#endregion


#region ===== TEST 9: Full Mini-Pipeline (Invoke-Patch Simulation) =====

Write-Host "--- Test 9: Full Invoke-Patch Mini-Pipeline ---" -ForegroundColor Yellow

# This function replicates the full Invoke-Patch flow at a high level:
# begin: Set-ExecutionPolicy, Main-Switch load, variable setup
# process: Invoke-RunspacePool (with progress), result collection, Format-Table

function Test-FullPipeline {
    [CmdletBinding()]
    param([switch]$UseOutHost)

    begin {
        # Simulate begin block
        Set-ExecutionPolicy Bypass -Scope Process -Force *> $null
    }

    process {
        $isConsole = $Host.Name -eq 'ConsoleHost'

        # --- Simulate Invoke-RunspacePool with progress display ---
        if ($isConsole) {
            $pTop = [Console]::CursorTop
            $pEnd = $pTop
            $pw   = [Console]::BufferWidth - 1

            [Console]::CursorVisible = $false
            [Console]::SetCursorPosition(0, $pTop)
            Write-Host (' ' * $pw)
            Write-Host "Waiting on " -NoNewline
            Write-Host "Chrome " -ForegroundColor Magenta -NoNewline
            Write-Host ("tasks...    Timeout: 30 min    (Progress: 0/3)".PadRight($pw - 50))
            Write-Host ("Queued: 3  |  Active: 0  |  Completed: 0  |  Failed: 0".PadRight($pw)) -ForegroundColor DarkGray
            $pEnd = [Console]::CursorTop
            [Console]::CursorVisible = $true

            Start-Sleep -Milliseconds 300

            # Cleanup
            $cw = [Console]::BufferWidth - 1
            for ($i = $pTop; $i -lt $pEnd; $i++) {
                try { [Console]::SetCursorPosition(0, $i); [Console]::Write(' ' * $cw) } catch { break }
            }
            [Console]::SetCursorPosition(0, $pTop)

            Write-Host ""
            Write-Host "Chrome " -ForegroundColor Magenta -NoNewline
            Write-Host "tasks complete!  3 Completed  0 Failed"
            Write-Host ""
        }

        # --- Simulate result collection (like Invoke-RunspacePool emitting objects) ---
        $pipelineResults = New-SampleResults -Count 3

        # --- Simulate Invoke-Patch result processing ---
        $results = @($pipelineResults | Where-Object {
            $_ -is [PSCustomObject] -and $null -ne $_.PSObject.Properties['SoftwareName']
        })

        $displayResults = $results   # Skip Add-Delimiter for simplicity

        $displayProperties = @("IPAddress", "Computername", "Status", "SoftwareName",
            "Version", "Compliant", "NewVersion", "ExitCode", "Comment")

        # --- Output ---
        if ($UseOutHost) {
            $displayResults | Select-Object $displayProperties | Sort-Object -Property (
                @{Expression = "Status"; Descending = $true },
                @{Expression = "Version"; Descending = $false }
            ) | Format-Table -AutoSize | Out-Host
        }
        else {
            $displayResults | Select-Object $displayProperties | Sort-Object -Property (
                @{Expression = "Status"; Descending = $true },
                @{Expression = "Version"; Descending = $false }
            ) | Format-Table -AutoSize
        }
    }
}

# --- Test WITHOUT Out-Host (original bug pattern) ---
Write-Host "  Sub-test A: Full pipeline WITHOUT Out-Host" -ForegroundColor DarkYellow
$outputA = Test-FullPipeline

$typesA = @()
if ($null -ne $outputA) {
    $typesA = @($outputA | ForEach-Object { $_.GetType().Name } | Sort-Object -Unique)
}
Write-Host "  Pipeline output types: $(if ($typesA.Count -eq 0) { '(none)' } else { $typesA -join ', ' })"

$nonFormatA = ($typesA | Where-Object { $_ -notmatch 'Format|Group' }).Count -gt 0
Assert-Test "FullPipeline" "No non-format objects leak (without Out-Host)" `
    -Expected "False" -Actual $nonFormatA

if ($nonFormatA) {
    Write-Host "    LEAKED OBJECTS:" -ForegroundColor Red
    foreach ($obj in $outputA) {
        $tn = $obj.GetType().Name
        if ($tn -notmatch 'Format|Group') {
            Write-Host "      Type: $($obj.GetType().FullName)" -ForegroundColor DarkRed
            if ($obj -is [PSCustomObject]) {
                $obj.PSObject.Properties | ForEach-Object {
                    Write-Host "        $($_.Name) = $($_.Value)" -ForegroundColor DarkRed
                }
            }
            else {
                Write-Host "        Value: $obj" -ForegroundColor DarkRed
            }
        }
    }
}

$canary9a = Test-FormatCanary "After full pipeline (no Out-Host)"
Assert-Test "FullPipeline" "Format-Table works after pipeline without Out-Host" `
    -Expected "True" -Actual $canary9a

# --- Test WITH Out-Host (fixed pattern) ---
Write-Host ""
Write-Host "  Sub-test B: Full pipeline WITH Out-Host" -ForegroundColor DarkYellow
$outputB = Test-FullPipeline -UseOutHost

$leakedB = if ($null -ne $outputB) { @($outputB).Count } else { 0 }

Assert-Test "FullPipeline" "Fixed pipeline leaks nothing" `
    -Expected 0 -Actual $leakedB

$canary9b = Test-FormatCanary "After full pipeline (with Out-Host)"
Assert-Test "FullPipeline" "Format-Table works after fixed pipeline" `
    -Expected "True" -Actual $canary9b

Write-Host ""

#endregion



#region ===== TEST 10: Bare RunspacePool Lifecycle =====

Write-Host "--- Test 10: Bare .NET RunspacePool Lifecycle ---" -ForegroundColor Yellow
Write-Host "  Creates a raw RunspacePool with `$Host, runs trivial work, disposes." -ForegroundColor Gray

$poolTestOK = $true
try {
    $iss10 = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $pool10 = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(
        1, 2, $iss10, $Host
    )
    $pool10.Open()

    # Run 3 trivial scriptblocks
    $handles10 = @()
    $psInstances10 = @()
    for ($i = 0; $i -lt 3; $i++) {
        $ps10 = [PowerShell]::Create()
        $ps10.RunspacePool = $pool10
        $ps10.AddScript('Start-Sleep -Milliseconds 200; [PSCustomObject]@{ Result = "OK" }') > $null
        $handles10 += [PSCustomObject]@{
            PS     = $ps10
            Handle = $ps10.BeginInvoke()
        }
        $psInstances10 += $ps10
    }

    # Collect results
    foreach ($h in $handles10) {
        try { $null = $h.PS.EndInvoke($h.Handle) } catch {}
    }

    # Dispose PowerShell instances
    foreach ($ps10 in $psInstances10) {
        try { $ps10.Dispose() } catch {}
    }

    # Dispose pool
    try { $pool10.Close() } catch {}
    try { $pool10.Dispose() } catch {}

    Write-Host "  Pool created, used, and disposed successfully." -ForegroundColor Gray
}
catch {
    $poolTestOK = $false
    Write-Host "  POOL ERROR: $_" -ForegroundColor Red
}

Assert-Test "BarePool" "Bare pool lifecycle completed without error" `
    -Expected "True" -Actual $poolTestOK

# Run ALL canaries after pool disposal
$healthy10 = Test-AllCanaries "BarePool"

Write-Host ""

#endregion


#region ===== TEST 11: Real Invoke-RunspacePool With Local Scriptblocks =====

Write-Host "--- Test 11: Real Invoke-RunspacePool (Local Scriptblocks) ---" -ForegroundColor Yellow
Write-Host "  Calls the actual module with local-only work. No network access." -ForegroundColor Gray

# Check if Invoke-RunspacePool is available
$irpAvailable = $null -ne (Get-Command Invoke-RunspacePool -ErrorAction SilentlyContinue)

if ($irpAvailable) {
    $irpResults = $null
    $irpOK = $true
    try {
        $irpScript = {
            $computer = $args[0]
            Start-Sleep -Milliseconds 300
            [PSCustomObject]@{
                ComputerName = $computer
                Status       = "Online"
                SoftwareName = "TestSoftware"
                Version      = "1.0.0.1"
                Comment      = "Local test"
            }
        }

        $irpArgs = @(
            ,@("LOCAL-PC-001")
            ,@("LOCAL-PC-002")
            ,@("LOCAL-PC-003")
        )

        $irpResults = Invoke-RunspacePool -ScriptBlock $irpScript `
            -ArgumentList $irpArgs -ThrottleLimit 3 `
            -TimeoutMinutes 1 -ActivityName "DiagTest11"
    }
    catch {
        $irpOK = $false
        Write-Host "  INVOKE-RUNSPACEPOOL ERROR: $_" -ForegroundColor Red
    }

    $resultCount = if ($null -ne $irpResults) { @($irpResults).Count } else { 0 }

    Assert-Test "RealIRP" "Invoke-RunspacePool completed without error" `
        -Expected "True" -Actual $irpOK

    Assert-Test "RealIRP" "Returned expected result count" `
        -Expected 3 -Actual $resultCount

    # Run ALL canaries after real Invoke-RunspacePool
    $healthy11 = Test-AllCanaries "RealIRP"
}
else {
    Write-Host "  (Skipped - Invoke-RunspacePool module not loaded)" -ForegroundColor DarkGray
    Assert-Test "RealIRP" "Invoke-RunspacePool available" `
        -Expected "True" -Actual "False (skipped)"
}

Write-Host ""

#endregion


#region ===== TEST 12: Full Invoke-Patch Pattern (RunspacePool + Format-Table) =====

Write-Host "--- Test 12: Full Invoke-Patch Pattern (RunspacePool + Format-Table) ---" -ForegroundColor Yellow
Write-Host "  Calls Invoke-RunspacePool then pipes results through Format-Table inside a function." -ForegroundColor Gray

function Test-FullPatchPattern {
    [CmdletBinding()]
    param()

    begin {
        Set-ExecutionPolicy Bypass -Scope Process -Force *> $null
    }

    process {
        # --- Call real Invoke-RunspacePool ---
        $irpScript = {
            $computer = $args[0]
            Start-Sleep -Milliseconds 200
            [PSCustomObject]@{
                IPAddress    = "10.0.0.1"
                ComputerName = $computer
                Status       = "Online"
                SoftwareName = "TestSoftware"
                Version      = "1.0.0.1"
                Compliant    = "Yes"
                NewVersion   = "2.0.0.1"
                ExitCode     = "0"
                Comment      = "Test"
                AdminName    = "test.admin"
                Date         = "2026/04/09 12:00"
            }
        }

        $irpArgs = @(
            ,@("PATCH-PC-001")
            ,@("PATCH-PC-002")
            ,@("PATCH-PC-003")
        )

        $pipelineResults = Invoke-RunspacePool -ScriptBlock $irpScript `
            -ArgumentList $irpArgs -ThrottleLimit 3 `
            -TimeoutMinutes 1 -ActivityName "DiagTest12"

        # --- Replicate Invoke-Patch result processing ---
        $results = @($pipelineResults | Where-Object {
            $_ -is [PSCustomObject] -and $null -ne $_.PSObject.Properties['SoftwareName']
        })

        $displayProperties = @("IPAddress", "Computername", "Status", "SoftwareName",
            "Version", "Compliant", "NewVersion", "ExitCode", "Comment")

        # --- THE EXACT PATTERN FROM INVOKE-PATCH ---
        $results | Select-Object $displayProperties | Sort-Object -Property (
            @{Expression = "Status"; Descending = $true },
            @{Expression = "Version"; Descending = $false }
        ) | Format-Table -AutoSize | Out-Host
    }
}

if ($irpAvailable) {
    $patchPatternOK = $true
    try {
        Test-FullPatchPattern
    }
    catch {
        $patchPatternOK = $false
        Write-Host "  PATCH PATTERN ERROR: $_" -ForegroundColor Red
    }

    Assert-Test "PatchPattern" "Full patch pattern completed without error" `
        -Expected "True" -Actual $patchPatternOK

    # Run ALL canaries after the full pattern
    $healthy12 = Test-AllCanaries "PatchPattern"
}
else {
    Write-Host "  (Skipped - Invoke-RunspacePool module not loaded)" -ForegroundColor DarkGray
    Assert-Test "PatchPattern" "Invoke-RunspacePool available" `
        -Expected "True" -Actual "False (skipped)"
}

Write-Host ""

#endregion


#region ===== TEST 13: Output Layer Diagnosis =====

Write-Host "--- Test 13: Output Layer Diagnosis (Final State) ---" -ForegroundColor Yellow
Write-Host "  Tests each output layer independently to pinpoint corruption." -ForegroundColor Gray

# Layer 1: Write-Host (bypasses everything, talks directly to host UI)
$layer1 = Test-WriteHostCanary "Layer diagnosis"
Assert-Test "LayerDiag" "Layer 1 - Write-Host (Host UI)" `
    -Expected "True" -Actual $layer1

# Layer 2: Out-String (format + capture, no display)
$layer2 = Test-OutHostCanary "Layer diagnosis"
Assert-Test "LayerDiag" "Layer 2 - Out-String (formatting engine)" `
    -Expected "True" -Actual $layer2

# Layer 3: Pipeline capture in scriptblock
$layer3 = Test-OutputCanary "Layer diagnosis"
Assert-Test "LayerDiag" "Layer 3 - Pipeline capture (output stream)" `
    -Expected "True" -Actual $layer3

# Layer 4: Format-Table specifically
$layer4 = Test-FormatCanary "Layer diagnosis"
Assert-Test "LayerDiag" "Layer 4 - Format-Table (table formatter)" `
    -Expected "True" -Actual $layer4

# Layer 5: Out-Default (the final output path for interactive display)
$layer5OK = $true
try {
    # Out-Default is what makes "$var" display at the prompt.
    # We test it by sending a known string through Out-Default and checking
    # if the formatter processes it (via Out-String as a proxy).
    $testObj = [PSCustomObject]@{ DiagKey = "DiagValue_$(Get-Random)" }
    $rendered5 = ($testObj | Format-List | Out-String).Trim()
    $layer5OK = $rendered5 -match 'DiagValue_'

    if (-not $layer5OK) {
        Write-Host "    Out-Default proxy test returned: '$rendered5'" -ForegroundColor DarkRed
    }
}
catch {
    $layer5OK = $false
    Write-Host "    LAYER 5 ERROR: $_" -ForegroundColor Red
}
Assert-Test "LayerDiag" "Layer 5 - Format-List + Out-String (default formatter)" `
    -Expected "True" -Actual $layer5OK

# Summary diagnosis
if ($layer1 -and $layer2 -and $layer3 -and $layer4 -and $layer5OK) {
    Write-Host "  All output layers are healthy." -ForegroundColor Green
}
else {
    Write-Host ""
    Write-Host "  CORRUPTION DETECTED - Layer Analysis:" -ForegroundColor Red
    if (-not $layer1) {
        Write-Host "  >> Host UI is DEAD. `$Host.UI.WriteLine is broken." -ForegroundColor Red
        Write-Host "     Cause: RunspacePool disposal corrupted the shared `$Host object." -ForegroundColor DarkRed
        Write-Host "     Fix: Probe [Console] access before pool creation; fall back to Write-Progress." -ForegroundColor DarkYellow
    }
    elseif (-not $layer3) {
        Write-Host "  >> Output STREAM is DEAD. Pipeline produces nothing." -ForegroundColor Red
        Write-Host "     Cause: An unclosed format operation is swallowing all output." -ForegroundColor DarkRed
        Write-Host "     Fix: Ensure Format-Table always completes (pipe to Out-Host)." -ForegroundColor DarkYellow
    }
    elseif (-not $layer2 -or -not $layer5OK) {
        Write-Host "  >> FORMATTER is stuck. Objects enter but never render." -ForegroundColor Red
        Write-Host "     Cause: FormatStartData emitted without FormatEndData." -ForegroundColor DarkRed
        Write-Host "     Fix: Pipe all Format-* calls to Out-Host to isolate the formatter." -ForegroundColor DarkYellow
    }
    elseif (-not $layer4) {
        Write-Host "  >> Format-TABLE specifically is broken, other output works." -ForegroundColor Red
        Write-Host "     Cause: Table column calculation or AutoSize failed." -ForegroundColor DarkRed
        Write-Host "     Fix: Check console buffer width; consider dropping -AutoSize." -ForegroundColor DarkYellow
    }
}

Write-Host ""

#endregion



# ================================================================
#region ===== SUMMARY =====

Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "  RESULTS SUMMARY" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host ""

$passed = @($script:TestResults | Where-Object { $_.Status -eq 'PASS' }).Count
$failed = @($script:TestResults | Where-Object { $_.Status -eq 'FAIL' }).Count
$total  = $script:TestResults.Count

if ($failed -eq 0) {
    Write-Host "  All $total tests PASSED" -ForegroundColor Green
}
else {
    Write-Host "  $passed/$total passed, $failed FAILED:" -ForegroundColor Red
    Write-Host ""
    $script:TestResults | Where-Object { $_.Status -eq 'FAIL' } | ForEach-Object {
        Write-Host "    FAIL #$($_.Number): [$($_.Test)] $($_.Description)" -ForegroundColor Red
        Write-Host "          Expected: $($_.Expected)  |  Actual: $($_.Actual)" -ForegroundColor DarkRed
    }
}

Write-Host ""
Write-Host "--- Final Canary Check (All Layers) ---" -ForegroundColor Yellow
$finalAll = Test-AllCanaries "Final"
if ($finalAll) {
    Write-Host "  All output layers are HEALTHY after all tests." -ForegroundColor Green
}
else {
    Write-Host "  OUTPUT IS CORRUPTED! Check the test results above." -ForegroundColor Red
    Write-Host "  The first failed canary identifies the guilty operation." -ForegroundColor Red
    Write-Host "  The Layer Diagnosis (Test 13) identifies WHAT is broken." -ForegroundColor Red
}

Write-Host ""
Write-Host "--- Post-Test Verification ---" -ForegroundColor Yellow
Write-Host '  Run these commands now to verify output still works:' -ForegroundColor Gray
Write-Host '    $PSVersionTable              (table output)' -ForegroundColor Gray
Write-Host '    $env:COMPUTERNAME             (bare output)' -ForegroundColor Gray
Write-Host '    $x = "test" ; $x             (variable display)' -ForegroundColor Gray
Write-Host ""

#endregion
