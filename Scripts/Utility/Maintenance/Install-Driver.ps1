# DOTS formatting comment

function Install-Driver {
    <#
        .SYNOPSIS
            Installs a driver package on remote machines via pnputil.
        .DESCRIPTION
            Copies a local driver folder to each remote machine's C:\Temp directory,
            then runs pnputil /add-driver to install all .inf files found in that
            folder (including subdirectories).

            Uses Invoke-RunspacePool for parallel execution across many machines.
        .PARAMETER ComputerName
            One or more target machine names. Accepts pipeline input.
        .PARAMETER DriverPath
            Local path to the driver folder containing .inf files to install.
            The folder will be copied to C:\Temp on each remote machine.
        .PARAMETER ThrottleLimit
            Maximum concurrent runspaces. Default 50.
        .PARAMETER Timeout
            Timeout in minutes before a runspace is stopped. Default 15.
        .EXAMPLE
            $list = Get-Content "$env:USERPROFILE\Desktop\Lists\Target_Machines.txt"
            Install-Driver -ComputerName $list -DriverPath "C:\Patches\Intel-Dynamic-Tuning-Driver"
        .EXAMPLE
            Install-Driver -ComputerName "PC001" -DriverPath "\\share\drivers\NIC-v2"
        .NOTES
            Written by Skyler Werner
            Version: 2.0.0
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [string[]]
        $ComputerName,

        [Parameter(Mandatory, Position = 1)]
        [ValidateScript({ Test-Path $_ -PathType Container })]
        [string]
        $DriverPath,

        [Parameter()]
        [ValidateRange(1, 300)]
        [int]
        $ThrottleLimit = 50,

        [Parameter()]
        [Alias("TimeoutMinutes")]
        [ValidateRange(1, 120)]
        [int]
        $Timeout = 15,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        $collectedNames = @()
    }

    process {
        foreach ($name in $ComputerName) {
            if ($name.Length -gt 0) {
                $collectedNames += $name
            }
        }
    }

    end {

    # --- Input cleanup ---
    $targets = @(Format-ComputerList $collectedNames -ToUpper)
    if ($targets.Count -eq 0) {
        Write-Warning "No valid computer names provided."
        return
    }

$driverFolder = Split-Path $DriverPath -Leaf


# --- Connectivity check ---
Write-Host ""
Write-Host "Checking for online machines..."

$pingResults = Test-ConnectionAsJob -ComputerName $targets

$onlineList  = @()
$allResults  = @()

foreach ($ping in $pingResults) {
    if ($ping.Reachable -eq $true) {
        $onlineList += $ping.ComputerName
    }
    else {
        $allResults += [PSCustomObject][ordered]@{
            ComputerName    = $ping.ComputerName
            Status          = "Offline"
            DriverInstalled = $null
            ExitCode        = $null
            Comment         = $null
        }
    }
}

if ($onlineList.Count -eq 0) {
    Write-Host "No online machines found." -ForegroundColor Yellow
    $allResults | Sort-Object ComputerName | Format-Table -AutoSize | Out-Host
    if ($PassThru) { return ($allResults | Sort-Object ComputerName) }
    return
}


# --- Capture WhatIf state for passing into runspaces ---
$isWhatIf = $WhatIfPreference


# --- Scriptblock executed in each runspace (one per machine) ---
$installDriverBlock = {

    $computer     = $args[0]
    $srcPath      = $args[1]
    $folderName   = $args[2]
    $dryRun       = $args[3]

    if ($dryRun) {
        $PhaseTracker[$computer] = "WhatIf"
        return [PSCustomObject][ordered]@{
            ComputerName    = $computer
            Status          = "Online"
            DriverInstalled = 'WhatIf'
            ExitCode        = $null
            Comment         = "Would install driver '$folderName'"
        }
    }

    $PhaseTracker[$computer] = "Copying Driver"

    try {
        # --- Copy driver folder to remote machine ---
        $remoteTempPath = "\\$computer\C$\Temp"
        $remoteDestPath = "$remoteTempPath\$folderName"

        if (-not (Test-Path $remoteTempPath)) {
            New-Item -Path $remoteTempPath -ItemType Directory -Force | Out-Null
        }

        if (Test-Path $remoteDestPath) {
            Remove-Item -Path $remoteDestPath -Recurse -Force
        }

        Copy-Item -Path $srcPath -Destination $remoteDestPath -Recurse -Force

        # --- Install the driver remotely ---
        $PhaseTracker[$computer] = "Installing Driver"

        $remoteResult = Invoke-Command -ComputerName $computer -ErrorAction Stop -ArgumentList $folderName -ScriptBlock {
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseUsingScopeModifierInNewRunspaces', '')]
            param($Folder)

            $driverDir = "C:\Temp\$Folder"

            if (-not (Test-Path $driverDir)) {
                return [PSCustomObject]@{
                    Installed = $false
                    ExitCode  = $null
                    Comment   = "Driver folder not found at $driverDir"
                }
            }

            $infFiles = @(Get-ChildItem $driverDir -Filter "*.inf" -Recurse)
            if ($infFiles.Count -eq 0) {
                return [PSCustomObject]@{
                    Installed = $false
                    ExitCode  = $null
                    Comment   = "No .inf files found in $driverDir"
                }
            }

            $output = & pnputil.exe /add-driver "$driverDir\*.inf" /subdirs /install 2>&1
            $exitCode = $LASTEXITCODE

            $comment = $null
            if ($exitCode -eq 3010) { $comment = "Reboot required" }

            [PSCustomObject]@{
                Installed = ($exitCode -eq 0 -or $exitCode -eq 3010)
                ExitCode  = $exitCode
                Comment   = $comment
            }
        }

        [PSCustomObject][ordered]@{
            ComputerName    = $computer
            Status          = "Online"
            DriverInstalled = $remoteResult.Installed
            ExitCode        = $remoteResult.ExitCode
            Comment         = $remoteResult.Comment
        }
    }
    catch {
        [PSCustomObject][ordered]@{
            ComputerName    = $computer
            Status          = "Online"
            DriverInstalled = $false
            ExitCode        = $null
            Comment         = "Failed: $_"
        }
    }
}


# --- Build argument sets (one per machine) ---
$argumentSets = @(
    foreach ($machine in $onlineList) {
        , @($machine, $DriverPath, $driverFolder, $isWhatIf)
    }
)


# --- Execute via Invoke-RunspacePool ---
if ($isWhatIf) {
    Write-Host "WhatIf: Would install driver '$driverFolder' on $($onlineList.Count) machine(s)..."
}

$runspaceParams = @{
    ScriptBlock    = $installDriverBlock
    ArgumentList   = $argumentSets
    ThrottleLimit  = $ThrottleLimit
    TimeoutMinutes = $Timeout
    ActivityName   = "Install Driver"
}

$poolResults = Invoke-RunspacePool @runspaceParams


# --- Post-processing ---
foreach ($result in $poolResults) {
    if ($result -isnot [PSCustomObject]) { continue }

    if ($null -ne $result.PSObject.Properties['Status']) {
        $allResults += $result
        continue
    }

    $allResults += [PSCustomObject][ordered]@{
        ComputerName    = $result.ComputerName
        Status          = "Online"
        DriverInstalled = $null
        ExitCode        = $null
        Comment         = if ($null -ne $result.Comment) { $result.Comment } else { "Task Stopped" }
    }
}


# --- Output results ---
    $sorted = $allResults | Sort-Object -Property @(
        @{ Expression = "Status";          Descending = $true  }
        @{ Expression = "DriverInstalled"; Descending = $true  }
        @{ Expression = "Comment";         Descending = $true  }
        @{ Expression = "ComputerName";    Descending = $false }
    )

    $sorted | Format-Table -AutoSize | Out-Host
    if ($PassThru) { return $sorted }

    } # end
}
