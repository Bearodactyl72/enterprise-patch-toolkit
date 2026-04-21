# DOTS formatting comment

<#
    .SYNOPSIS
        Reconstructs a project tree from an export created by Export-Package.ps1.
    .DESCRIPTION
        Reads _MANIFEST.txt from an export folder, recreates the original directory structure,
        copies files to their correct locations, strips the .txt extension, and unblocks them.

        Features:
            -WhatIf   Dry-run mode. Shows what would be created/overwritten without making changes.
            -Undo     Restores the most recent pre-import backup, reversing the last import.

        Before overwriting any existing file, a backup is saved to _PreImport_Backup_<date>/
        inside the destination folder. The backup manifest tracks every change so -Undo can
        reverse the operation cleanly.

        Written by Skyler Werner
        Date: 2026/03/04
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    # Reverse the most recent import using the backup folder
    [switch]$Undo
)

[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')



#region --- GUI Helpers ---


function Show-FolderPicker {
    param(
        [string]$Description,
        [string]$DefaultPath,
        [switch]$ShowNewFolderButton
    )

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.RootFolder  = 'MyComputer'
    $dialog.Description = $Description
    $dialog.ShowNewFolderButton = $ShowNewFolderButton.IsPresent

    if ($DefaultPath) {
        $dialog.SelectedPath = $DefaultPath
    }

    if ($dialog.ShowDialog() -eq 'Cancel') {
        return $null
    }

    return $dialog.SelectedPath
}


#endregion --- GUI Helpers ---



#region --- Undo Mode ---


if ($Undo.IsPresent) {

    $targetFolder = Show-FolderPicker `
        -Description "Select the folder where the import was applied (contains _PreImport_Backup_ folder)."

    if (-not $targetFolder) {
        Write-Host "No folder selected. Undo cancelled." -ForegroundColor Yellow
        return
    }

    # Find most recent backup
    $backups = Get-ChildItem $targetFolder -Directory -Filter '_PreImport_Backup_*' |
        Sort-Object Name -Descending

    if ($backups.Count -eq 0) {
        Write-Host "No backup folders found in $targetFolder" -ForegroundColor Red
        return
    }

    $latestBackup   = $backups[0]
    $backupManifest = Join-Path $latestBackup.FullName '_BACKUP_MANIFEST.txt'

    if (-not (Test-Path $backupManifest)) {
        Write-Host "Backup manifest not found in $($latestBackup.FullName)" -ForegroundColor Red
        return
    }

    Write-Host ""
    Write-Host "Restoring from: $($latestBackup.Name)" -ForegroundColor Cyan
    Write-Host ""

    $entries       = Import-Csv $backupManifest -Delimiter "`t"
    $restoredCount = 0

    foreach ($entry in $entries) {
        $originalPath = Join-Path $targetFolder $entry.RelativePath

        if ($entry.Action -eq 'Overwritten') {
            $backupFile = Join-Path $latestBackup.FullName $entry.BackupName

            if (Test-Path $backupFile) {
                # Ensure parent directory exists (in case it was removed)
                $parentDir = Split-Path $originalPath -Parent
                if (-not (Test-Path $parentDir)) {
                    New-Item -ItemType Directory -Force -Path $parentDir | Out-Null
                }

                Move-Item $backupFile -Destination $originalPath -Force
                $restoredCount++
            }
            else {
                Write-Host "  Backup file missing: $($entry.BackupName)" -ForegroundColor Red
            }
        }
        elseif ($entry.Action -eq 'Created') {
            # File was newly created by the import -- remove it to undo
            if (Test-Path $originalPath) {
                Remove-Item $originalPath -Force
                $restoredCount++
            }
        }
    }

    # Remove empty directories left behind (bottom-up)
    Get-ChildItem $targetFolder -Directory -Recurse |
        Sort-Object { $_.FullName.Length } -Descending |
        Where-Object {
            $_.Name -notmatch '^_PreImport_Backup_' -and
            (Get-ChildItem $_.FullName -Force).Count -eq 0
        } |
        ForEach-Object { Remove-Item $_.FullName -Force }

    # Clean up the backup folder itself
    Remove-Item $latestBackup.FullName -Recurse -Force

    Write-Host "Undo complete! $restoredCount files restored." -ForegroundColor Green
    Write-Host "Backup folder removed: $($latestBackup.Name)" -ForegroundColor Gray
    Write-Host ""
    return
}


#endregion --- Undo Mode ---



#region --- Source & Destination Selection ---


# Pick the export folder (must contain _MANIFEST.txt)
$sourceFolder = Show-FolderPicker `
    -Description "Select the export folder (contains _MANIFEST.txt)."

if (-not $sourceFolder) {
    [void][Microsoft.VisualBasic.Interaction]::MsgBox(
        "No source folder selected. Import cancelled.",
        "OKOnly,SystemModal,Exclamation,DefaultButton1",
        " Import Cancelled"
    )
    return
}

$manifestPath = Join-Path $sourceFolder '_MANIFEST.txt'
if (-not (Test-Path $manifestPath)) {
    [void][Microsoft.VisualBasic.Interaction]::MsgBox(
        "_MANIFEST.txt not found in the selected folder.`nPlease select a valid export folder.",
        "OKOnly,SystemModal,Exclamation,DefaultButton1",
        " Manifest Not Found"
    )
    return
}


# Pick destination folder
$destFolder = Show-FolderPicker `
    -Description "Select or create the destination folder to reconstruct the project." `
    -ShowNewFolderButton

if (-not $destFolder) {
    [void][Microsoft.VisualBasic.Interaction]::MsgBox(
        "No destination folder selected. Import cancelled.",
        "OKOnly,SystemModal,Exclamation,DefaultButton1",
        " Import Cancelled"
    )
    return
}


#endregion --- Source & Destination Selection ---



#region --- Read Manifest ---


$manifestLines = Get-Content $manifestPath | Select-Object -Skip 1

# Patterns to skip during import (git/Claude Code metadata)
$skipPatterns = @(
    '^\.git[\\/]',
    '^\.claude[\\/]',
    '^CLAUDE\.md$',
    '^\.gitignore$',
    '^\.gitattributes$'
)

$entries = foreach ($line in $manifestLines) {
    $parts = $line -split "`t"

    if ($parts.Count -ge 2 -and $parts[0].Length -gt 0) {
        $originalPath = $parts[1]
        $skip = $false
        foreach ($pattern in $skipPatterns) {
            if ($originalPath -match $pattern) { $skip = $true; break }
        }
        if (-not $skip) {
            [PSCustomObject]@{
                EncodedName  = $parts[0]
                OriginalPath = $originalPath
            }
        }
    }
}

if ($null -eq $entries -or @($entries).Count -eq 0) {
    Write-Host "Manifest is empty or could not be parsed." -ForegroundColor Red
    return
}

# Normalize to array
$entries = @($entries)


#endregion --- Read Manifest ---



#region --- WhatIf (Dry Run) ---


if ($WhatIfPreference) {

    Write-Host ""
    Write-Host "=== DRY RUN ===" -ForegroundColor Cyan
    Write-Host "Would import $($entries.Count) manifest entries to: $destFolder" -ForegroundColor Cyan
    Write-Host ""

    $newCount       = 0
    $overwriteCount = 0
    $missingCount   = 0
    $bundleCount    = 0

    foreach ($entry in $entries) {
        $sourceFile = Join-Path $sourceFolder $entry.EncodedName
        $destFile = Join-Path $destFolder   $entry.OriginalPath

        if (-not (Test-Path $sourceFile)) {
            Write-Host "  MISSING    $($entry.EncodedName)" -ForegroundColor Red
            $missingCount++
            continue
        }

        # Check main file
        if (Test-Path $destFile) {
            Write-Host "  OVERWRITE  $($entry.OriginalPath)" -ForegroundColor Yellow
            $overwriteCount++
        }
        else {
            Write-Host "  NEW        $($entry.OriginalPath)" -ForegroundColor Green
            $newCount++
        }

        # Check for bundled .psd1
        $content = Get-Content -Path $sourceFile -Raw
        if ($content -match '(?m)^# __BUNDLE_PSD1_START__ (.+)$') {
            $psd1RelPath = $Matches[1].Trim()
            $psd1Dest    = Join-Path $destFolder $psd1RelPath
            $bundleCount++

            if (Test-Path $psd1Dest) {
                Write-Host "  OVERWRITE  $psd1RelPath (bundled)" -ForegroundColor Yellow
                $overwriteCount++
            }
            else {
                Write-Host "  NEW        $psd1RelPath (bundled)" -ForegroundColor Green
                $newCount++
            }
        }
    }

    Write-Host ""
    Write-Host "Summary:" -ForegroundColor Cyan
    Write-Host "  New files:       $newCount"
    Write-Host "  Overwrites:      $overwriteCount"
    Write-Host "  Missing sources: $missingCount"
    if ($bundleCount -gt 0) {
        Write-Host "  Bundled .psd1:   $bundleCount (extracted from parent scripts)"
    }
    Write-Host ""
    return
}


#endregion --- WhatIf (Dry Run) ---



#region --- Full Import ---


# Create backup folder for undo support
$formattedDate = Get-Date -Format 'yyyyMMdd_HHmm'
$backupFolder  = Join-Path $destFolder "_PreImport_Backup_$formattedDate"
New-Item -ItemType Directory -Force -Path $backupFolder | Out-Null

$backupManifest = [System.Collections.Generic.List[string]]::new()
$backupManifest.Add("BackupName`tRelativePath`tAction")

$importedCount = 0
$skippedCount  = 0
$backedUpCount = 0
$dirsCreated   = [System.Collections.Generic.HashSet[string]]::new()


$unbundledCount = 0

foreach ($entry in $entries) {
    $sourceFile = Join-Path $sourceFolder $entry.EncodedName
    $destFile = Join-Path $destFolder   $entry.OriginalPath

    # Skip if source file is missing from the export folder
    if (-not (Test-Path $sourceFile)) {
        Write-Host "  MISSING: $($entry.EncodedName)" -ForegroundColor Red
        $skippedCount++
        continue
    }

    # Read file content to check for bundled .psd1
    $content    = Get-Content -Path $sourceFile -Raw
    $isBundled  = $false
    $psd1Path   = $null
    $mainBody   = $null
    $psd1Body   = $null

    if ($content -match '(?m)^# __BUNDLE_PSD1_START__ (.+)$') {
        $isBundled     = $true
        $psd1RelPath   = $Matches[1].Trim()
        $psd1Path      = Join-Path $destFolder $psd1RelPath

        # The export prepends CRLF before the marker, so search for that exact
        # sequence to split precisely -- this preserves the original file content
        # including its trailing newline (if any) without adding or removing bytes.
        $markerPrefix = "`r`n# __BUNDLE_PSD1_START__"
        $startIdx     = $content.IndexOf($markerPrefix)

        # Main file content: everything before the CRLF that the export added
        $mainBody = $content.Substring(0, $startIdx)

        # PSD1 content: after the START line, before the CRLF + END marker
        $afterMarker  = $content.Substring($startIdx + $markerPrefix.Length)
        $lineEnd      = $afterMarker.IndexOf("`n") + 1
        $endMarker    = "`r`n# __BUNDLE_PSD1_END__"
        $endIdx       = $afterMarker.IndexOf($endMarker, $lineEnd)
        $psd1Body     = $afterMarker.Substring($lineEnd, $endIdx - $lineEnd)
    }

    # --- Process main file ---

    # Ensure target directory exists
    $targetDir = Split-Path $destFile -Parent
    if (-not (Test-Path $targetDir)) {
        New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
        [void]$dirsCreated.Add($targetDir)
    }

    # Backup existing file before overwriting
    if (Test-Path $destFile) {
        $backupName = $entry.EncodedName -replace '\.txt$', ''
        Copy-Item $destFile -Destination (Join-Path $backupFolder $backupName)
        $backupManifest.Add("$backupName`t$($entry.OriginalPath)`tOverwritten")
        $backedUpCount++
    }
    else {
        $backupManifest.Add("`t$($entry.OriginalPath)`tCreated")
    }

    if ($isBundled) {
        # Write only the main file content (without the bundled .psd1)
        [System.IO.File]::WriteAllText($destFile, $mainBody, [System.Text.UTF8Encoding]::new($true))
    }
    else {
        Copy-Item $sourceFile -Destination $destFile -Force
    }

    Unblock-File -Path $destFile
    $importedCount++

    # --- Process bundled .psd1 (if present) ---

    if ($isBundled) {
        $psd1Dir = Split-Path $psd1Path -Parent
        if (-not (Test-Path $psd1Dir)) {
            New-Item -ItemType Directory -Force -Path $psd1Dir | Out-Null
            [void]$dirsCreated.Add($psd1Dir)
        }

        if (Test-Path $psd1Path) {
            $psd1BackupName = ($psd1RelPath -replace '\\', '--')
            Copy-Item $psd1Path -Destination (Join-Path $backupFolder $psd1BackupName)
            $backupManifest.Add("$psd1BackupName`t$psd1RelPath`tOverwritten")
            $backedUpCount++
        }
        else {
            $backupManifest.Add("`t$psd1RelPath`tCreated")
        }

        [System.IO.File]::WriteAllText($psd1Path, $psd1Body, [System.Text.UTF8Encoding]::new($true))
        Unblock-File -Path $psd1Path

        $unbundledCount++
        $importedCount++
    }
}


# Write backup manifest
$backupManifestPath = Join-Path $backupFolder '_BACKUP_MANIFEST.txt'
$backupManifest | Out-File -FilePath $backupManifestPath -Encoding UTF8

# If nothing was backed up, clean up the empty backup folder
if ($backedUpCount -eq 0 -and $importedCount -gt 0) {
    # Keep the backup folder anyway -- it tracks "Created" entries for undo
}


#endregion --- Full Import ---



#region --- Summary ---


Write-Host ""
Write-Host "Import complete!" -ForegroundColor Green
Write-Host "  Files imported:    $importedCount"
if ($unbundledCount -gt 0) {
    Write-Host "  Unbundled .psd1:   $unbundledCount (extracted from parent scripts)"
}
Write-Host "  Files skipped:     $skippedCount"
Write-Host "  Dirs created:      $($dirsCreated.Count)"
Write-Host "  Files backed up:   $backedUpCount"
Write-Host "  Backup location:   $backupFolder"
Write-Host ""
Write-Host "To preview before importing, run:" -ForegroundColor Gray
Write-Host "  .\Import-Package.ps1 -WhatIf" -ForegroundColor Gray
Write-Host ""
Write-Host "To undo this import, run:" -ForegroundColor Gray
Write-Host "  .\Import-Package.ps1 -Undo" -ForegroundColor Gray
Write-Host ""


#endregion --- Summary ---
