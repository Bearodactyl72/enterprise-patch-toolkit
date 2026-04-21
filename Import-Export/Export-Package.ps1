# DOTS formatting comment

<#
    .SYNOPSIS
        Exports the project to a single folder with .txt extensions and a manifest for reconstruction.
    .DESCRIPTION
        Creates a packaged copy of the project where every file is renamed with a path-encoded
        name and .txt extension (safe for email attachment filters). A _MANIFEST.txt file
        preserves the original directory structure so Import-Package.ps1 can reconstruct the tree.

        Naming convention:
            Modules\Add-Delimiter.psm1  -->  Modules--Add-Delimiter.psm1.txt
            Scripts\Invoke-Patch.ps1    -->  Scripts--Invoke-Patch.ps1.txt
            Export-Package.ps1          -->  Export-Package.ps1.txt

        Delta export (-ReferenceExport):
            Point to a previous export folder to export only files that have been
            added or modified since that package was created. The reference manifest's
            SHA256 hashes are compared against current file hashes to detect changes.
            The resulting package can be imported normally with Import-Package.ps1.

        Written by Skyler Werner
        Date: 2026/03/04
    .PARAMETER ReferenceExport
        Path to a previous export folder. When provided, only new and modified
        files (compared by SHA256 hash) are included in the export.
    .EXAMPLE
        .\Export-Package.ps1
        Full export of all project files.
    .EXAMPLE
        .\Export-Package.ps1 -ReferenceExport ".\Export_20260319"
        Delta export containing only files changed since the 2026-03-19 package.
#>

param(
    [Alias('Diff', 'Since')]
    [string]$ReferenceExport
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



#region --- Folder Selection ---


# Default to project root (this script lives in Import-Export\, so one level up)
$defaultPath = if ($psISE) {
    Split-Path -Path $psISE.CurrentFile.FullPath
}
elseif ($PSScriptRoot) {
    Split-Path $PSScriptRoot -Parent
}
else {
    (Get-Location).Path
}


# Pick source folder
$sourceFolder = Show-FolderPicker `
    -Description "Select the project folder to export." `
    -DefaultPath $defaultPath

if (-not $sourceFolder) {
    [void][Microsoft.VisualBasic.Interaction]::MsgBox(
        "No source folder selected. Export cancelled.",
        "OKOnly,SystemModal,Exclamation,DefaultButton1",
        " Export Cancelled"
    )
    return
}


# Pick destination folder
$destFolder = Show-FolderPicker `
    -Description "Select or create a destination folder for the export." `
    -DefaultPath $PSScriptRoot `
    -ShowNewFolderButton

if (-not $destFolder) {
    [void][Microsoft.VisualBasic.Interaction]::MsgBox(
        "No destination folder selected. Export cancelled.",
        "OKOnly,SystemModal,Exclamation,DefaultButton1",
        " Export Cancelled"
    )
    return
}


#endregion --- Folder Selection ---



#region --- Validate Reference Export ---


$isDelta    = $false
$refHashes  = @{}

if ($ReferenceExport) {
    # Resolve to full path
    $ReferenceExport = (Resolve-Path $ReferenceExport -ErrorAction SilentlyContinue).Path

    if (-not $ReferenceExport -or -not (Test-Path $ReferenceExport)) {
        [void][Microsoft.VisualBasic.Interaction]::MsgBox(
            "Reference export folder not found:`n$ReferenceExport",
            "OKOnly,SystemModal,Exclamation,DefaultButton1",
            " Reference Not Found"
        )
        return
    }

    $refManifest = Join-Path $ReferenceExport '_MANIFEST.txt'
    if (-not (Test-Path $refManifest)) {
        [void][Microsoft.VisualBasic.Interaction]::MsgBox(
            "_MANIFEST.txt not found in reference folder.`nCannot perform delta comparison.",
            "OKOnly,SystemModal,Exclamation,DefaultButton1",
            " Manifest Not Found"
        )
        return
    }

    # Parse the reference manifest -- expects tab-delimited with optional Hash column
    $refManifestLines = @{}
    $refLines = Get-Content $refManifest | Select-Object -Skip 1
    foreach ($line in $refLines) {
        $parts = $line -split "`t"
        if ($parts.Count -ge 3 -and $parts[0].Length -gt 0) {
            # Manifest has hash column: EncodedName, OriginalPath, SHA256
            $refHashes[$parts[1]]        = $parts[2]
            $refManifestLines[$parts[1]] = $line
        }
    }

    if ($refHashes.Count -eq 0) {
        [void][Microsoft.VisualBasic.Interaction]::MsgBox(
            "The reference manifest does not contain SHA256 hashes.`n`nA full export with hash support must be created first`nbefore delta exports can be used.`n`nRun a full export (without -ReferenceExport) to generate`na hash-enabled manifest, then use that as your reference.",
            "OKOnly,SystemModal,Exclamation,DefaultButton1",
            " No Hashes in Reference"
        )
        return
    }

    $isDelta = $true
    Write-Host ""
    Write-Host "Delta mode: comparing against $ReferenceExport" -ForegroundColor Cyan
    Write-Host "  Reference contains $($refHashes.Count) file hashes." -ForegroundColor Gray
}


#endregion --- Validate Reference Export ---



#region --- File Enumeration & Exclusion ---


$formattedDate = Get-Date -Format 'yyyyMMdd'
$folderPrefix  = if ($isDelta) { "Delta_$formattedDate" } else { "Export_$formattedDate" }

$exportFolder = Join-Path $destFolder $folderPrefix
if (Test-Path $exportFolder) {
    $suffix = 2
    while (Test-Path (Join-Path $destFolder "${folderPrefix}_$suffix")) {
        $suffix++
    }
    $exportFolder = Join-Path $destFolder "${folderPrefix}_$suffix"
}

New-Item -ItemType Directory -Force -Path $exportFolder | Out-Null


# Enumerate all files
$allFiles = Get-ChildItem $sourceFolder -Recurse -File


# Exclusion patterns (matched against full path)
$excludePatterns = @(
    '[\\/]\.git[\\/]',
    '[\\/]\.git$',
    '[\\/]\.claude[\\/]',
    '[\\/]CLAUDE\.md$',
    '[\\/]AGENTS\.md$',
    '[\\/]\.gitignore$',
    '[\\/]\.gitattributes$',
    '[\\/]\.vscode[\\/]',
    '[\\/]PSScriptAnalyzerSettings\.psd1$',
    '[\\/]Export_(Flat_)?[\d]',
    '[\\/](Export_)?Delta_',
    '[\\/]SwitchBackups[\\/]',
    'Working',
    'Final Upload',
    '[\\/]desktop\.ini$',
    '[\\/]_MANIFEST\.txt$'
)

$files = $allFiles | Where-Object {
    $path     = $_.FullName
    $excluded = $false

    foreach ($pattern in $excludePatterns) {
        if ($path -match $pattern) {
            $excluded = $true
            break
        }
    }

    -not $excluded
}


#endregion --- File Enumeration & Exclusion ---



#region --- Identify PSD1 Bundle Pairs ---


# Build a lookup of all file relative paths for quick matching
$filesByRelPath = @{}
foreach ($file in $files) {
    $rel = $file.FullName.Substring($sourceFolder.Length + 1)
    $filesByRelPath[$rel] = $file
}

# Find .psd1 files that can be bundled into a matching .ps1 or .psm1
# Key = main file relative path, Value = .psd1 file info object
$bundleMap    = @{}
$bundledPsd1s = @{}

foreach ($rel in @($filesByRelPath.Keys)) {
    if ($rel -notmatch '\.psd1$') { continue }

    $dir      = Split-Path $rel -Parent
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($rel)

    # Look for a matching .psm1 first, then .ps1
    $psm1Path = if ($dir) { "$dir\$baseName.psm1" } else { "$baseName.psm1" }
    $ps1Path  = if ($dir) { "$dir\$baseName.ps1"  } else { "$baseName.ps1"  }

    $mainPath = $null
    if ($filesByRelPath.ContainsKey($psm1Path)) {
        $mainPath = $psm1Path
    }
    elseif ($filesByRelPath.ContainsKey($ps1Path)) {
        $mainPath = $ps1Path
    }

    if ($mainPath) {
        $bundleMap[$mainPath]  = [PSCustomObject]@{
            Psd1RelativePath = $rel
            Psd1File         = $filesByRelPath[$rel]
        }
        $bundledPsd1s[$rel] = $true
    }
}

# Remove bundled .psd1 files from the file list -- they will not get their own entries
$files = $files | Where-Object {
    $rel = $_.FullName.Substring($sourceFolder.Length + 1)
    -not $bundledPsd1s.ContainsKey($rel)
}

if ($bundleMap.Count -gt 0) {
    Write-Host ""
    Write-Host "Bundling $($bundleMap.Count) .psd1 file(s) with their parent scripts." -ForegroundColor Gray
}


#endregion --- Identify PSD1 Bundle Pairs ---



#region --- Classify Files (Delta) ---


# Hash all current files and classify them for delta mode
$fileEntries   = [System.Collections.Generic.List[PSCustomObject]]::new()
$newFiles      = [System.Collections.Generic.List[string]]::new()
$modifiedFiles = [System.Collections.Generic.List[string]]::new()
$unchangedFiles = [System.Collections.Generic.List[string]]::new()
$deletedFiles  = [System.Collections.Generic.List[string]]::new()
$seenPaths     = @{}

foreach ($file in $files) {
    $relativePath = $file.FullName.Substring($sourceFolder.Length + 1)

    # For bundled files, compute hash on the combined content so delta
    # detection covers changes to either the main file or its .psd1
    if ($bundleMap.ContainsKey($relativePath)) {
        $mainContent = Get-Content -Path $file.FullName -Raw
        $psd1Content = Get-Content -Path $bundleMap[$relativePath].Psd1File.FullName -Raw
        $combined    = $mainContent + "`r`n# __BUNDLE_PSD1_START__ $($bundleMap[$relativePath].Psd1RelativePath)`r`n" + $psd1Content + "`r`n# __BUNDLE_PSD1_END__"

        $stream  = [System.IO.MemoryStream]::new([System.Text.Encoding]::UTF8.GetBytes($combined))
        $currentHash = (Get-FileHash -InputStream $stream -Algorithm SHA256).Hash
        $stream.Dispose()
    }
    else {
        $currentHash = (Get-FileHash -Path $file.FullName -Algorithm SHA256).Hash
    }

    $encodedName = ($relativePath -replace '\\', '--') + '.txt'

    $seenPaths[$relativePath] = $true

    # In delta mode, also mark the bundled .psd1 path as seen so it does not
    # appear as "deleted" when comparing against a pre-bundling reference manifest
    if ($bundleMap.ContainsKey($relativePath)) {
        $seenPaths[$bundleMap[$relativePath].Psd1RelativePath] = $true
    }

    if ($isDelta) {
        if ($refHashes.ContainsKey($relativePath) -and $refHashes[$relativePath] -eq $currentHash) {
            $unchangedFiles.Add($relativePath)
            continue
        }

        if ($refHashes.ContainsKey($relativePath)) {
            $modifiedFiles.Add($relativePath)
        } else {
            $newFiles.Add($relativePath)
        }
    }

    $fileEntries.Add([PSCustomObject]@{
        File         = $file
        RelativePath = $relativePath
        EncodedName  = $encodedName
        Hash         = $currentHash
    })
}

# Find deleted files: in reference but not in current source
if ($isDelta) {
    foreach ($refPath in $refHashes.Keys) {
        if (-not $seenPaths.ContainsKey($refPath)) {
            $deletedFiles.Add($refPath)
        }
    }
}


#endregion --- Classify Files (Delta) ---



#region --- Delta Preview & Confirmation ---


if ($isDelta) {
    Write-Host ""

    # --- New Files ---
    if ($newFiles.Count -gt 0) {
        Write-Host "--- New Files ($($newFiles.Count)) ---" -ForegroundColor Green
        foreach ($f in $newFiles) {
            Write-Host "  + $f" -ForegroundColor Green
        }
        Write-Host ""
    }

    # --- Modified Files ---
    if ($modifiedFiles.Count -gt 0) {
        Write-Host "--- Modified Files ($($modifiedFiles.Count)) ---" -ForegroundColor Yellow
        foreach ($f in $modifiedFiles) {
            Write-Host "  ~ $f" -ForegroundColor Yellow
        }
        Write-Host ""
    }

    # --- Deleted Files ---
    if ($deletedFiles.Count -gt 0) {
        Write-Host "--- Deleted Files ($($deletedFiles.Count)) ---" -ForegroundColor Red
        foreach ($f in $deletedFiles) {
            Write-Host "  - $f" -ForegroundColor Red
        }
        Write-Host ""
    }

    # --- Unchanged ---
    Write-Host "--- $($unchangedFiles.Count) unchanged files not included ---" -ForegroundColor DarkGray
    Write-Host ""

    # --- Totals ---
    $deltaTotal = $newFiles.Count + $modifiedFiles.Count
    Write-Host "  $deltaTotal file(s) will be exported." -ForegroundColor Cyan

    if ($deletedFiles.Count -gt 0) {
        Write-Host "  $($deletedFiles.Count) deleted file(s) noted (must be removed manually on import)." -ForegroundColor DarkGray
    }

    Write-Host ""

    # --- Confirmation ---
    if ($deltaTotal -eq 0) {
        Write-Host "No changes detected. Nothing to export." -ForegroundColor Gray
        # Clean up the empty export folder
        Remove-Item $exportFolder -Force
        return
    }

    $confirm = Read-Host "Proceed with delta export? [Y/n]"
    if ($confirm -and $confirm -notmatch '^[Yy]') {
        Write-Host "Export cancelled." -ForegroundColor Yellow
        Remove-Item $exportFolder -Force
        return
    }

    Write-Host ""
}


#endregion --- Delta Preview & Confirmation ---



#region --- Build Manifest & Copy Files ---


$manifest = [System.Collections.Generic.List[string]]::new()
$manifest.Add("EncodedName`tOriginalPath`tSHA256")

$fileCount = 0
$totalSize = 0

$bundledCount = 0

foreach ($entry in $fileEntries) {

    # Record in manifest (with hash)
    $manifest.Add("$($entry.EncodedName)`t$($entry.RelativePath)`t$($entry.Hash)")

    $destPath = Join-Path $exportFolder $entry.EncodedName

    if ($bundleMap.ContainsKey($entry.RelativePath)) {
        # Bundled file: combine main content + delimiter + .psd1 content
        $mainContent = Get-Content -Path $entry.File.FullName -Raw
        $psd1Info    = $bundleMap[$entry.RelativePath]
        $psd1Content = Get-Content -Path $psd1Info.Psd1File.FullName -Raw

        $combined = $mainContent + "`r`n# __BUNDLE_PSD1_START__ $($psd1Info.Psd1RelativePath)`r`n" + $psd1Content + "`r`n# __BUNDLE_PSD1_END__"

        [System.IO.File]::WriteAllText($destPath, $combined, [System.Text.UTF8Encoding]::new($false))

        $bundledCount++
        $totalSize += $entry.File.Length + $psd1Info.Psd1File.Length
    }
    else {
        # Normal file: straight copy
        Copy-Item $entry.File.FullName -Destination $destPath
        $totalSize += $entry.File.Length
    }

    $fileCount++
}


# In delta mode, carry forward unchanged entries so this manifest is a complete
# state snapshot -- enabling it to be used as a reference for future deltas.
if ($isDelta) {
    foreach ($path in $unchangedFiles) {
        if ($refManifestLines.ContainsKey($path)) {
            $manifest.Add($refManifestLines[$path])
        }
    }
}


# Write manifest as first file in the export
$manifestPath = Join-Path $exportFolder '_MANIFEST.txt'
$manifest | Out-File -FilePath $manifestPath -Encoding UTF8


#endregion --- Build Manifest & Copy Files ---



#region --- Summary ---


$sizeMB = [math]::Round($totalSize / 1MB, 2)

Write-Host ""

if ($isDelta) {
    Write-Host "Delta export complete!" -ForegroundColor Green
    Write-Host "  New files:       $($newFiles.Count)"
    Write-Host "  Modified files:  $($modifiedFiles.Count)"
    Write-Host "  Deleted files:   $($deletedFiles.Count)"
    Write-Host "  Unchanged:       $($unchangedFiles.Count) (not included)"
    Write-Host "  -------------------------"
    Write-Host "  Total exported:  $fileCount"
    Write-Host "  Total size:      $sizeMB MB"
    Write-Host "  Destination:     $exportFolder"
    Write-Host "  Reference:       $ReferenceExport"
    Write-Host ""
    Write-Host "Import this delta package the same way as a full export:" -ForegroundColor Gray
    Write-Host "  .\Import-Package.ps1" -ForegroundColor Gray
    Write-Host "  (select the delta folder, then point to your existing project)" -ForegroundColor Gray
} else {
    Write-Host "Export complete!" -ForegroundColor Green
    Write-Host "  Files:       $fileCount"
    if ($bundledCount -gt 0) {
        Write-Host "  Bundled:     $bundledCount .psd1 file(s) embedded in parent scripts"
    }
    Write-Host "  Total size:  $sizeMB MB"
    Write-Host "  Destination: $exportFolder"
    Write-Host "  Manifest:    $manifestPath"
    Write-Host ""
    Write-Host "This export can be used as a reference for future delta exports:" -ForegroundColor Gray
    Write-Host "  .\Export-Package.ps1 -ReferenceExport `"$exportFolder`"" -ForegroundColor Gray
}

Write-Host ""


#endregion --- Summary ---
