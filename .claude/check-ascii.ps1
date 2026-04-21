# Hook script: Check for non-ASCII characters in PowerShell files
$j = [Console]::In.ReadToEnd() | ConvertFrom-Json
$f = $j.tool_input.file_path
if ($f -match '\.(ps1|psm1|psd1)$') {
    $bytes = [System.IO.File]::ReadAllBytes($f)
    $bad = $bytes | Where-Object { $_ -gt 127 }
    if ($bad) {
        Write-Error "Non-ASCII bytes found in: $f"
    }
}
