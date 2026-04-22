# Contributing

Conventions and constraints for working in this codebase. Most of these exist because the target environment is Windows PowerShell 5.1 (Desktop edition), which has quirks that newer shells don't.

## Target Environment

- PSVersion:  5.1.20348.4294
- PSEdition:  Desktop
- CLRVersion: 4.0.30319.42000
- OS:         Windows Server (Build 20348) / Windows 11

PowerShell 7 is not available in the target environment. Every pattern in the repo is chosen to work on 5.1.

## ASCII-Only in PowerShell Files

All `.ps1`, `.psm1`, and `.psd1` files must contain only ASCII characters (bytes `0x00-0x7F`). No em-dashes, en-dashes, curly/smart quotes, box-drawing characters, or trademark symbols.

Why: PowerShell ISE on 5.1 reads BOM-less files as Windows-1252, which silently corrupts any multi-byte UTF-8 character. The project enforces ASCII-only as a belt-and-suspenders measure on top of the UTF-8-with-BOM file encoding.

Use only: regular hyphens (`-`), straight quotes (`"` and `'`), pipes (`|`), plus (`+`), equals (`=`).

For decorative comment borders, use repeated hyphens or equals signs:

    # --- Section Name ---
    # =====================

For software names with special characters (e.g. `Intel(R) PROSet`), use wildcards (`*` or `%`) in matches rather than embedding the special character.

## PowerShell 5.1 Quirks to Remember

- No null-coalescing (`??`), null-conditional (`?.`), or ternary (`?:`) operators.
- No `pwsh`-only cmdlets.
- Single-element collections don't have `.Count`. Always `@()`-wrap collections that might be single items.
- `$null` comparisons go on the left: `if ($null -eq $x)`, not `if ($x -eq $null)`.

## Runspace + Invoke-Command Double-Serialization

Scripts that use `Invoke-RunspacePool` pass arguments via `AddArgument()`, then the outer scriptblock calls `Invoke-Command -ArgumentList` against a remote session. This creates two serialization boundaries. Arrays and strings received from `$args[n]` arrive in the remote block as `Deserialized.*` types, which fail implicit string conversion in cmdlets like `Get-ChildItem`, `Test-Path`, `Remove-Item`, and registry cmdlets.

Rule: any time an array received from `$args[n]` is forwarded through `Invoke-Command -ArgumentList` and used as a path or string parameter in the remote block, force-cast it at the top of the remote block:

    $searchPaths = [string[]]($searchPaths | ForEach-Object { "$_" })

Scalars (`int`, `bool`, single `string`) survive deserialization fine. Only arrays and values used directly as `-Path` arguments need this treatment.

See `docs/ARCHITECTURE.md` for the full deep dive.

## File Encoding

- `.ps1`, `.psm1`, `.psd1` files are UTF-8 with BOM, CRLF line endings.
- PowerShell 5.1 ISE treats BOM-less files as Windows-1252; the BOM forces correct UTF-8 interpretation.

## Output Format

Scripts standardize on table output via `Format-Table` (not list view). Error `Comment` fields have computer names and hostnames stripped so the column sorts cleanly.

## Pre-Commit Sanity Checks

Before committing changes that touch `.ps1`/`.psm1`/`.psd1` files:

1. No file should have 0 lines (catches accidental wipes).
2. No non-ASCII bytes beyond the 3-byte UTF-8 BOM at offset 0-2.

`.claude/check-ascii.ps1` covers check #2.

## Invoke-RunspacePool Timeout Guard

Every `Invoke-RunspacePool` call needs post-processing that normalizes timed-out or failed jobs into full-width objects matching the standard result schema (`IP, ComputerName, Status, SoftwareName, Version, Compliant, NewVersion, ExitCode, Comment, AdminName, Date`). Without the guard, a timed-out job leaves a ragged object in the result array and the `Format-Table` output misaligns.
