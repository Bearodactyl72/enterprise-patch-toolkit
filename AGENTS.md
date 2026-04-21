# Project Rules

## CRITICAL: ASCII-Only in PowerShell Files
All .ps1, .psm1, and .psd1 files MUST contain only ASCII characters (bytes 0x00-0x7F).
NEVER use Unicode characters such as em-dashes, en-dashes, curly/smart quotes,
box-drawing characters, or trademark symbols. This project targets PowerShell 5.1,
which reads files as Windows-1252 and corrupts multi-byte UTF-8 characters.

## Target Environment
- PSVersion:  5.1.20348.4294
- PSEdition:  Desktop
- CLRVersion: 4.0.30319.42000
- OS:         Windows Server (Build 20348) / Windows 11

Use only: regular hyphens (-), straight quotes (" and '), pipes (|), plus (+), equals (=).

For decorative comment borders, use repeated hyphens or equals signs:
  # --- Section Name ---
  # =====================

## PowerShell Conventions
All scripts in this repo target PowerShell 5.1. Be aware of PS 5.1 quirks: single-element arrays don't have .Count, $null differences, no ternary operator, no null-coalescing. Always use @() wrapping for collections that might be single items.

## Runspace + Invoke-Command Double-Serialization
Scripts that use Invoke-RunspacePool pass arguments via AddArgument(), then the outer
scriptblock calls Invoke-Command -ArgumentList to a remote session. This creates TWO
serialization boundaries. Arrays and strings received from $args[n] arrive in the remote
block as Deserialized.* types, which fail implicit string conversion in cmdlets like
Get-ChildItem, Test-Path, Remove-Item, and registry cmdlets.

Rule: Any time an array received from $args[n] is forwarded through Invoke-Command
-ArgumentList and used as a path or string parameter in the remote block, force-cast it
at the top of the remote block:

    $searchPaths = [string[]]($searchPaths | ForEach-Object { "$_" })

Scalars (int, bool, single string) survive deserialization fine. Only arrays and values
used directly as -Path arguments need this treatment.

## PowerShell File Editing Rules
NEVER modify .ps1, .psm1, or .psd1 files via Bash. This includes:
  - No echo/cat/sed/awk redirects
  - No PowerShell.exe -Command with file writes (Set-Content, Out-File, WriteAllText, etc.)
  - No [System.IO.File] calls via Bash
  - No regex-based find-and-replace via Bash (this has wiped files before)

- The ONLY safe way to modify PowerShell files is the Edit tool (or Write tool for new
files). This preserves the UTF-8 BOM and Windows line endings (CRLF) that PowerShell
5.1 and ISE expect.
- Never use bash `echo`/`cat`/heredoc to prepend or modify PowerShell files — this corrupts BOM and line endings. Use the Edit or Write tool instead.
- Preserve UTF-8 with BOM encoding on existing .ps1/.psm1/.psd1 files.

## Pre-Commit Validation
Before every commit that touches .ps1/.psm1/.psd1 files, run a sanity check:
  1. No file should have 0 lines (corruption/wipe detection)
  2. No non-ASCII bytes beyond the 3-byte UTF-8 BOM at offset 0-2
If either check fails, STOP and investigate before committing.

## Design Before Implementing
- For modernization, refactoring, or multi-file rewrites, discuss design goals and naming choices with the user BEFORE editing files. Propose parameter names and get confirmation.

## Workflow Rules
When asked for a commit message, write it in chat as text. Do NOT run git commit or any git commands unless explicitly asked to.

## Output Format
- Standardize script output on table format (Format-Table), not list view. Strip computer names/hostnames from error Comments fields for proper sorting.
