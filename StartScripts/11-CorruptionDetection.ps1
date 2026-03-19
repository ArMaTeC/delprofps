#Requires -Version 5.1
<#
.SYNOPSIS
    Demo: Corrupted Profile Detection

.DESCRIPTION
    Runs DelprofPS with -IncludeCorrupted to detect profiles that have
    issues such as:
    - Missing profile path (registry entry exists but folder is gone)
    - Missing NTUSER.DAT (profile folder exists but hive file is missing)

    DelprofPS normally skips corrupted profiles. This flag includes them
    in the output so you can identify and address them.

    WHAT THIS DEMONSTRATES:
    - The -IncludeCorrupted parameter
    - Profile state detection (Corrupted, Local, Roaming, etc.)
    - The -FixCorruption parameter (mentioned, not executed)
    - Get-ProfileType function behavior

.NOTES
    SAFE TO RUN - Only detects corruption, no changes made.
#>

$mainScript = Join-Path $PSScriptRoot '..\DelprofPS.ps1'

Write-Host "`n  DEMO: Corrupted Profile Detection" -ForegroundColor Cyan
Write-Host "  Scans for profiles with missing paths or NTUSER.DAT files.`n" -ForegroundColor Gray

& $mainScript -DaysInactive 0 -IncludeCorrupted -ShowSpace

Write-Host ""
Write-Host "  ────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  To interactively repair corrupted profiles, run:" -ForegroundColor Yellow
Write-Host "    .\DelprofPS.ps1 -IncludeCorrupted -FixCorruption -Interactive" -ForegroundColor White
Write-Host ""
Write-Host "  Repair options include:" -ForegroundColor Gray
Write-Host "    [R] Recreate NTUSER.DAT from default template" -ForegroundColor DarkGray
Write-Host "    [D] Delete corrupted profile entirely" -ForegroundColor DarkGray
Write-Host "    [F] Remove orphaned registry key" -ForegroundColor DarkGray
Write-Host "    [S] Skip and continue" -ForegroundColor DarkGray
Write-Host ""
