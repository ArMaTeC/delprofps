#Requires -Version 5.1
<#
.SYNOPSIS
    Demo: Preview / Simulation Mode

.DESCRIPTION
    Runs DelprofPS with -Preview to simulate what WOULD be deleted.
    Shows full profile details including age, size, and eligibility
    with a clear "PREVIEW MODE" banner. Nothing is deleted.

    WHAT THIS DEMONSTRATES:
    - The -Preview parameter
    - Full profile analysis without deletion
    - Preview banner and completion message
    - Profile eligibility assessment

.NOTES
    SAFE TO RUN - Simulation only, no profiles are deleted.
#>

$mainScript = Join-Path $PSScriptRoot '..\DelprofPS.ps1'

Write-Host "`n  DEMO: Preview / Simulation Mode" -ForegroundColor Cyan
Write-Host "  Shows what WOULD be deleted without actually deleting.`n" -ForegroundColor Gray

& $mainScript -Preview -DaysInactive 60
