#Requires -Version 5.1
<#
.SYNOPSIS
    Demo: Test / Validate Mode

.DESCRIPTION
    Runs DelprofPS with -Test to validate prerequisites and connectivity
    without scanning or modifying any profiles. Useful for verifying that:
    - Admin rights are available
    - Target computer is reachable
    - Profile registry is accessible
    - Profile count can be retrieved

    WHAT THIS DEMONSTRATES:
    - The -Test parameter
    - Prerequisite validation
    - Connection testing for local/remote targets

.NOTES
    SAFE TO RUN - Only checks prerequisites, no profile changes.
#>

$mainScript = Join-Path $PSScriptRoot '..\DelprofPS.ps1'

Write-Host "`n  DEMO: Test / Validate Mode" -ForegroundColor Cyan
Write-Host "  Validates prerequisites and connectivity without changes.`n" -ForegroundColor Gray

& $mainScript -Test
