#Requires -Version 5.1
<#
.SYNOPSIS
    Demo: Include / Exclude / Size Filtering

.DESCRIPTION
    Runs DelprofPS with various filtering options to show how you can
    target specific profiles:

    1. Include filter  - Only process profiles matching a pattern
    2. Exclude filter  - Skip profiles matching a pattern
    3. Size filter     - Only profiles above a minimum size
    4. Combined        - Multiple filters together

    WHAT THIS DEMONSTRATES:
    - The -Include parameter (wildcard patterns)
    - The -Exclude parameter (wildcard patterns)
    - The -MinProfileSizeMB parameter
    - The -MaxProfileSizeMB parameter
    - How filters combine (AND logic)

.NOTES
    SAFE TO RUN - Read-only listing, no changes made.
#>

$mainScript = Join-Path $PSScriptRoot '..\DelprofPS.ps1'

Write-Host "`n  DEMO: Profile Filtering" -ForegroundColor Cyan

# Example 1: Exclude admin accounts
Write-Host "`n  ── Filter 1: Exclude *admin* accounts ────────────────────" -ForegroundColor Yellow
& $mainScript -DaysInactive 0 -Exclude "*admin*" -ShowSpace

# Example 2: Only include specific pattern
Write-Host "`n  ── Filter 2: Include only 'user*' pattern ────────────────" -ForegroundColor Yellow
& $mainScript -DaysInactive 0 -Include "user*" -ShowSpace

# Example 3: Size filter - profiles over 500MB
Write-Host "`n  ── Filter 3: Only profiles larger than 500 MB ────────────" -ForegroundColor Yellow
& $mainScript -DaysInactive 0 -MinProfileSizeMB 500 -ShowSpace

Write-Host "`n  TIP: Combine filters for precision:" -ForegroundColor Green
Write-Host "       .\DelprofPS.ps1 -Exclude '*admin*','*service*' -MinProfileSizeMB 100 -DaysInactive 90`n" -ForegroundColor Gray
