#Requires -Version 5.1
<#
.SYNOPSIS
    Demo: Disk Space & Detailed Folder Breakdown

.DESCRIPTION
    Runs DelprofPS with -ShowSpace and -Detailed to display how much disk
    space each profile consumes, with a per-folder breakdown showing sizes
    for Documents, Downloads, Desktop, AppData, Pictures, Videos, and Music.

    WHAT THIS DEMONSTRATES:
    - The -ShowSpace parameter (total profile size)
    - The -Detailed parameter (folder-level breakdown)
    - Get-ProfileFolderSize and Get-ProfileSizeBreakdown functions
    - Format-Byte human-readable size formatting

.NOTES
    SAFE TO RUN - Read-only analysis, no changes made.
#>

$mainScript = Join-Path $PSScriptRoot '..\DelprofPS.ps1'

Write-Host "`n  DEMO: Disk Space & Detailed Folder Breakdown" -ForegroundColor Cyan
Write-Host "  Shows profile sizes with per-folder breakdown.`n" -ForegroundColor Gray

& $mainScript -DaysInactive 0 -ShowSpace -Detailed
