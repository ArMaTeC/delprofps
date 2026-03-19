#Requires -Version 5.1
<#
.SYNOPSIS
    Demo: Basic Profile Listing (Default Dry-Run Mode)

.DESCRIPTION
    Runs DelprofPS in its default mode - lists all user profiles older than
    30 days without making any changes. This is the safest way to see what
    DelprofPS finds on your system.

    WHAT THIS DEMONSTRATES:
    - Default dry-run behavior (no -Delete = no changes)
    - Profile enumeration from the registry
    - Age calculation using NTUSER.DAT timestamps
    - Protected profile detection (system accounts skipped)
    - Color-coded age output

.NOTES
    SAFE TO RUN - No profiles are modified or deleted.
#>

$mainScript = Join-Path $PSScriptRoot '..\DelprofPS.ps1'

Write-Host "`n  DEMO: Basic Profile Listing" -ForegroundColor Cyan
Write-Host "  This lists all profiles older than 30 days (default)." -ForegroundColor Gray
Write-Host "  No changes will be made - this is a read-only operation.`n" -ForegroundColor Green

& $mainScript
