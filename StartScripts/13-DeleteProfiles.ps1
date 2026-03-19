#Requires -Version 5.1
<#
.SYNOPSIS
    Demo: Delete Profiles (with Safety Prompts)

.DESCRIPTION
    Runs DelprofPS with -Delete to actually remove profiles. Multiple
    safety layers are in place:

    1. Confirmation prompt before any deletion
    2. Active session detection (logged-in users are skipped)
    3. System profile protection (Administrator, SYSTEM, etc. skipped)
    4. Mass deletion warning if > 50 profiles would be deleted
    5. -WhatIf support via PowerShell's ShouldProcess

    WHAT THIS DEMONSTRATES:
    - The -Delete parameter (actual deletion)
    - The -Exclude parameter (safety exclusions)
    - Confirmation prompts and safety checks
    - Active session protection

.NOTES
    *** CAUTION: This script WILL delete profiles if you confirm. ***
    Profiles older than 120 days (excluding admins/services) are targeted.
    You will be prompted before any deletion occurs.
#>

$mainScript = Join-Path $PSScriptRoot '..\DelprofPS.ps1'

Write-Host "`n  DEMO: Delete Profiles" -ForegroundColor Red
Write-Host "  *** This will DELETE profiles older than 120 days ***" -ForegroundColor Red
Write-Host "  Excluding: *admin*, *service*, Administrator*" -ForegroundColor Yellow
Write-Host "  You will be prompted for confirmation before deletion.`n" -ForegroundColor Gray

$confirm = Read-Host "  Proceed with deletion demo? (YES to continue, anything else to cancel)"
if ($confirm -ne 'YES') {
    Write-Host "  Cancelled.`n" -ForegroundColor Yellow
    exit 0
}

& $mainScript -DaysInactive 120 -Delete -Exclude "*admin*", "*service*", "Administrator*" -ShowSpace -UnloadHives
