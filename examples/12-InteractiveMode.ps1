#Requires -Version 5.1
<#
.SYNOPSIS
    Demo: Interactive Profile Selection

.DESCRIPTION
    Runs DelprofPS with -Interactive to present a visual menu where you
    can manually select which profiles to process. The menu shows profile
    details and lets you toggle selections before confirming.

    WHAT THIS DEMONSTRATES:
    - The -Interactive parameter
    - Select-ProfilesInteractive visual menu
    - Manual profile selection before any action
    - Safe, user-controlled deletion workflow

.NOTES
    INTERACTIVE - You will be prompted to select profiles.
    No profiles are deleted unless you explicitly confirm.
#>

$mainScript = Join-Path $PSScriptRoot '..\DelprofPS.ps1'

Write-Host "`n  DEMO: Interactive Profile Selection" -ForegroundColor Cyan
Write-Host "  You will see a menu to select profiles manually." -ForegroundColor Gray
Write-Host "  No profiles are deleted unless you explicitly choose to.`n" -ForegroundColor Green

& $mainScript -DaysInactive 0 -Interactive -ShowSpace
