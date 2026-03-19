#Requires -Version 5.1
<#
.SYNOPSIS
    Demo: Age Calculation Methods

.DESCRIPTION
    Runs DelprofPS multiple times, each with a different -AgeCalculation
    method, so you can compare how each method determines profile age:

    - NTUSER_DAT : Uses the last-modified time of NTUSER.DAT (most reliable)
    - ProfilePath : Uses the last-modified time of the profile folder
    - Registry    : Uses LocalProfileLoadTime from the registry
    - LastLogon   : Uses the user's last logon time via ADSI

    WHAT THIS DEMONSTRATES:
    - The -AgeCalculation parameter with all supported methods
    - How different methods may return different ages for the same profile

.NOTES
    SAFE TO RUN - Read-only listing, no changes made.
#>

$mainScript = Join-Path $PSScriptRoot '..\DelprofPS.ps1'

$methods = @('NTUSER_DAT', 'ProfilePath', 'Registry', 'LastLogon')

foreach ($method in $methods) {
    Write-Host "`n  ── Age Method: $method ──────────────────────────────────────" -ForegroundColor Cyan
    & $mainScript -DaysInactive 0 -AgeCalculation $method -Quiet:$false
    Write-Host ""
}

Write-Host "  Compare the ages above - different methods may report different values." -ForegroundColor Yellow
