#Requires -Version 5.1
<#
.SYNOPSIS
    Demo: Remote Computer Targeting

.DESCRIPTION
    Shows how DelprofPS targets remote computers for profile management.
    Supports multiple input methods:

    - Direct names: -ComputerName SERVER1,SERVER2
    - From file: -ComputerName (Get-Content servers.txt)
    - From CSV: -ComputerName (Import-Csv servers.csv).Name
    - Pipeline: Get-ADComputer ... | .\DelprofPS.ps1

    Also demonstrates parallel processing for multiple targets.

    WHAT THIS DEMONSTRATES:
    - The -ComputerName parameter (remote targets)
    - The -UseParallel parameter (concurrent processing)
    - The -ThrottleLimit parameter (concurrency control)
    - Test-ComputerConnection for remote validation
    - Pipeline input capability

.NOTES
    SAFE TO RUN - Lists profiles only, no deletion.
    Remote computers must have WinRM/remote management enabled.
#>

$mainScript = Join-Path $PSScriptRoot '..\DelprofPS.ps1'

Write-Host "`n  DEMO: Remote Computer Targeting" -ForegroundColor Cyan
Write-Host ""

Write-Host "  Usage examples:" -ForegroundColor Yellow
Write-Host @"

    # Single remote computer
    .\DelprofPS.ps1 -ComputerName SERVER01 -ShowSpace

    # Multiple computers
    .\DelprofPS.ps1 -ComputerName SERVER01,SERVER02,SERVER03

    # From a text file (one name per line)
    .\DelprofPS.ps1 -ComputerName (Get-Content C:\servers.txt)

    # From a CSV file
    .\DelprofPS.ps1 -ComputerName (Import-Csv C:\servers.csv).ComputerName

    # Parallel processing for many computers
    .\DelprofPS.ps1 -ComputerName (Get-Content servers.txt) ``
        -UseParallel -ThrottleLimit 10

    # Pipeline from Active Directory
    Get-ADComputer -Filter * -SearchBase "OU=Workstations,DC=corp,DC=com" |
        Select-Object -ExpandProperty Name |
        .\DelprofPS.ps1 -DaysInactive 90 -ShowSpace

"@ -ForegroundColor White

Write-Host "  Running against local computer as demo...`n" -ForegroundColor Gray

& $mainScript -ComputerName $env:COMPUTERNAME -Test

Write-Host ""
Write-Host "  To target remote computers, replace `$env:COMPUTERNAME with" -ForegroundColor Green
Write-Host "  your server names in the command above.`n" -ForegroundColor Green
