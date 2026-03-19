#Requires -Version 5.1
<#
.SYNOPSIS
    Demo: Generate HTML Report

.DESCRIPTION
    Runs DelprofPS with -HtmlReport to produce a professional styled HTML
    report containing all profile analysis results. The report includes:
    - Summary statistics (computers, profiles, space)
    - Color-coded profile table with age, size, type, status
    - Professional CSS styling suitable for management reporting

    WHAT THIS DEMONSTRATES:
    - The -HtmlReport parameter
    - Export-HtmlReport function output
    - Professional report suitable for stakeholders

.NOTES
    SAFE TO RUN - Generates a report file, no profiles are deleted.
#>

$mainScript = Join-Path $PSScriptRoot '..\DelprofPS.ps1'
$reportPath = Join-Path $PSScriptRoot '..\DelprofPS_DemoReport.html'

Write-Host "`n  DEMO: HTML Report Generation" -ForegroundColor Cyan
Write-Host "  Generating report to: $reportPath`n" -ForegroundColor Gray

& $mainScript -DaysInactive 0 -ShowSpace -HtmlReport $reportPath

if (Test-Path $reportPath) {
    Write-Host "`n  Report saved to: $reportPath" -ForegroundColor Green
    $open = Read-Host "  Open report in browser? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') {
        Start-Process $reportPath
    }
}
