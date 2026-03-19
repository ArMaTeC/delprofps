#Requires -Version 5.1
<#
.SYNOPSIS
    Demo: Scheduled Task / Quiet Mode Example

.DESCRIPTION
    Shows how to configure DelprofPS for unattended operation in a
    scheduled task. Key features for automation:

    - -Quiet: Suppresses all console output
    - -Force: Skips confirmation prompts
    - -LogPath: Writes audit log to file
    - -OutputPath: Exports CSV results
    - -HtmlReport: Generates management report
    - Email notifications via SMTP

    This demo shows the command that would be used but does NOT
    actually delete anything (no -Delete flag).

    WHAT THIS DEMONSTRATES:
    - The -Quiet parameter
    - The -Force parameter
    - Combining logging, export, and notifications
    - Typical scheduled task configuration

.NOTES
    SAFE TO RUN - Shows configuration only, no deletion.
#>

$mainScript = Join-Path $PSScriptRoot '..\DelprofPS.ps1'
$logDir = Join-Path $PSScriptRoot '..\Logs'

if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

Write-Host "`n  DEMO: Scheduled Task Configuration" -ForegroundColor Cyan
Write-Host "  Shows how DelprofPS is configured for unattended use.`n" -ForegroundColor Gray

Write-Host "  Typical scheduled task command:" -ForegroundColor Yellow
Write-Host @"

    powershell.exe -ExecutionPolicy Bypass -File "DelprofPS.ps1" ``
        -DaysInactive 60 ``
        -Delete ``
        -Force ``
        -Quiet ``
        -Exclude "*admin*", "*service*" ``
        -UnloadHives ``
        -LogPath "C:\Logs\DelprofPS.log" ``
        -OutputPath "C:\Logs\DelprofPS_Results.csv" ``
        -HtmlReport "C:\Logs\DelprofPS_Report.html" ``
        -BackupPath "C:\Backups\Profiles" ``
        -SmtpServer "mail.company.com" ``
        -EmailTo "admin@company.com"

"@ -ForegroundColor White

Write-Host "  Running a preview with logging (no -Delete, not quiet)...`n" -ForegroundColor Gray

& $mainScript -DaysInactive 60 `
    -Exclude "*admin*", "*service*" `
    -ShowSpace `
    -LogPath (Join-Path $logDir 'DelprofPS_Demo.log') `
    -OutputPath (Join-Path $logDir 'DelprofPS_DemoResults.csv')

Write-Host ""
if (Test-Path (Join-Path $logDir 'DelprofPS_Demo.log')) {
    Write-Host "  Log saved to: $(Join-Path $logDir 'DelprofPS_Demo.log')" -ForegroundColor Green
}
if (Test-Path (Join-Path $logDir 'DelprofPS_DemoResults.csv')) {
    Write-Host "  CSV saved to: $(Join-Path $logDir 'DelprofPS_DemoResults.csv')" -ForegroundColor Green
}
Write-Host ""
