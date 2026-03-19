#Requires -Version 5.1
<#
.SYNOPSIS
    Demo: CSV Export & Log File Output

.DESCRIPTION
    Runs DelprofPS with -OutputPath and -LogPath to demonstrate the
    logging and export capabilities:

    - CSV export: Machine-readable results for further analysis in Excel
    - Log file: Timestamped operational log for audit trails

    WHAT THIS DEMONSTRATES:
    - The -OutputPath parameter (CSV export)
    - The -LogPath parameter (detailed log file)
    - How results can be piped to other tools

.NOTES
    SAFE TO RUN - Creates output files, no profiles are deleted.
#>

$mainScript = Join-Path $PSScriptRoot '..\DelprofPS.ps1'
$csvPath = Join-Path $PSScriptRoot '..\DelprofPS_DemoResults.csv'
$logPath = Join-Path $PSScriptRoot '..\DelprofPS_DemoLog.log'

Write-Host "`n  DEMO: CSV Export & Logging" -ForegroundColor Cyan
Write-Host "  CSV Output:  $csvPath" -ForegroundColor Gray
Write-Host "  Log File:    $logPath`n" -ForegroundColor Gray

& $mainScript -DaysInactive 0 -ShowSpace -OutputPath $csvPath -LogPath $logPath

Write-Host ""
if (Test-Path $csvPath) {
    $records = Import-Csv $csvPath
    Write-Host "  CSV exported: $($records.Count) records" -ForegroundColor Green
    Write-Host "  Columns: $($records[0].PSObject.Properties.Name -join ', ')" -ForegroundColor DarkGray
}

if (Test-Path $logPath) {
    $lines = (Get-Content $logPath).Count
    Write-Host "  Log written: $lines lines" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Last 5 log entries:" -ForegroundColor Yellow
    Get-Content $logPath -Tail 5 | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
}
Write-Host ""
