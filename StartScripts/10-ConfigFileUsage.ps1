#Requires -Version 5.1
<#
.SYNOPSIS
    Demo: Load Settings from JSON Configuration File

.DESCRIPTION
    Runs DelprofPS with -ConfigFile to load settings from a JSON file.
    This is ideal for standardizing settings across an organization or
    for scheduled tasks where you want consistent behavior.

    The config file can set: DaysInactive, Exclude, Include, LogPath,
    OutputPath, HtmlReport, BackupPath, email settings, and more.

    WHAT THIS DEMONSTRATES:
    - The -ConfigFile parameter
    - JSON configuration file format
    - How config values are applied (command-line overrides config)

.NOTES
    SAFE TO RUN - Uses default config, no deletion (no -Delete flag).
#>

$mainScript = Join-Path $PSScriptRoot '..\DelprofPS.ps1'
$configFile = Join-Path $PSScriptRoot '..\DelprofPS.config.json'

Write-Host "`n  DEMO: JSON Configuration File" -ForegroundColor Cyan

if (Test-Path $configFile) {
    Write-Host "  Config file: $configFile" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Config contents:" -ForegroundColor Yellow
    $config = Get-Content $configFile | ConvertFrom-Json
    Write-Host "    DaysInactive:  $($config.DaysInactive)" -ForegroundColor DarkGray
    Write-Host "    Exclude:       $($config.Exclude -join ', ')" -ForegroundColor DarkGray
    Write-Host "    LogPath:       $($config.LogPath)" -ForegroundColor DarkGray
    Write-Host "    OutputPath:    $($config.OutputPath)" -ForegroundColor DarkGray
    Write-Host "    BackupPath:    $($config.BackupPath)" -ForegroundColor DarkGray
    Write-Host ""

    # Run with config file (dry-run, no -Delete)
    & $mainScript -ConfigFile $configFile -ShowSpace
}
else {
    Write-Host "  Config file not found: $configFile" -ForegroundColor Red
    Write-Host "  Copy DelprofPS.config.json to the project root to use this demo." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "  TIP: Command-line parameters override config file values." -ForegroundColor Green
Write-Host "       .\DelprofPS.ps1 -ConfigFile config.json -DaysInactive 120`n" -ForegroundColor Gray
