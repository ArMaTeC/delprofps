#Requires -Version 5.1
<#
.SYNOPSIS
    DelprofPS Feature Launcher - Interactive menu to run any demo script.

.DESCRIPTION
    Presents a numbered menu of all DelprofPS feature demos.
    Select a number to launch the corresponding demo script.

.NOTES
    Run as Administrator for full functionality.
#>

$scriptDir = $PSScriptRoot
Clear-Host

Write-Host ""
Write-Host "  ============================================================================" -ForegroundColor Cyan
Write-Host "   DelprofPS - Feature Demo Launcher" -ForegroundColor Cyan
Write-Host "  ============================================================================" -ForegroundColor Cyan
Write-Host ""

$demos = @(
    @{ Num = 1;  File = "01-BasicListProfiles.ps1";       Desc = "Basic Profile Listing (dry-run, default mode)" }
    @{ Num = 2;  File = "02-TestMode.ps1";                Desc = "Test / Validate Mode (check prerequisites)" }
    @{ Num = 3;  File = "03-PreviewMode.ps1";             Desc = "Preview Mode (simulate what would be deleted)" }
    @{ Num = 4;  File = "04-ShowSpaceDetailed.ps1";       Desc = "Disk Space & Detailed Folder Breakdown" }
    @{ Num = 5;  File = "05-AgeCalculationMethods.ps1";   Desc = "Age Calculation Methods (NTUSER, Path, Registry, Logon)" }
    @{ Num = 6;  File = "06-FilteringProfiles.ps1";       Desc = "Include / Exclude / Size Filtering" }
    @{ Num = 7;  File = "07-HTMLReport.ps1";              Desc = "Generate HTML Report" }
    @{ Num = 8;  File = "08-CSVExportAndLogging.ps1";     Desc = "CSV Export & Log File Output" }
    @{ Num = 9;  File = "09-BackupAndDelete.ps1";         Desc = "Backup Profiles Before Deletion" }
    @{ Num = 10; File = "10-ConfigFileUsage.ps1";         Desc = "Load Settings from JSON Config File" }
    @{ Num = 11; File = "11-CorruptionDetection.ps1";     Desc = "Corrupted Profile Detection" }
    @{ Num = 12; File = "12-InteractiveMode.ps1";         Desc = "Interactive Profile Selection" }
    @{ Num = 13; File = "13-DeleteProfiles.ps1";          Desc = "Delete Profiles (with safety prompts)" }
    @{ Num = 14; File = "14-GUIMode.ps1";                 Desc = "Graphical User Interface (WPF GUI)" }
    @{ Num = 15; File = "15-ScheduledTask.ps1";           Desc = "Scheduled Task / Quiet Mode Example" }
    @{ Num = 16; File = "16-RemoteComputers.ps1";         Desc = "Remote Computer Targeting" }
)

foreach ($d in $demos) {
    $color = if ($d.Num -le 9) { "White" } else { "Gray" }
    Write-Host ("  [{0,2}] {1}" -f $d.Num, $d.Desc) -ForegroundColor $color
}

Write-Host ""
Write-Host "  [ 0] Exit" -ForegroundColor DarkGray
Write-Host ""
$choice = Read-Host "  Select a demo (0-$($demos.Count))"

if ($choice -eq '0' -or [string]::IsNullOrWhiteSpace($choice)) {
    Write-Host "  Exiting." -ForegroundColor Yellow
    exit 0
}

$selected = $demos | Where-Object { $_.Num -eq [int]$choice }
if ($selected) {
    $scriptPath = Join-Path $scriptDir $selected.File
    if (Test-Path $scriptPath) {
        Write-Host ""
        Write-Host "  Launching: $($selected.Desc)" -ForegroundColor Green
        Write-Host "  ────────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
        Write-Host ""
        & $scriptPath
    }
    else {
        Write-Host "  Script not found: $scriptPath" -ForegroundColor Red
    }
}
else {
    Write-Host "  Invalid selection." -ForegroundColor Red
}
