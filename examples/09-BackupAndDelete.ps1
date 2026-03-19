#Requires -Version 5.1
<#
.SYNOPSIS
    Demo: Backup Profiles Before Deletion

.DESCRIPTION
    Shows how DelprofPS can back up profiles to ZIP files before deleting
    them. This provides a safety net - if a profile is accidentally deleted,
    it can be restored from the backup archive.

    This demo runs in PREVIEW mode so nothing is actually deleted.
    To actually backup+delete, remove -Preview and keep -Delete.

    WHAT THIS DEMONSTRATES:
    - The -BackupPath parameter
    - The -Delete parameter (combined with -Preview for safety)
    - Backup-Profile function (ZIP compression)
    - How backup + delete work together

.NOTES
    SAFE TO RUN AS-IS - Preview mode, no actual deletion.
    To perform real backup+delete, edit the command below.
#>

$mainScript = Join-Path $PSScriptRoot '..\DelprofPS.ps1'
$backupDir = Join-Path $PSScriptRoot '..\DelprofPS_DemoBackups'

Write-Host "`n  DEMO: Backup Profiles Before Deletion" -ForegroundColor Cyan
Write-Host "  Backup directory: $backupDir" -ForegroundColor Gray
Write-Host "  Running in PREVIEW mode - no profiles will be deleted.`n" -ForegroundColor Green

# Preview mode - shows what WOULD happen
& $mainScript -DaysInactive 90 -Preview -BackupPath $backupDir -ShowSpace

Write-Host ""
Write-Host "  ────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  To actually backup and delete, run:" -ForegroundColor Yellow
Write-Host "    .\DelprofPS.ps1 -DaysInactive 90 -Delete -BackupPath `"$backupDir`"" -ForegroundColor White
Write-Host ""
