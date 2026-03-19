#Requires -Version 5.1
<#
.SYNOPSIS
    Demo: Graphical User Interface (WPF GUI)

.DESCRIPTION
    Launches the DelprofPS WPF-based graphical interface. The GUI provides
    visual access to all script functionality:

    - Connection tab: Local or remote computer targeting
    - Filters tab: Age, patterns, size, profile type
    - Actions tab: Preview/Delete mode, all options as checkboxes
    - Output tab: Backup, logging, HTML report, email settings
    - Real-time output console with progress bar

    WHAT THIS DEMONSTRATES:
    - The -UI parameter
    - Show-DelprofPSGUI WPF interface
    - Full visual control over all DelprofPS features

.NOTES
    INTERACTIVE - Opens a GUI window.
    No changes are made unless you click RUN in the GUI.
#>

$mainScript = Join-Path $PSScriptRoot '..\DelprofPS.ps1'

Write-Host "`n  DEMO: Graphical User Interface" -ForegroundColor Cyan
Write-Host "  Launching the DelprofPS WPF GUI...`n" -ForegroundColor Gray

& $mainScript -UI
