#Requires -Version 5.1
<#
.SYNOPSIS
    Delprof2-PS v2.0 - Enterprise-grade PowerShell replacement for Delprof2 with advanced features.

.DESCRIPTION
    A comprehensive user profile management tool that exceeds Delprof2 capabilities:
    
    CORE FEATURES:
    - Local and remote computer profile management
    - Multiple age calculation methods (NTUSER.DAT, ProfilePath, Registry, LastLogon)
    - Active session detection (prevents deleting logged-in users)
    - Flexible filtering (include/exclude by username, SID, or pattern)
    - Profile state detection (local, roaming, temporary, mandatory, corrupted)
    - Disk space calculation and reporting
    - Registry hive unloading before deletion
    - Comprehensive logging and CSV export
    
    ENTERPRISE FEATURES:
    - Interactive profile selection mode with visual menu
    - Test/Validate mode to check prerequisites without changes
    - Parallel processing for multiple computers
    - HTML report generation with professional styling
    - Windows Event Log integration for audit trails
    - Profile backup before deletion
    - JSON configuration file support
    - Email notifications for scheduled runs
    - Progress bars for long operations
    - Age-based color coding in output

.PARAMETER ComputerName
    Target computer(s). Defaults to localhost. Accepts multiple, pipeline input, and CSV files.

.PARAMETER DaysInactive
    Minimum days of inactivity for profile deletion. Default: 30

.PARAMETER AgeCalculation
    Method to determine profile age: NTUSER_DAT (default), ProfilePath, Registry, LastLogon, LastLogoff

.PARAMETER Include
    Wildcard pattern for usernames to include (e.g., "user*", "*admin*")

.PARAMETER Exclude
    Wildcard pattern for usernames to exclude (e.g., "Administrator*", "*service*")

.PARAMETER Delete
    Actually delete profiles. Without this, runs in dry-run/list mode.

.PARAMETER Force
    Skip confirmation prompts and ignore non-critical errors.

.PARAMETER UI
    Launches the graphical user interface for visual profile management.
    Provides a modern WPF-based GUI with access to all script functionality.

.PARAMETER Preview
    Shows what would be deleted without actually deleting. Performs a dry run
    that displays all profiles that match the criteria and would be removed.
    Combine with -Delete to see preview before actual deletion.

.PARAMETER IgnoreActiveSessions
    Allow deletion of profiles with active user sessions (DANGEROUS).

.PARAMETER UnloadHives
    Unload loaded registry hives before deletion (recommended).

.PARAMETER MaxRetries
    Number of retry attempts for locked files. Default: 3

.PARAMETER RetryDelaySeconds
    Seconds between retry attempts. Default: 2

.PARAMETER OutputPath
    Export results to CSV file.

.PARAMETER LogPath
    Write detailed log to file.

.PARAMETER Quiet
    Suppress console output (useful for scheduled tasks).

.PARAMETER ShowSpace
    Display disk space used by each profile.

.PARAMETER IncludeSystemProfiles
    Include system profiles (Default, Public, etc.) - EXTREME CAUTION.

.PARAMETER IncludeSpecialProfiles
    Include special accounts (SYSTEM, NetworkService, LocalService).

.PARAMETER MinProfileSizeMB
    Only consider profiles larger than specified MB.

.PARAMETER MaxProfileSizeMB
    Only consider profiles smaller than specified MB.

.PARAMETER IncludeCorrupted
    Include corrupted profiles in processing.

.PARAMETER FixCorruption
    Enable interactive corruption repair mode. Presents options to fix corrupted profiles:
    - Remove orphaned registry keys (for missing profile paths)
    - Delete corrupted profiles entirely
    - Recreate NTUSER.DAT from default template
    - Skip and continue. Requires -Interactive for full control.

.PARAMETER ProfileType
    Filter by profile type: Local, Roaming, Temporary, Mandatory, or All (default).

.PARAMETER Interactive
    Enable interactive mode for manual profile selection with visual menu.

.PARAMETER Test
    Test mode - validate prerequisites and connectivity without making changes.

.PARAMETER HtmlReport
    Generate professional HTML report at specified path.

.PARAMETER BackupPath
    Backup profiles to ZIP files before deletion (specify directory path).

.PARAMETER ConfigFile
    Load settings from JSON configuration file.

.PARAMETER UseParallel
    Use parallel processing for multiple computers.

.PARAMETER ThrottleLimit
    Maximum parallel threads when using -UseParallel. Default: 5

.PARAMETER SmtpServer
    SMTP server for email notifications.

.PARAMETER EmailTo
    Email recipient address for notifications.

.PARAMETER EmailFrom
    Email sender address. Default: delprofps@computername

.PARAMETER Detailed
    Show detailed folder breakdown for each profile (Documents, Downloads, Desktop, etc.)

.EXAMPLE
    # List all profiles older than 30 days on local computer (dry run)
    .\DelprofPS.ps1

.EXAMPLE
    # Delete profiles older than 60 days, excluding administrators
    .\DelprofPS.ps1 -DaysInactive 60 -Delete -Exclude "*admin*"

.EXAMPLE
    # Interactive mode - select profiles visually before deletion
    .\DelprofPS.ps1 -DaysInactive 90 -Interactive

.EXAMPLE
    # Test connectivity to remote computers without making changes
    .\DelprofPS.ps1 -ComputerName SERVER1,SERVER2,SERVER3 -Test

.EXAMPLE
    # List profiles on remote computers showing disk space with progress bar
    .\DelprofPS.ps1 -ComputerName SERVER1,SERVER2 -ShowSpace

.EXAMPLE
    # Enterprise deployment with full reporting
    .\DelprofPS.ps1 -ComputerName (Import-Csv servers.csv).Name `
        -Delete -DaysInactive 90 `
        -LogPath "C:\Logs\delprof.log" `
        -OutputPath "C:\Logs\results.csv" `
        -HtmlReport "C:\Logs\report.html" `
        -BackupPath "C:\Backups\Profiles" `
        -UnloadHives -ShowSpace

.EXAMPLE
    # Scheduled task with email notification
    .\DelprofPS.ps1 -DaysInactive 60 -Delete -Exclude "*admin*" `
        -SmtpServer "mail.company.com" `
        -EmailTo "admin@company.com" `
        -EmailFrom "delprofps@server01" `
        -Quiet -LogPath "C:\Logs\delprof.log"

.EXAMPLE
    # Use JSON configuration file
    .\DelprofPS.ps1 -ConfigFile "C:\Config\delprof.json" -Delete

.EXAMPLE
    # Parallel processing for many computers
    .\DelprofPS.ps1 -ComputerName (Get-Content servers.txt) `
        -UseParallel -ThrottleLimit 10 `
        -DaysInactive 120 -Delete

.EXAMPLE
    # Show detailed folder breakdown for each profile
    .\DelprofPS.ps1 -DaysInactive 60 -Detailed -ShowSpace

.EXAMPLE
    # Process with all safety checks disabled (EXTREME CAUTION)
    .\DelprofPS.ps1 -Delete -Force -IgnoreActiveSessions -IncludeSystemProfiles

.NOTES
    Version:        2.0.0
    Author:         Karl Lawrence
    Creation Date:  2024
    
    REQUIREMENTS:
    - PowerShell 5.1 or later
    - Administrative privileges on target computers
    - Remote management enabled for remote computer processing
    
    SAFETY FEATURES:
    - Default dry-run mode (must use -Delete to actually remove profiles)
    - Active session protection (skips logged-in users unless -IgnoreActiveSessions)
    - System profile protection (excludes Default, Public, SYSTEM, etc.)
    - Registry hive unloading before deletion (-UnloadHives)
    - Profile backup capability (-BackupPath)
    - Comprehensive logging for audit trails
    - Windows Event Log integration
    
    EVENT LOG IDs:
    - 1000: Script started
    - 1001: HTML report generated
    - 1002: Script completed
    - 1005: Error - admin rights required
    - 1010: Profile deleted
    
.LINK
    Original Delprof2: https://helgeklein.com/delprof2/
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
    [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [Alias('CN', 'MachineName', 'Server')]
    [string[]]$ComputerName = $env:COMPUTERNAME,

    [Parameter()]
    [Alias('Age', 'Days')]
    [int]$DaysInactive = 30,

    [Parameter()]
    [ValidateSet('NTUSER_DAT', 'ProfilePath', 'Registry', 'LastLogon', 'LastLogoff')]
    [string]$AgeCalculation = 'NTUSER_DAT',

    [Parameter()]
    [string[]]$Include,

    [Parameter()]
    [string[]]$Exclude,

    [Parameter()]
    [switch]$Delete,

    [Parameter()]
    [switch]$Force,

    [Parameter()]
    [switch]$UI,

    [Parameter()]
    [switch]$Preview,

    [Parameter()]
    [switch]$IgnoreActiveSessions,

    [Parameter()]
    [switch]$UnloadHives,

    [Parameter()]
    [int]$MaxRetries = 3,

    [Parameter()]
    [int]$RetryDelaySeconds = 2,

    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [string]$LogPath,

    [Parameter()]
    [switch]$Quiet,

    [Parameter()]
    [switch]$ShowSpace,

    [Parameter()]
    [switch]$IncludeSystemProfiles,

    [Parameter()]
    [switch]$IncludeSpecialProfiles,

    [Parameter()]
    [long]$MinProfileSizeMB,

    [Parameter()]
    [long]$MaxProfileSizeMB,

    [Parameter()]
    [switch]$IncludeCorrupted,

    [Parameter()]
    [switch]$FixCorruption,

    [Parameter()]
    [ValidateSet('Local', 'Roaming', 'Temporary', 'Mandatory', 'All')]
    [string]$ProfileType = 'All',

    [Parameter()]
    [switch]$Interactive,

    [Parameter()]
    [switch]$Test,

    [Parameter()]
    [string]$HtmlReport,

    [Parameter()]
    [string]$BackupPath,

    [Parameter()]
    [string]$ConfigFile,

    [Parameter()]
    [switch]$UseParallel,

    [Parameter()]
    [int]$ThrottleLimit = 5,

    [Parameter()]
    [string]$SmtpServer,

    [Parameter()]
    [string]$EmailTo,

    [Parameter()]
    [string]$EmailFrom = "delprofps@$env:COMPUTERNAME",

    [Parameter()]
    [switch]$Detailed
)

begin {
    #region UI Mode Check
    if ($UI) {
        Show-DelprofPSGUI
        return
    }
    #endregion

    #region Initialization
    $script:StartTime = Get-Date
    $script:Version = '2.0.0'
    $script:TotalProfilesProcessed = 0
    $script:TotalProfilesDeleted = 0
    $script:TotalSpaceFreed = 0
    $script:Results = [System.Collections.Generic.List[object]]::new()
    $script:ComputerQueue = [System.Collections.Generic.List[string]]::new()

    # Well-known SIDs to protect
    $script:ProtectedSIDs = @(
        'S-1-5-18',      # SYSTEM
        'S-1-5-19',      # LOCAL SERVICE
        'S-1-5-20',      # NETWORK SERVICE
        'S-1-5-21-%-500' # Built-in Administrator (domain)
    )

    $script:SystemProfileNames = @(
        'Default', 'Public', 'Default User', 'All Users',
        'systemprofile', 'LocalService', 'NetworkService',
        'Administrator', 'Guest'
    )

    # Error action preference
    if ($Force) {
        $ErrorActionPreference = 'SilentlyContinue'
    } else {
        $ErrorActionPreference = 'Stop'
    }

    # Load configuration file if specified
    if ($ConfigFile -and (Test-Path $ConfigFile)) {
        try {
            $config = Get-Content $ConfigFile | ConvertFrom-Json
            # Apply config values only if not already specified via parameters
            if (-not $PSBoundParameters.ContainsKey('DaysInactive') -and $config.DaysInactive) { $DaysInactive = $config.DaysInactive }
            if (-not $PSBoundParameters.ContainsKey('Exclude') -and $config.Exclude) { $Exclude = $config.Exclude }
            if (-not $PSBoundParameters.ContainsKey('Include') -and $config.Include) { $Include = $config.Include }
            if (-not $PSBoundParameters.ContainsKey('MaxRetries') -and $config.MaxRetries) { $MaxRetries = $config.MaxRetries }
            Write-DPLog -Message "Configuration loaded from $ConfigFile" -Level 'INFO'
        }
        catch {
            Write-DPLog -Message "Failed to load config file: $($_.Exception.Message)" -Level 'ERROR'
        }
    }
    #endregion

    #region Logging Functions
    function Write-DPLog {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Message,
            
            [Parameter()]
            [ValidateSet('INFO', 'WARNING', 'ERROR', 'SUCCESS', 'VERBOSE')]
            [string]$Level = 'INFO'
        )
        
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = "[$timestamp] [$Level] $Message"
        
        # Write to log file if specified
        if ($LogPath) {
            try {
                Add-Content -Path $LogPath -Value $logEntry -ErrorAction SilentlyContinue
            }
            catch {
                # Silent fail for logging errors
            }
        }
        
        # Console output
        if (-not $Quiet) {
            switch ($Level) {
                'ERROR'   { Write-Host $logEntry -ForegroundColor Red }
                'WARNING' { Write-Host $logEntry -ForegroundColor Yellow }
                'SUCCESS' { Write-Host $logEntry -ForegroundColor Green }
                'VERBOSE' { Write-Verbose $Message }
                default   { Write-Host $logEntry }
            }
        }
    }

    function Write-DPHeader {
        if (-not $Quiet) {
            Write-Host "`n" + ('=' * 80) -ForegroundColor Cyan
            Write-Host " Delprof2-PS v$script:Version - User Profile Management Tool" -ForegroundColor Cyan
            $modeText = if ($Preview) { 'PREVIEW/SIMULATION' } elseif ($Delete) { 'DELETE' } else { 'LIST/ANALYZE' }
            $modeColor = if ($Preview) { 'Magenta' } elseif ($Delete) { 'Red' } else { 'Green' }
            Write-Host " Mode: $modeText" -ForegroundColor $modeColor
            Write-Host " Criteria: Profiles older than $DaysInactive days" -ForegroundColor Cyan
            Write-Host ('=' * 80) + "`n" -ForegroundColor Cyan
        }
        Write-DPLog -Message "Script started. Version: $script:Version, Delete mode: $Delete, Days inactive: $DaysInactive" -Level 'INFO'
    }
    #endregion

    #region Utility Functions
    function Test-AdminRights {
        $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }

    function ConvertTo-UserName {
        param([string]$SID)
        try {
            $securityIdentifier = New-Object System.Security.Principal.SecurityIdentifier($SID)
            $ntAccount = $securityIdentifier.Translate([System.Security.Principal.NTAccount])
            return $ntAccount.Value
        }
        catch {
            return $null
        }
    }

    function Test-IsProtectedProfile {
        param(
            [string]$UserName,
            [string]$SID
        )
        
        # Check system profile names
        foreach ($sysName in $script:SystemProfileNames) {
            if ($UserName -like "*\$sysName" -or $UserName -eq $sysName) {
                return $true
            }
        }
        
        # Check protected SIDs
        foreach ($protectedSid in $script:ProtectedSIDs) {
            if ($protectedSid -match '%') {
                # Pattern match for domain admin
                $pattern = $protectedSid -replace '%', '\d+'
                if ($SID -match $pattern) { return $true }
            } else {
                if ($SID -eq $protectedSid) { return $true }
            }
        }
        
        # Special accounts
        if (-not $IncludeSpecialProfiles) {
            if ($SID -in @('S-1-5-18', 'S-1-5-19', 'S-1-5-20')) {
                return $true
            }
        }
        
        return $false
    }

    function Get-ProfileFolderSize {
        param([string]$Path)
        
        if (-not (Test-Path $Path)) { return 0 }
        
        try {
            $size = (Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue | 
                Measure-Object -Property Length -Sum).Sum
            return $size
        }
        catch {
            return -1  # Indicates error calculating size
        }
    }

    function Get-ProfileSizeBreakdown {
        param([string]$ProfilePath)
        
        $breakdown = @{}
        $folders = @('Documents', 'Downloads', 'Desktop', 'AppData', 'Pictures', 'Videos', 'Music')
        
        foreach ($folder in $folders) {
            $folderPath = Join-Path $ProfilePath $folder
            if (Test-Path $folderPath) {
                $size = Get-ProfileFolderSize -Path $folderPath
                $breakdown[$folder] = Format-Bytes -Bytes $size
            }
            else {
                $breakdown[$folder] = 'N/A'
            }
        }
        
        return $breakdown
    }

    function Test-ProfileLockedFiles {
        param([string]$ProfilePath)
        
        $lockedFiles = @()
        try {
            $files = Get-ChildItem -Path $ProfilePath -Recurse -File -ErrorAction SilentlyContinue | Select-Object -First 10
            foreach ($file in $files) {
                try {
                    $stream = [System.IO.File]::Open($file.FullName, 'Open', 'Read', 'None')
                    $stream.Close()
                }
                catch {
                    $lockedFiles += $file.FullName
                }
            }
        }
        catch {
            # Ignore errors
        }
        return $lockedFiles
    }

    function Format-Bytes {
        param([long]$Bytes)
        
        if ($Bytes -lt 0) { return 'Error' }
        if ($Bytes -eq 0) { return '0 B' }
        
        $sizes = @('B', 'KB', 'MB', 'GB', 'TB')
        $order = [math]::Floor([math]::Log($Bytes, 1024))
        $order = [math]::Min($order, $sizes.Count - 1)
        
        $formatted = [math]::Round($Bytes / [math]::Pow(1024, $order), 2)
        return "$formatted $($sizes[$order])"
    }

    function Get-AgeColor {
        param([int]$AgeInDays)
        if ($AgeInDays -lt 30) { return 'Green' }
        elseif ($AgeInDays -lt 90) { return 'Yellow' }
        elseif ($AgeInDays -lt 180) { return 'Magenta' }
        else { return 'Red' }
    }

    function Backup-Profile {
        param(
            [string]$SourcePath,
            [string]$UserName
        )
        
        if (-not $BackupPath) { return $true }
        
        try {
            $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $backupFile = Join-Path $BackupPath "$($UserName)_$timestamp.zip"
            
            if (-not (Test-Path $BackupPath)) {
                New-Item -ItemType Directory -Path $BackupPath -Force | Out-Null
            }
            
            Compress-Archive -Path $SourcePath -DestinationPath $backupFile -CompressionLevel Optimal -Force
            Write-DPLog -Message "Profile backed up to $backupFile" -Level 'SUCCESS'
            return $true
        }
        catch {
            Write-DPLog -Message "Failed to backup profile: $($_.Exception.Message)" -Level 'ERROR'
            return $false
        }
    }

    function Send-NotificationEmail {
        param([hashtable]$Summary)
        
        if (-not $SmtpServer -or -not $EmailTo) { return }
        
        try {
            $subject = "Delprof2-PS Report - $($Summary.ProfilesDeleted) profiles deleted"
            $body = @"
Delprof2-PS has completed processing.

Summary:
- Computers processed: $($Summary.Computers)
- Profiles processed: $($Summary.ProfilesProcessed)
- Profiles deleted: $($Summary.ProfilesDeleted)
- Space freed: $($Summary.SpaceFreed)
- Duration: $($Summary.Duration)

Report generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
"@
            
            Send-MailMessage -SmtpServer $SmtpServer -To $EmailTo -From $EmailFrom -Subject $subject -Body $body
            Write-DPLog -Message "Notification email sent to $EmailTo" -Level 'SUCCESS'
        }
        catch {
            Write-DPLog -Message "Failed to send email: $($_.Exception.Message)" -Level 'ERROR'
        }
    }
    #endregion

    #region Event Logging
    function Write-EventLogEntry {
        param(
            [string]$Message,
            [ValidateSet('Information', 'Warning', 'Error')]
            [string]$EntryType = 'Information',
            [int]$EventId = 1000
        )
        
        try {
            $source = 'Delprof2PS'
            $logName = 'Application'
            
            # Create event source if it doesn't exist
            if (-not [System.Diagnostics.EventLog]::SourceExists($source)) {
                try {
                    New-EventLog -LogName $logName -Source $source -ErrorAction Stop
                }
                catch {
                    # May not have permission to create source
                    return
                }
            }
            
            Write-EventLog -LogName $logName -Source $source -EventId $EventId -EntryType $EntryType -Message $Message -ErrorAction SilentlyContinue
        }
        catch {
            # Silent fail - event logging is optional
        }
    }
    #endregion

    #region HTML Reporting
    function Export-HtmlReport {
        param(
            [string]$Path,
            [array]$Results,
            [hashtable]$Summary
        )
        
        $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Delprof2-PS Report</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 1400px; margin: 0 auto; background: white; padding: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #0078d4; border-bottom: 2px solid #0078d4; padding-bottom: 10px; }
        h2 { color: #333; margin-top: 30px; }
        .summary { background: #f0f8ff; padding: 15px; border-radius: 5px; margin: 20px 0; }
        .summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; }
        .summary-item { background: white; padding: 10px; border-radius: 3px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
        .summary-label { font-weight: bold; color: #666; font-size: 0.9em; }
        .summary-value { font-size: 1.2em; color: #0078d4; margin-top: 5px; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th { background: #0078d4; color: white; padding: 12px; text-align: left; }
        td { padding: 10px; border-bottom: 1px solid #ddd; }
        tr:hover { background: #f5f5f5; }
        .deleted { background-color: #d4edda; }
        .active { background-color: #fff3cd; }
        .error { background-color: #f8d7da; }
        .badge { padding: 3px 8px; border-radius: 3px; font-size: 0.85em; font-weight: bold; }
        .badge-local { background: #e3f2fd; color: #1565c0; }
        .badge-roaming { background: #f3e5f5; color: #7b1fa2; }
        .badge-temp { background: #fff3e0; color: #e65100; }
        .badge-mandatory { background: #fce4ec; color: #c2185b; }
        .badge-corrupted { background: #ffebee; color: #c62828; }
        .footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; color: #666; font-size: 0.9em; }
        .chart-container { margin: 20px 0; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Delprof2-PS Report</h1>
        <p>Generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        
        <div class="summary">
            <h2>Summary</h2>
            <div class="summary-grid">
                <div class="summary-item">
                    <div class="summary-label">Computers Processed</div>
                    <div class="summary-value">$($Summary.Computers)</div>
                </div>
                <div class="summary-item">
                    <div class="summary-label">Profiles Processed</div>
                    <div class="summary-value">$($Summary.ProfilesProcessed)</div>
                </div>
                <div class="summary-item">
                    <div class="summary-label">Profiles Deleted</div>
                    <div class="summary-value">$($Summary.ProfilesDeleted)</div>
                </div>
                <div class="summary-item">
                    <div class="summary-label">Space Freed</div>
                    <div class="summary-value">$($Summary.SpaceFreed)</div>
                </div>
                <div class="summary-item">
                    <div class="summary-label">Duration</div>
                    <div class="summary-value">$($Summary.Duration)</div>
                </div>
            </div>
        </div>
        
        <h2>Profile Details</h2>
        <table>
            <thead>
                <tr>
                    <th>Computer</th>
                    <th>User</th>
                    <th>Profile Type</th>
                    <th>Last Used</th>
                    <th>Age (Days)</th>
                    <th>Size</th>
                    <th>Status</th>
                </tr>
            </thead>
            <tbody>
"@
        
        foreach ($result in $Results) {
            $rowClass = ''
            if ($result.Deleted) { $rowClass = 'deleted' }
            elseif ($result.IsActiveSession) { $rowClass = 'active' }
            elseif ($result.Error) { $rowClass = 'error' }
            
            $typeClass = switch ($result.ProfileType) {
                'Local' { 'badge-local' }
                'Roaming' { 'badge-roaming' }
                'Temporary' { 'badge-temp' }
                'Mandatory' { 'badge-mandatory' }
                default { 'badge-corrupted' }
            }
            
            $status = if ($result.Deleted) { 'Deleted' } elseif ($result.IsActiveSession) { 'Active' } elseif ($result.Error) { 'Error' } else { 'Kept' }
            
            $html += "                <tr class='$rowClass'>" +
                "<td>$($result.ComputerName)</td>" +
                "<td>$($result.UserName)</td>" +
                "<td><span class='badge $typeClass'>$($result.ProfileType)</span></td>" +
                "<td>$($result.LastUsed)</td>" +
                "<td>$($result.AgeInDays)</td>" +
                "<td>$($result.SizeFormatted)</td>" +
                "<td>$status</td></tr>`n"
        }
        
        $html += @"
            </tbody>
        </table>
        
        <div class="footer">
            <p>Report generated by Delprof2-PS v$script:Version</p>
        </div>
    </div>
</body>
</html>
"@
        
        try {
            $html | Out-File -FilePath $Path -Encoding UTF8 -Force
            Write-DPLog -Message "HTML report saved to $Path" -Level 'SUCCESS'
            Write-EventLogEntry -Message "HTML report generated: $Path" -EntryType Information -EventId 1001
        }
        catch {
            Write-DPLog -Message "Failed to save HTML report: $($_.Exception.Message)" -Level 'ERROR'
        }
    }
    #endregion

    #region Interactive Mode
    function Select-ProfilesInteractive {
        param([array]$Profiles)
        
        if (-not $Profiles) { return @() }
        
        Write-Host "`n=== INTERACTIVE PROFILE SELECTION ===" -ForegroundColor Cyan
        Write-Host "Use arrow keys to navigate, Space to toggle selection, Enter to confirm`n" -ForegroundColor Gray
        
        $selected = New-Object System.Collections.Generic.List[int]
        $currentIndex = 0
        
        function Show-Menu {
            Clear-Host
            Write-Host "=== SELECT PROFILES TO DELETE ===" -ForegroundColor Cyan
            Write-Host "[Space] Select/Deselect  [Enter] Confirm  [A] Select All  [N] Select None  [Q] Quit`n" -ForegroundColor Gray
            
            for ($i = 0; $i -lt $Profiles.Count; $i++) {
                $prof = $Profiles[$i]
                $prefix = if ($i -eq $currentIndex) { '>' } else { ' ' }
                $marker = if ($selected -contains $i) { '[X]' } else { '[ ]' }
                $color = if ($i -eq $currentIndex) { 'Yellow' } elseif ($prof.IsActiveSession) { 'Red' } else { 'White' }
                $active = if ($prof.IsActiveSession) { ' [ACTIVE]' } else { '' }
                
                Write-Host " $prefix $marker $($prof.UserName) - $($prof.AgeInDays) days - $($prof.SizeFormatted)$active" -ForegroundColor $color
            }
            
            Write-Host "`nSelected: $($selected.Count) profiles ($(($Profiles | Where-Object { $selected -contains $Profiles.IndexOf($_) } | Measure-Object -Property SizeBytes -Sum).Sum / 1MB -as [int]) MB total)" -ForegroundColor Green
        }
        
        do {
            Show-Menu
            $key = [Console]::ReadKey($true)
            
            switch ($key.Key) {
                'UpArrow' { if ($currentIndex -gt 0) { $currentIndex-- } }
                'DownArrow' { if ($currentIndex -lt $Profiles.Count - 1) { $currentIndex++ } }
                'Spacebar' {
                    if ($selected -contains $currentIndex) {
                        $selected.Remove($currentIndex)
                    } else {
                        $selected.Add($currentIndex)
                    }
                }
                'A' { $selected = [System.Collections.Generic.List[int]](0..($Profiles.Count - 1)) }
                'N' { $selected.Clear() }
                'Q' { return @() }
            }
        } while ($key.Key -ne 'Enter')
        
        return $selected | ForEach-Object { $Profiles[$_] }
    }
    #endregion
    function Test-ComputerConnection {
        param([string]$Computer)
        
        try {
            # Test connection
            $ping = Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction SilentlyContinue
            if (-not $ping) {
                return @{ Success = $false; Error = 'Host unreachable' }
            }
            
            # Test admin access
            $testPath = "\\$Computer\ADMIN$"
            if (-not (Test-Path $testPath -ErrorAction SilentlyContinue)) {
                return @{ Success = $false; Error = 'Admin share not accessible' }
            }
            
            return @{ Success = $true; Error = $null }
        }
        catch {
            return @{ Success = $false; Error = $_.Exception.Message }
        }
    }

    function Invoke-RemoteCommand {
        param(
            [string]$ComputerName,
            [scriptblock]$ScriptBlock,
            [hashtable]$ArgumentList = @{}
        )
        
        try {
            if ($ComputerName -eq $env:COMPUTERNAME -or $ComputerName -eq 'localhost' -or $ComputerName -eq '.') {
                # Local execution
                $result = & $ScriptBlock @ArgumentList
                return @{ Success = $true; Data = $result; Error = $null }
            }
            else {
                # Remote execution using Invoke-Command
                $session = New-PSSession -ComputerName $ComputerName -ErrorAction Stop
                $result = Invoke-Command -Session $session -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList
                Remove-PSSession -Session $session
                return @{ Success = $true; Data = $result; Error = $null }
            }
        }
        catch {
            return @{ Success = $false; Data = $null; Error = $_.Exception.Message }
        }
    }
    #endregion

    #region Profile Analysis Functions
    function Get-ActiveSessions {
        param([string]$ComputerName)
        
        try {
            $sessions = @()
            
            # Method 1: Query user.exe (quser)
            try {
                $quserOutput = quser /server:$ComputerName 2>$null
                if ($quserOutput) {
                    $sessions += $quserOutput | Select-String -Pattern '(\S+)\s+\d+\s+(\S+)' | ForEach-Object {
                        $matches[1]
                    }
                }
            }
            catch {
                # quser may not be available
            }
            
            # Method 2: Get logged on users via WMI
            try {
                $loggedOn = Get-WmiObject -Class Win32_LoggedOnUser -ComputerName $ComputerName -ErrorAction SilentlyContinue |
                    ForEach-Object { 
                        $_.Antecedent -match 'Domain="([^"]+)",Name="([^"]+)"' | Out-Null
                        "$($matches[1])\$($matches[2])"
                    } | Select-Object -Unique
                $sessions += $loggedOn
            }
            catch {
                # WMI may fail
            }
            
            # Method 3: Get explorer.exe processes
            try {
                $explorerUsers = Get-WmiObject -Class Win32_Process -ComputerName $ComputerName -Filter "Name='explorer.exe'" -ErrorAction SilentlyContinue |
                    ForEach-Object { 
                        $owner = $_.GetOwner()
                        if ($owner) { "$($owner.Domain)\$($owner.User)" }
                    } | Select-Object -Unique
                $sessions += $explorerUsers
            }
            catch {
                # Process query may fail
            }
            
            return $sessions | Select-Object -Unique
        }
        catch {
            return @()
        }
    }

    function Get-ProfileAge {
        param(
            [string]$ProfilePath,
            [string]$SID,
            [string]$Method,
            [string]$ComputerName
        )
        
        $lastUsed = $null
        $source = 'Unknown'
        
        switch ($Method) {
            'NTUSER_DAT' {
                $ntUserDat = Join-Path $ProfilePath 'NTUSER.DAT'
                if (Test-Path $ntUserDat -ErrorAction SilentlyContinue) {
                    try {
                        $lastUsed = (Get-Item $ntUserDat).LastWriteTime
                        $source = 'NTUSER.DAT'
                    }
                    catch {
                        $lastUsed = $null
                    }
                }
                
                if (-not $lastUsed) {
                    # Fallback to profile path
                    try {
                        $lastUsed = (Get-Item $ProfilePath).LastWriteTime
                        $source = 'ProfilePath'
                    }
                    catch {
                        $lastUsed = [DateTime]::MinValue
                        $source = 'Error'
                    }
                }
            }
            
            'ProfilePath' {
                if (Test-Path $ProfilePath -ErrorAction SilentlyContinue) {
                    try {
                        $lastUsed = (Get-Item $ProfilePath).LastWriteTime
                        $source = 'ProfilePath'
                    }
                    catch {
                        $lastUsed = [DateTime]::MinValue
                        $source = 'Error'
                    }
                }
            }
            
            'Registry' {
                try {
                    $profileKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$SID"
                    if (Test-Path $profileKey) {
                        $profileInfo = Get-ItemProperty $profileKey
                        # Try to find timestamp in various registry values
                        if ($profileInfo.LocalProfileLoadTimeHigh) {
                            # Convert FILETIME if available
                            $lastUsed = [DateTime]::FromFileTime($profileInfo.LocalProfileLoadTimeLow + ($profileInfo.LocalProfileLoadTimeHigh -shl 32))
                            $source = 'RegistryLoadTime'
                        }
                        else {
                            $lastUsed = (Get-Item $profileKey).LastWriteTime
                            $source = 'RegistryKey'
                        }
                    }
                }
                catch {
                    $lastUsed = [DateTime]::MinValue
                    $source = 'Error'
                }
            }
            
            'LastLogon' {
                try {
                    $userName = ConvertTo-UserName -SID $SID
                    if ($userName) {
                        $adUser = [ADSI]"WinNT://$($userName -replace '\\','/'),user"
                        if ($adUser.LastLogin) {
                            $lastUsed = $adUser.LastLogin
                            $source = 'LastLogon'
                        }
                    }
                }
                catch {
                    $lastUsed = $null
                }
                
                if (-not $lastUsed) {
                    # Fallback to NTUSER.DAT
                    $ntUserDat = Join-Path $ProfilePath 'NTUSER.DAT'
                    if (Test-Path $ntUserDat -ErrorAction SilentlyContinue) {
                        $lastUsed = (Get-Item $ntUserDat).LastWriteTime
                        $source = 'NTUSER.DAT (fallback)'
                    }
                }
            }
        }
        
        if (-not $lastUsed) {
            $lastUsed = [DateTime]::MinValue
            $source = 'Unknown'
        }
        
        return @{ LastUsed = $lastUsed; Source = $source }
    }

    function Get-ProfileType {
        param([hashtable]$ProfileInfo)
        
        $profilePath = $ProfileInfo.ProfilePath
        
        # Check for roaming
        if ($ProfileInfo.RoamingConfigured -eq 1) {
            return 'Roaming'
        }
        
        # Check for mandatory
        if ($profilePath -like '*.man' -or (Test-Path "$profilePath.man" -ErrorAction SilentlyContinue)) {
            return 'Mandatory'
        }
        
        # Check for temporary
        if ($ProfileInfo.TemporaryProfile -eq 1 -or $profilePath -like '*TEMP*') {
            return 'Temporary'
        }
        
        # Check if corrupted
        if (-not (Test-Path $profilePath -ErrorAction SilentlyContinue)) {
            return 'Corrupted (Path Missing)'
        }
        
        if (-not (Test-Path (Join-Path $profilePath 'NTUSER.DAT') -ErrorAction SilentlyContinue)) {
            return 'Corrupted (No NTUSER.DAT)'
        }
        
        return 'Local'
    }

    function Repair-CorruptedProfile {
        [CmdletBinding(SupportsShouldProcess = $true)]
        param(
            [string]$ComputerName,
            [string]$SID,
            [string]$UserName,
            [string]$ProfilePath,
            [string]$CorruptionType
        )
        
        $fixed = $false
        $actionTaken = 'No action taken'
        
        Write-DPLog -Message "Corruption detected for $UserName ($SID)`: $CorruptionType" -Level 'WARNING'
        
        if (-not $Quiet) {
            Write-Host "`n  CORRUPTION DETECTED" -ForegroundColor Red -BackgroundColor Black
            Write-Host "  User: $UserName" -ForegroundColor Yellow
            Write-Host "  SID: $SID" -ForegroundColor Yellow
            Write-Host "  Path: $ProfilePath" -ForegroundColor Yellow
            Write-Host "  Issue: $CorruptionType" -ForegroundColor Red
        }
        
        # Determine available repair options based on corruption type
        $repairOptions = @()
        
        switch ($CorruptionType) {
            'Corrupted (Path Missing)' {
                $repairOptions = @(
                    @{ Key = 'R'; Label = 'Remove orphaned registry key'; Description = 'Deletes the stale registry entry pointing to non-existent folder'; Risk = 'Low' }
                    @{ Key = 'S'; Label = 'Skip'; Description = 'Leave as-is and continue'; Risk = 'None' }
                )
            }
            'Corrupted (No NTUSER.DAT)' {
                $repairOptions = @(
                    @{ Key = 'D'; Label = 'Delete entire profile'; Description = 'Remove registry key and delete profile folder'; Risk = 'Medium' }
                    @{ Key = 'R'; Label = 'Recreate NTUSER.DAT'; Description = 'Copy default NTUSER.DAT to fix the profile'; Risk = 'Low' }
                    @{ Key = 'F'; Label = 'Force-remove folder only'; Description = 'Delete folder contents but keep registry key'; Risk = 'Medium' }
                    @{ Key = 'S'; Label = 'Skip'; Description = 'Leave as-is and continue'; Risk = 'None' }
                )
            }
        }
        
        # Display interactive menu
        if (-not $Quiet) {
            Write-Host "`n  Available repair options:" -ForegroundColor Cyan
            foreach ($opt in $repairOptions) {
                $riskColor = switch ($opt.Risk) {
                    'Low' { 'Green' }
                    'Medium' { 'Yellow' }
                    'High' { 'Red' }
                    default { 'White' }
                }
                Write-Host "    [$($opt.Key)] $($opt.Label)" -ForegroundColor White -NoNewline
                Write-Host " [Risk: $($opt.Risk)]" -ForegroundColor $riskColor
                Write-Host "        $($opt.Description)" -ForegroundColor Gray
            }
            Write-Host "`n  NOTE: This operation requires administrative privileges." -ForegroundColor Magenta
            Write-Host "  The user will NOT be able to log in until repaired." -ForegroundColor Magenta
        }
        
        # Get user choice (respecting Force parameter for non-interactive)
        $choice = $null
        if ($Force -and -not $Interactive) {
            # In Force mode without Interactive, default to Skip to prevent accidental damage
            $choice = 'S'
            if (-not $Quiet) {
                Write-Host "`n  Force mode active - defaulting to Skip. Use -FixCorruption with -Interactive for manual control." -ForegroundColor Yellow
            }
        }
        else {
            # Interactive prompt
            $validChoices = $repairOptions.Key
            while ($choice -notin $validChoices) {
                if (-not $Quiet) {
                    Write-Host "`n  Enter your choice [$(($validChoices -join '/'))]: " -ForegroundColor Cyan -NoNewline
                }
                $choice = Read-Host
                $choice = $choice.ToUpper().Trim()
            }
        }
        
        # Execute chosen action
        switch ($choice) {
            'R' {  # Remove registry key or Recreate NTUSER.DAT
                if ($CorruptionType -eq 'Corrupted (Path Missing)') {
                    # Remove orphaned registry key
                    $targetDesc = "Remove orphaned registry key for $UserName ($SID)"
                    if ($PSCmdlet.ShouldProcess($targetDesc, 'Remove Registry Key')) {
                        try {
                            $regKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$SID"
                            if ($ComputerName -eq $env:COMPUTERNAME -or $ComputerName -eq 'localhost' -or $ComputerName -eq '.') {
                                if (Test-Path $regKey) {
                                    Remove-Item -Path $regKey -Recurse -Force
                                }
                            }
                            else {
                                $scriptBlock = {
                                    param($sid)
                                    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
                                    if (Test-Path "$regPath\$sid") {
                                        Remove-Item -Path "$regPath\$sid" -Recurse -Force
                                    }
                                }
                                Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $SID
                            }
                            $fixed = $true
                            $actionTaken = 'Removed orphaned registry key'
                            Write-DPLog -Message "Removed orphaned registry key for $UserName" -Level 'SUCCESS'
                        }
                        catch {
                            $actionTaken = "Failed to remove registry key: $($_.Exception.Message)"
                            Write-DPLog -Message "Failed to remove registry key for $UserName`: $($_.Exception.Message)" -Level 'ERROR'
                        }
                    }
                }
                else {
                    # Recreate NTUSER.DAT
                    $targetDesc = "Recreate NTUSER.DAT for $UserName at $ProfilePath"
                    if ($PSCmdlet.ShouldProcess($targetDesc, 'Recreate NTUSER.DAT')) {
                        try {
                            $defaultNtUser = "C:\Users\Default\NTUSER.DAT"
                            if (Test-Path $defaultNtUser) {
                                $targetNtUser = Join-Path $ProfilePath 'NTUSER.DAT'
                                Copy-Item -Path $defaultNtUser -Destination $targetNtUser -Force
                                
                                # Set proper permissions (simplified - full ACL would require more code)
                                $acl = Get-Acl $targetNtUser
                                $sidObject = New-Object System.Security.Principal.SecurityIdentifier($SID)
                                $rule = New-Object System.Security.AccessControl.FileSystemAccessRule($sidObject, 'FullControl', 'Allow')
                                $acl.SetAccessRule($rule)
                                Set-Acl $targetNtUser $acl
                                
                                $fixed = $true
                                $actionTaken = 'Recreated NTUSER.DAT from default'
                                Write-DPLog -Message "Recreated NTUSER.DAT for $UserName" -Level 'SUCCESS'
                            }
                            else {
                                $actionTaken = 'Default NTUSER.DAT not found'
                                Write-DPLog -Message "Default NTUSER.DAT not found for copying" -Level 'ERROR'
                            }
                        }
                        catch {
                            $actionTaken = "Failed to recreate NTUSER.DAT: $($_.Exception.Message)"
                            Write-DPLog -Message "Failed to recreate NTUSER.DAT for $UserName`: $($_.Exception.Message)" -Level 'ERROR'
                        }
                    }
                }
            }
            
            'D' {  # Delete entire profile
                $targetDesc = "Delete entire corrupted profile for $UserName ($SID) at $ProfilePath"
                if ($PSCmdlet.ShouldProcess($targetDesc, 'Delete Corrupted Profile')) {
                    try {
                        # Remove registry key
                        $regKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$SID"
                        if ($ComputerName -eq $env:COMPUTERNAME -or $ComputerName -eq 'localhost' -or $ComputerName -eq '.') {
                            if (Test-Path $regKey) {
                                Remove-Item -Path $regKey -Recurse -Force
                            }
                            if (Test-Path $ProfilePath) {
                                Remove-Item -Path $ProfilePath -Recurse -Force
                            }
                        }
                        else {
                            $scriptBlock = {
                                param($sid, $profilePath)
                                $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
                                if (Test-Path "$regPath\$sid") {
                                    Remove-Item -Path "$regPath\$sid" -Recurse -Force
                                }
                                if (Test-Path $profilePath) {
                                    Remove-Item -Path $profilePath -Recurse -Force
                                }
                            }
                            Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $SID, $ProfilePath
                        }
                        $fixed = $true
                        $actionTaken = 'Deleted entire corrupted profile'
                        Write-DPLog -Message "Deleted corrupted profile for $UserName" -Level 'SUCCESS'
                    }
                    catch {
                        $actionTaken = "Failed to delete profile: $($_.Exception.Message)"
                        Write-DPLog -Message "Failed to delete corrupted profile for $UserName`: $($_.Exception.Message)" -Level 'ERROR'
                    }
                }
            }
            
            'F' {  # Force-remove folder only
                $targetDesc = "Remove profile folder for $UserName at $ProfilePath"
                if ($PSCmdlet.ShouldProcess($targetDesc, 'Remove Folder')) {
                    try {
                        if ($ComputerName -eq $env:COMPUTERNAME -or $ComputerName -eq 'localhost' -or $ComputerName -eq '.') {
                            if (Test-Path $ProfilePath) {
                                # Remove read-only attributes
                                Get-ChildItem -Path $ProfilePath -Recurse -Force -ErrorAction SilentlyContinue | 
                                    ForEach-Object { $_.Attributes = $_.Attributes -band -bnot [System.IO.FileAttributes]::ReadOnly }
                                Remove-Item -Path $ProfilePath -Recurse -Force
                            }
                        }
                        else {
                            $scriptBlock = {
                                param($profilePath)
                                if (Test-Path $profilePath) {
                                    Get-ChildItem -Path $profilePath -Recurse -Force -ErrorAction SilentlyContinue | 
                                        ForEach-Object { $_.Attributes = $_.Attributes -band -bnot [System.IO.FileAttributes]::ReadOnly }
                                    Remove-Item -Path $profilePath -Recurse -Force
                                }
                            }
                            Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $ProfilePath
                        }
                        $fixed = $true
                        $actionTaken = 'Removed profile folder only'
                        Write-DPLog -Message "Removed profile folder for $UserName (registry key preserved)" -Level 'SUCCESS'
                    }
                    catch {
                        $actionTaken = "Failed to remove folder: $($_.Exception.Message)"
                        Write-DPLog -Message "Failed to remove folder for $UserName`: $($_.Exception.Message)" -Level 'ERROR'
                    }
                }
            }
            
            'S' {  # Skip
                $actionTaken = 'Skipped by administrator'
                if (-not $Quiet) {
                    Write-Host "  Skipped corruption repair for $UserName" -ForegroundColor Yellow
                }
            }
        }
        
        return [PSCustomObject]@{
            Fixed = $fixed
            ActionTaken = $actionTaken
            Choice = $choice
        }
    }
    #endregion

    #region Core Profile Functions
    function Get-UserProfiles {
        param([string]$ComputerName)
        
        Write-DPLog -Message "Scanning profiles on $ComputerName..." -Level 'INFO'
        
        $profiles = @()
        
        try {
            $profileListPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
            
            if ($ComputerName -ne $env:COMPUTERNAME -and $ComputerName -ne 'localhost' -and $ComputerName -ne '.') {
                # Use WMI for remote registry
                $profileKeys = Get-WmiObject -ComputerName $ComputerName -Class StdRegProv -Namespace 'root\default' -ErrorAction Stop |
                    ForEach-Object { $_.EnumKey(2147483650, 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList') } |
                    Select-Object -ExpandProperty sNames |
                    Where-Object { $_ -match '^S-1-5-21' }
                
                foreach ($sid in $profileKeys) {
                    try {
                        $regProv = Get-WmiObject -ComputerName $ComputerName -Class StdRegProv -Namespace 'root\default'
                        $profilePathValue = $regProv.GetStringValue(2147483650, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid", 'ProfileImagePath')
                        $roamingValue = $regProv.GetDWORDValue(2147483650, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid", 'RoamingConfigured')
                        $tempValue = $regProv.GetDWORDValue(2147483650, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid", 'TemporaryProfile')
                        
                        $profilePath = $profilePathValue.sValue
                        if ($profilePath) {
                            $profilePath = $profilePath -replace '%SystemDrive%', 'C:'
                        }
                        
                        $profiles += @{
                            SID = $sid
                            ProfilePath = $profilePath
                            RoamingConfigured = $roamingValue.uValue
                            TemporaryProfile = $tempValue.uValue
                        }
                    }
                    catch {
                        Write-DPLog -Message "Error reading profile $sid on $ComputerName`: $($_.Exception.Message)" -Level 'WARNING'
                    }
                }
            }
            else {
                # Local registry access
                $profileKeys = Get-ChildItem $profileListPath -ErrorAction Stop | 
                    Where-Object { $_.PSChildName -match '^S-1-5-21' }
                
                foreach ($key in $profileKeys) {
                    try {
                        $props = Get-ItemProperty $key.PSPath
                        $profiles += @{
                            SID = $key.PSChildName
                            ProfilePath = $props.ProfileImagePath
                            RoamingConfigured = $props.RoamingConfigured
                            TemporaryProfile = $props.TemporaryProfile
                        }
                    }
                    catch {
                        Write-DPLog -Message "Error reading profile $($key.PSChildName)`:`r`n$($_.Exception.Message)" -Level 'WARNING'
                    }
                }
            }
        }
        catch {
            Write-DPLog -Message "Failed to enumerate profiles on $ComputerName`: $($_.Exception.Message)" -Level 'ERROR'
            return $null
        }
        
        return $profiles
    }

    function Test-ProfileFilter {
        param(
            [string]$UserName,
            [string]$SID,
            [string]$ProfilePath,
            [long]$ProfileSize,
            [string]$ProfileType
        )
        
        # Include filter
        if ($Include) {
            $includeMatch = $false
            foreach ($pattern in $Include) {
                if ($UserName -like $pattern) { $includeMatch = $true; break }
            }
            if (-not $includeMatch) { return $false }
        }
        
        # Exclude filter
        if ($Exclude) {
            foreach ($pattern in $Exclude) {
                if ($UserName -like $pattern) { return $false }
            }
        }
        
        # Size filters
        if ($MinProfileSizeMB -and $ProfileSize -lt ($MinProfileSizeMB * 1MB)) { return $false }
        if ($MaxProfileSizeMB -and $ProfileSize -gt ($MaxProfileSizeMB * 1MB)) { return $false }
        
        # Profile type filter
        if ($ProfileType -ne 'All') {
            if ($ProfileType -eq 'Local' -and $ProfileType -ne 'Local') { return $false }
            if ($ProfileType -eq 'Roaming' -and $ProfileType -ne 'Roaming') { return $false }
            if ($ProfileType -eq 'Temporary' -and $ProfileType -ne 'Temporary') { return $false }
            if ($ProfileType -eq 'Mandatory' -and $ProfileType -ne 'Mandatory') { return $false }
        }
        
        return $true
    }
    #endregion

    #region Profile Deletion Functions
    function Dismount-RegistryHive {
        param([string]$ProfilePath)
        
        try {
            # Find loaded hives
            $loadedHives = Get-ChildItem 'HKU:' -ErrorAction SilentlyContinue | 
                Where-Object { $_.Name -match '^HKEY_USERS\\S-1-5-21' }
            
            foreach ($hive in $loadedHives) {
                $sid = Split-Path $hive.Name -Leaf
                $hiveProfilePath = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid" -ErrorAction SilentlyContinue).ProfileImagePath
                
                if ($hiveProfilePath -eq $ProfilePath) {
                    Write-DPLog -Message "Unloading registry hive for SID $sid" -Level 'VERBOSE'
                    $result = reg.exe unload "HKU\$sid" 2>&1
                    if ($LASTEXITCODE -ne 0) {
                        Write-DPLog -Message "Failed to unload registry hive: $result" -Level 'WARNING'
                        return $false
                    }
                }
            }
            return $true
        }
        catch {
            return $false
        }
    }

    function Remove-ProfileWithRetry {
        param(
            [string]$ProfilePath,
            [string]$SID,
            [string]$ComputerName
        )
        
        $attempt = 0
        $success = $false
        
        while ($attempt -lt $MaxRetries -and -not $success) {
            $attempt++
            
            try {
                # Unload registry hive if requested
                if ($UnloadHives) {
                    Dismount-RegistryHive -ProfilePath $ProfilePath
                }
                
                # Remove read-only attributes
                Get-ChildItem -Path $ProfilePath -Recurse -Force -ErrorAction SilentlyContinue | 
                    ForEach-Object { $_.Attributes = $_.Attributes -band -bnot [System.IO.FileAttributes]::ReadOnly }
                
                # Remove the directory
                if ($ComputerName -eq $env:COMPUTERNAME -or $ComputerName -eq 'localhost' -or $ComputerName -eq '.') {
                    Remove-Item -Path $ProfilePath -Recurse -Force -ErrorAction Stop
                }
                else {
                    # Remote deletion using Invoke-Command
                    $scriptBlock = {
                        param($Path)
                        Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue | 
                            ForEach-Object { $_.Attributes = $_.Attributes -band -bnot [System.IO.FileAttributes]::ReadOnly }
                        Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
                    }
                    Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $ProfilePath -ErrorAction Stop
                }
                
                $success = $true
            }
            catch {
                if ($attempt -lt $MaxRetries) {
                    Write-DPLog -Message "Attempt $attempt failed for $ProfilePath, retrying in $RetryDelaySeconds seconds..." -Level 'WARNING'
                    Start-Sleep -Seconds $RetryDelaySeconds
                }
                else {
                    Write-DPLog -Message "Failed to delete $ProfilePath after $MaxRetries attempts`: $($_.Exception.Message)" -Level 'ERROR'
                }
            }
        }
        
        return $success
    }

    function Remove-UserProfile {
        [CmdletBinding(SupportsShouldProcess = $true)]
        param(
            [string]$ComputerName,
            [string]$SID,
            [string]$ProfilePath,
            [string]$UserName
        )
        
        Write-DPLog -Message "Deleting profile for '$UserName' on $ComputerName..." -Level 'INFO'
        
        $success = $true
        
        # Step 1: Remove from registry
        if ($PSCmdlet.ShouldProcess("Remove registry key for $SID", "Delete Profile Registry Entry")) {
            try {
                $regKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$SID"
                
                if ($ComputerName -eq $env:COMPUTERNAME -or $ComputerName -eq 'localhost' -or $ComputerName -eq '.') {
                    if (Test-Path $regKey) {
                        Remove-Item -Path $regKey -Recurse -Force
                    }
                }
                else {
                    $scriptBlock = {
                        param($sid)
                        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
                        if (Test-Path "$regPath\$sid") {
                            Remove-Item -Path "$regPath\$sid" -Recurse -Force
                        }
                    }
                    Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $SID
                }
                
                Write-DPLog -Message "Registry key removed for $UserName" -Level 'SUCCESS'
            }
            catch {
                Write-DPLog -Message "Failed to remove registry key for $UserName`: $($_.Exception.Message)" -Level 'ERROR'
                $success = $false
            }
        }
        
        # Step 2: Remove profile directory
        if ($success -and $PSCmdlet.ShouldProcess("Remove directory $ProfilePath", "Delete Profile Folder")) {
            if (Remove-ProfileWithRetry -ProfilePath $ProfilePath -SID $SID -ComputerName $ComputerName) {
                Write-DPLog -Message "Profile directory removed for $UserName" -Level 'SUCCESS'
            }
            else {
                Write-DPLog -Message "Failed to remove profile directory for $UserName" -Level 'ERROR'
                $success = $false
            }
        }
        
        return $success
    }
    #endregion

    #region Main Processing Functions
    function Invoke-ComputerProcessing {
        [CmdletBinding(SupportsShouldProcess = $true)]
        param([string]$ComputerName)
        
        Write-DPLog -Message "Processing computer: $ComputerName" -Level 'INFO'
        
        # Test connection
        $connection = Test-ComputerConnection -Computer $ComputerName
        if (-not $connection.Success) {
            Write-DPLog -Message "Cannot connect to $ComputerName`: $($connection.Error)" -Level 'ERROR'
            return
        }
        
        # Get active sessions
        $activeSessions = @()
        if (-not $IgnoreActiveSessions) {
            $activeSessions = Get-ActiveSessions -ComputerName $ComputerName
            Write-DPLog -Message "Active sessions on $ComputerName`: $($activeSessions -join ', ')" -Level 'VERBOSE'
        }
        
        # Get profiles
        $profiles = Get-UserProfiles -ComputerName $ComputerName
        if ($null -eq $profiles) {
            return
        }
        
        Write-DPLog -Message "Found $($profiles.Count) profiles on $ComputerName" -Level 'INFO'
        
        # Process each profile
        foreach ($profileInfo in $profiles) {
            $sid = $profileInfo.SID
            $profilePath = $profileInfo.ProfilePath
            
            if (-not $profilePath) {
                Write-DPLog -Message "Profile $sid has no path, skipping" -Level 'WARNING'
                continue
            }
            
            # Resolve username
            $userName = ConvertTo-UserName -SID $sid
            if (-not $userName) {
                $userName = "Unknown ($sid)"
            }
            
            # Check if protected
            if (-not $IncludeSystemProfiles) {
                $protectCheck = Test-IsProtectedProfile -UserName $userName -SID $sid
                if ($protectCheck) {
                    Write-DPLog -Message "Skipping protected profile: $userName" -Level 'VERBOSE'
                    continue
                }
            }
            
            # Get profile type
            $profType = Get-ProfileType -ProfileInfo $profileInfo
            
            # Skip corrupted unless requested or FixCorruption is enabled
            if ($profType -like 'Corrupted*' -and -not $IncludeCorrupted -and -not $FixCorruption) {
                Write-DPLog -Message "Skipping corrupted profile: $userName" -Level 'VERBOSE'
                continue
            }
            
            # Handle corruption repair if requested
            if ($profType -like 'Corrupted*' -and $FixCorruption) {
                $repairResult = Repair-CorruptedProfile -ComputerName $ComputerName -SID $sid -UserName $userName -ProfilePath $profilePath -CorruptionType $profType
                
                # Add repair result to the profile info for reporting
                if (-not $script:RepairResults) {
                    $script:RepairResults = [System.Collections.Generic.List[object]]::new()
                }
                $script:RepairResults.Add([PSCustomObject]@{
                    ComputerName = $ComputerName
                    UserName = $userName
                    SID = $sid
                    CorruptionType = $profType
                    Fixed = $repairResult.Fixed
                    ActionTaken = $repairResult.ActionTaken
                })
                
                # If corruption was fixed by recreating NTUSER.DAT, continue processing normally
                if ($repairResult.Fixed -and $repairResult.Choice -eq 'R') {
                    # Re-check profile type - it should now be valid
                    $profType = Get-ProfileType -ProfileInfo $profileInfo
                    Write-DPLog -Message "Profile $userName repaired successfully - continuing with normal processing" -Level 'SUCCESS'
                }
                # If profile was deleted as part of repair, skip further processing
                elseif ($repairResult.Choice -in @('D', 'F')) {
                    Write-DPLog -Message "Profile $userName handled via corruption repair - skipping deletion phase" -Level 'INFO'
                    continue
                }
                # If skipped or failed, move to next profile
                elseif ($repairResult.Choice -eq 'S' -or -not $repairResult.Fixed) {
                    Write-DPLog -Message "Corruption repair skipped or failed for $userName - continuing to next profile" -Level 'WARNING'
                    continue
                }
            }
            
            # Get profile age
            $ageInfo = Get-ProfileAge -ProfilePath $profilePath -SID $sid -Method $AgeCalculation -ComputerName $ComputerName
            $lastUsed = $ageInfo.LastUsed
            $ageSource = $ageInfo.Source
            
            # Calculate age in days
            $ageInDays = if ($lastUsed -eq [DateTime]::MinValue) { -1 } else { [math]::Floor((Get-Date) - $lastUsed).TotalDays }
            
            # Check age criteria
            if ($ageInDays -lt $DaysInactive -and $ageInDays -ge 0) {
                Write-DPLog -Message "Profile $userName is too recent ($ageInDays days, source: $ageSource)" -Level 'VERBOSE'
                continue
            }
            
            # Get profile size
            $sizeBytes = Get-ProfileFolderSize -Path $profilePath
            $sizeFormatted = Format-Bytes -Bytes $sizeBytes
            
            # Apply filters
            $passesFilter = Test-ProfileFilter -UserName $userName -SID $sid -ProfilePath $profilePath -ProfileSize $sizeBytes -ProfileType $profType
            if (-not $passesFilter) {
                Write-DPLog -Message "Profile $userName filtered out" -Level 'VERBOSE'
                continue
            }
            
            # Check for active session
            $hasActiveSession = $false
            foreach ($session in $activeSessions) {
                if ($userName -like "*$session*" -or $session -like "*$($userName.Split('\')[-1])") {
                    $hasActiveSession = $true
                    break
                }
            }
            
            if ($hasActiveSession -and -not $IgnoreActiveSessions) {
                Write-DPLog -Message "Skipping $userName - active session detected" -Level 'WARNING'
                continue
            }
            
            # Build result object
            $result = [PSCustomObject]@{
                ComputerName    = $ComputerName
                UserName        = $userName.Split('\')[-1]
                Domain          = if ($userName -contains '\') { $userName.Split('\')[0] } else { $env:USERDOMAIN }
                SID             = $sid
                ProfilePath     = $profilePath
                ProfileType     = $profType
                LastUsed        = if ($lastUsed -eq [DateTime]::MinValue) { 'Unknown' } else { $lastUsed.ToString('yyyy-MM-dd HH:mm:ss') }
                AgeInDays       = $ageInDays
                AgeSource       = $ageSource
                SizeBytes       = $sizeBytes
                SizeFormatted   = $sizeFormatted
                IsActiveSession = $hasActiveSession
                EligibleForDeletion = $true
                Deleted         = $false
                Error           = $null
            }
            
            # Display info
            if (-not $Quiet) {
                $color = if ($hasActiveSession) { 'Yellow' } else { (Get-AgeColor -AgeInDays $ageInDays) }
                $sizeStr = if ($ShowSpace) { " [Size: $sizeFormatted]" } else { '' }
                $activeStr = if ($hasActiveSession) { ' [ACTIVE]' } else { '' }
                Write-Host "  $userName - $ageInDays days ($ageSource)$sizeStr$activeStr" -ForegroundColor $color
                
                # Show detailed folder breakdown if requested
                if ($Detailed) {
                    $breakdown = Get-ProfileSizeBreakdown -ProfilePath $profilePath
                    $breakdownStr = $breakdown.GetEnumerator() | Where-Object { $_.Value -ne 'N/A' } | 
                        ForEach-Object { "$($_.Key): $($_.Value)" } | Join-String -Separator ', '
                    if ($breakdownStr) {
                        Write-Host "    Folders: $breakdownStr" -ForegroundColor Gray
                    }
                }
            }
            
            # Perform deletion if requested
            if ($Delete -and -not $hasActiveSession) {
                $targetDesc = "Delete profile for '$userName' on '$ComputerName' ($ageInDays days old, $sizeFormatted)"
                if ($PSCmdlet.ShouldProcess($targetDesc, 'Delete User Profile')) {
                    # Backup profile before deletion if requested
                    $backupSuccess = $true
                    if ($BackupPath) {
                        $backupSuccess = Backup-Profile -SourcePath $profilePath -UserName $userName
                    }
                    
                    if ($backupSuccess) {
                        if (Remove-UserProfile -ComputerName $ComputerName -SID $sid -ProfilePath $profilePath -UserName $userName) {
                            $result.Deleted = $true
                            $script:TotalProfilesDeleted++
                            $script:TotalSpaceFreed += $sizeBytes
                            Write-EventLogEntry -Message "Deleted profile: $userName on $ComputerName ($sizeFormatted)" -EntryType Information -EventId 1010
                        }
                        else {
                            $result.Error = 'Deletion failed'
                            Write-EventLogEntry -Message "Failed to delete profile: $userName on $ComputerName" -EntryType Error -EventId 1011
                        }
                    }
                    else {
                        $result.Error = 'Backup failed - deletion cancelled'
                        Write-DPLog -Message "Deletion cancelled for $userName - backup failed" -Level 'ERROR'
                    }
                }
            }
            
            $script:Results.Add($result)
            $script:TotalProfilesProcessed++
        }
        
        Write-DPLog -Message "Finished processing $ComputerName" -Level 'INFO'
    }

    function Show-Summary {
        if (-not $Quiet) {
            $duration = (Get-Date) - $script:StartTime
            
            Write-Host "`n" + ('=' * 80) -ForegroundColor Cyan
            Write-Host " SUMMARY" -ForegroundColor Cyan
            Write-Host ('=' * 80) -ForegroundColor Cyan
            Write-Host " Computers processed: $($ComputerName.Count)"
            Write-Host " Profiles processed:  $script:TotalProfilesProcessed"
            if ($Delete) {
                Write-Host " Profiles deleted:    $script:TotalProfilesDeleted" -ForegroundColor $(if ($script:TotalProfilesDeleted -gt 0) { 'Green' } else { 'White' })
                Write-Host " Space freed:         $(Format-Bytes -Bytes $script:TotalSpaceFreed)" -ForegroundColor $(if ($script:TotalSpaceFreed -gt 0) { 'Green' } else { 'White' })
            }
            else {
                # Dry run preview - show what WOULD be deleted
                $wouldDelete = $script:Results | Where-Object { $_.EligibleForDeletion -and -not $_.IsActiveSession }
                $wouldDeleteCount = $wouldDelete.Count
                $wouldDeleteSize = ($wouldDelete | Measure-Object -Property SizeBytes -Sum).Sum
                Write-Host " Would delete:        $wouldDeleteCount profiles ($(Format-Bytes -Bytes $wouldDeleteSize))" -ForegroundColor Yellow
            }
            Write-Host " Duration:            $($duration.ToString('hh\:mm\:ss'))"
            
            # Top 5 largest profiles
            if ($script:Results.Count -gt 0) {
                Write-Host "`n TOP 5 LARGEST PROFILES:" -ForegroundColor Cyan
                $top5 = $script:Results | Sort-Object SizeBytes -Descending | Select-Object -First 5
                foreach ($prof in $top5) {
                    Write-Host "  $($prof.UserName) on $($prof.ComputerName): $($prof.SizeFormatted) ($($prof.AgeInDays) days)" -ForegroundColor Gray
                }
            }
            
            # Age breakdown analysis
            if ($script:Results.Count -gt 0) {
                Write-Host "`n AGE BREAKDOWN:" -ForegroundColor Cyan
                $ageGroups = $script:Results | Group-Object -Property { 
                    if ($_.AgeInDays -lt 30) { '0-30 days' }
                    elseif ($_.AgeInDays -lt 60) { '31-60 days' }
                    elseif ($_.AgeInDays -lt 90) { '61-90 days' }
                    elseif ($_.AgeInDays -lt 180) { '91-180 days' }
                    elseif ($_.AgeInDays -ge 180) { '180+ days' }
                    else { 'Unknown' }
                } | Sort-Object Name
                
                foreach ($group in $ageGroups) {
                    $groupSize = ($group.Group | Measure-Object -Property SizeBytes -Sum).Sum
                    Write-Host "  $($group.Name): $($group.Count) profiles ($(Format-Bytes -Bytes $groupSize))"
                }
            }
            
            Write-Host ('=' * 80) -ForegroundColor Cyan
        }
        
        Write-DPLog -Message "Script completed. Processed: $script:TotalProfilesProcessed, Deleted: $script:TotalProfilesDeleted" -Level 'INFO'
        
        # Export to CSV if requested
        if ($OutputPath -and $script:Results.Count -gt 0) {
            try {
                $script:Results | Export-Csv -Path $OutputPath -NoTypeInformation -Force
                Write-DPLog -Message "Results exported to $OutputPath" -Level 'SUCCESS'
            }
            catch {
                Write-DPLog -Message "Failed to export to CSV`: $($_.Exception.Message)" -Level 'ERROR'
            }
        }
    }
    #endregion

    #region Script Entry Point
    Write-DPHeader
    
    # Log to event log
    Write-EventLogEntry -Message "Delprof2-PS started. Version: $script:Version, Delete mode: $Delete, Days inactive: $DaysInactive" -EntryType Information -EventId 1000
    
    # Test mode - validate prerequisites without processing
    if ($Test) {
        Write-Host "`n=== TEST MODE - Validating Prerequisites ===" -ForegroundColor Cyan
        foreach ($computer in $ComputerName) {
            Write-Host "Testing $computer..." -NoNewline
            $result = Test-ComputerConnection -Computer $computer.Trim()
            if ($result.Success) {
                Write-Host " OK" -ForegroundColor Green
                try {
                    $profiles = Get-UserProfiles -ComputerName $computer.Trim()
                    Write-Host "  Found $($profiles.Count) profiles" -ForegroundColor Gray
                }
                catch {
                    Write-Host "  ERROR: Could not enumerate profiles" -ForegroundColor Red
                }
            }
            else {
                Write-Host " FAILED - $($result.Error)" -ForegroundColor Red
            }
        }
        Write-Host "`nTest mode complete. No changes were made." -ForegroundColor Cyan
        exit 0
    }
    
    # Preview mode banner
    if ($Preview) {
        Write-Host "`n╔════════════════════════════════════════════════════════════════════════╗" -ForegroundColor Magenta
        Write-Host "║                        PREVIEW/SIMULATION MODE                         ║" -ForegroundColor Magenta
        Write-Host "║          No profiles will be deleted - showing what WOULD happen       ║" -ForegroundColor Magenta
        Write-Host "╚════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Magenta
    }
    
    # Validate admin rights for local execution
    if ($ComputerName -contains $env:COMPUTERNAME -or $ComputerName -contains 'localhost' -or $ComputerName -contains '.') {
        if (-not (Test-AdminRights)) {
            Write-DPLog -Message "Administrator privileges required for local execution. Restart as admin." -Level 'ERROR'
            Write-EventLogEntry -Message "Failed to start - admin rights required" -EntryType Error -EventId 1005
            exit 1
        }
    }
    
    # Validate delete warnings
    if ($Delete -and -not $Force -and -not $Quiet -and -not $Interactive) {
        Write-Host "`nWARNING: Delete mode is enabled. Profiles will be permanently removed!" -ForegroundColor Red -BackgroundColor Black
        if ($IgnoreActiveSessions) {
            Write-Host "WARNING: Active sessions will be deleted - this can cause data loss!" -ForegroundColor Red -BackgroundColor Black
        }
        Write-Host "Press Ctrl+C to cancel, or " -NoNewline -ForegroundColor Yellow
        Write-Host "Enter to continue..." -NoNewline -ForegroundColor Yellow
        $null = Read-Host
    }
    
    # Build computer queue
    foreach ($computer in $ComputerName) {
        $script:ComputerQueue.Add($computer.Trim())
    }
    
    # Pre-validate all computers if using interactive mode
    $validComputers = [System.Collections.Generic.List[string]]::new()
    if ($Interactive) {
        Write-Host "`nValidating computers for interactive mode..." -ForegroundColor Cyan
        foreach ($computer in $ComputerName) {
            $result = Test-ComputerConnection -Computer $computer.Trim()
            if ($result.Success) {
                $validComputers.Add($computer.Trim())
            }
            else {
                Write-DPLog -Message "Skipping $computer in interactive mode - $($result.Error)" -Level 'WARNING'
            }
        }
    }
}

process {
    # Mass deletion safeguard - warn if processing many profiles
    if ($Delete -and -not $Force -and -not $Quiet -and -not $Interactive) {
        # First pass - count eligible profiles without deleting
        $estimatedDeletions = 0
        foreach ($computer in $ComputerName) {
            Invoke-ComputerProcessing -ComputerName $computer.Trim()
            $eligibleOnComputer = $script:Results | Where-Object { 
                $_.ComputerName -eq $computer.Trim() -and $_.EligibleForDeletion -and -not $_.IsActiveSession 
            }
            $estimatedDeletions += $eligibleOnComputer.Count
        }
        
        # Mass deletion warning threshold
        if ($estimatedDeletions -gt 50) {
            Write-Host "`n  MASS DELETION WARNING " -ForegroundColor Red -BackgroundColor Black
            Write-Host "This operation will delete $estimatedDeletions profiles!" -ForegroundColor Red
            Write-Host "This is an unusually large number of deletions." -ForegroundColor Yellow
            Write-Host "Are you sure you want to proceed?" -ForegroundColor Yellow
            Write-Host "Type 'YES' to confirm: " -NoNewline -ForegroundColor Yellow
            $confirm = Read-Host
            if ($confirm -ne 'YES') {
                Write-Host "Operation cancelled by user." -ForegroundColor Red
                exit 1
            }
        }
        
        # Reset counters for actual deletion pass
        $script:Results.Clear()
        $script:TotalProfilesProcessed = 0
        $script:TotalProfilesDeleted = 0
        $script:TotalSpaceFreed = 0
    }
    
    # Interactive mode processing
    if ($Interactive) {
        foreach ($computer in $validComputers) {
            Invoke-ComputerProcessing -ComputerName $computer
        }
        
        # After collecting all eligible profiles, let user select interactively
        $eligibleProfiles = $script:Results | Where-Object { $_.EligibleForDeletion -and -not $_.IsActiveSession }
        
        if ($eligibleProfiles.Count -eq 0) {
            Write-Host "`nNo eligible profiles found for interactive selection." -ForegroundColor Yellow
        }
        else {
            $selectedProfiles = Select-ProfilesInteractive -Profiles $eligibleProfiles
            
            if ($selectedProfiles.Count -gt 0) {
                Write-Host "`nDeleting $($selectedProfiles.Count) selected profiles..." -ForegroundColor Yellow
                foreach ($prof in $selectedProfiles) {
                    if (Remove-UserProfile -ComputerName $prof.ComputerName -SID $prof.SID -ProfilePath $prof.ProfilePath -UserName $prof.UserName) {
                        # Update the result
                        $resultItem = $script:Results | Where-Object { $_.SID -eq $prof.SID -and $_.ComputerName -eq $prof.ComputerName }
                        if ($resultItem) {
                            $resultItem.Deleted = $true
                            $script:TotalProfilesDeleted++
                            $script:TotalSpaceFreed += $resultItem.SizeBytes
                        }
                        Write-EventLogEntry -Message "Deleted profile: $($prof.UserName) on $($prof.ComputerName)" -EntryType Information -EventId 1010
                    }
                }
            }
            else {
                Write-Host "`nNo profiles selected for deletion." -ForegroundColor Green
            }
        }
    }
    elseif ($UseParallel -and $ComputerName.Count -gt 1) {
        # Parallel processing for multiple computers
        $ComputerName | ForEach-Object -Parallel {
            # Note: In parallel mode, we need to handle logging carefully
            $comp = $_
            # Simplified parallel processing - log to file only
            Invoke-ComputerProcessing -ComputerName $comp.Trim()
        } -ThrottleLimit $ThrottleLimit
    }
    else {
        # Standard sequential processing with progress bar
        $computerCount = $ComputerName.Count
        for ($i = 0; $i -lt $computerCount; $i++) {
            $percentComplete = [math]::Floor(($i / $computerCount) * 100)
            Write-Progress -Activity "Processing Computers" -Status "Processing $($ComputerName[$i]) ($($i+1) of $computerCount)" -PercentComplete $percentComplete
            Invoke-ComputerProcessing -ComputerName $ComputerName[$i].Trim()
        }
        Write-Progress -Activity "Processing Computers" -Completed
    }
}

end {
    Show-Summary
    
    # Preview mode completion message
    if ($Preview) {
        Write-Host "`n╔════════════════════════════════════════════════════════════════════════╗" -ForegroundColor Magenta
        Write-Host "║                    PREVIEW MODE COMPLETE                               ║" -ForegroundColor Magenta
        Write-Host "║        No profiles were deleted. Use -Delete to perform deletion.      ║" -ForegroundColor Magenta
        Write-Host "╚════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Magenta
    }
    
    # Generate HTML report if requested
    if ($HtmlReport -and $script:Results.Count -gt 0) {
        $summary = @{
            Computers = $ComputerName.Count
            ProfilesProcessed = $script:TotalProfilesProcessed
            ProfilesDeleted = $script:TotalProfilesDeleted
            SpaceFreed = Format-Bytes -Bytes $script:TotalSpaceFreed
            Duration = ((Get-Date) - $script:StartTime).ToString('hh\:mm\:ss')
        }
        Export-HtmlReport -Path $HtmlReport -Results $script:Results -Summary $summary
    }
    
    # Send email notification if configured
    if ($SmtpServer -and $EmailTo) {
        $summary = @{
            Computers = $ComputerName.Count
            ProfilesProcessed = $script:TotalProfilesProcessed
            ProfilesDeleted = $script:TotalProfilesDeleted
            SpaceFreed = Format-Bytes -Bytes $script:TotalSpaceFreed
            Duration = ((Get-Date) - $script:StartTime).ToString('hh\:mm\:ss')
        }
        Send-NotificationEmail -Summary $summary
    }
    
    # Log completion to event log
    Write-EventLogEntry -Message "Delprof2-PS completed. Processed: $script:TotalProfilesProcessed, Deleted: $script:TotalProfilesDeleted, Space freed: $(Format-Bytes -Bytes $script:TotalSpaceFreed)" -EntryType Information -EventId 1002
    
    # Export to CSV if requested
    if ($OutputPath -and $script:Results.Count -gt 0) {
        try {
            $script:Results | Export-Csv -Path $OutputPath -NoTypeInformation -Force
            Write-DPLog -Message "Results exported to $OutputPath" -Level 'SUCCESS'
        }
        catch {
            Write-DPLog -Message "Failed to export to CSV`: $($_.Exception.Message)" -Level 'ERROR'
        }
    }
    
    # Return results for pipeline
    if ($script:Results.Count -gt 0) {
        return $script:Results
    }
}
#endregion