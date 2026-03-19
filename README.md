# Delprof2-PS

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://docs.microsoft.com/powershell/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Version](https://img.shields.io/badge/version-2.0.0-brightgreen.svg)](VERSION)

> **A comprehensive, enterprise-grade PowerShell replacement for Helge Klein's Delprof2 with advanced features, safety mechanisms, and modern capabilities.**

---

## 📋 Overview

Delprof2-PS is a feature-rich PowerShell script that **exceeds** the capabilities of the original [Delprof2](https://helgeklein.com/delprof2/) tool. It provides enterprise-grade user profile management with comprehensive safety features, detailed reporting, and modern PowerShell capabilities.

### ✨ Why Choose Delprof2-PS?

- 🚀 **Modern PowerShell** - Built for PowerShell 5.1+ with full cmdlet support
- 🛡️ **Safety First** - Multiple layers of protection prevent accidental deletions
- 📊 **Rich Reporting** - HTML reports, CSV export, and Windows Event Log integration
- 🔧 **Corruption Repair** - Interactive mode to fix corrupted profiles safely
- 🖥️ **Visual Interface** - Interactive selection with color-coded output
- ⚡ **Parallel Processing** - Process hundreds of computers efficiently

---

## 🎯 Features

### Core Capabilities

- **Local and Remote Computer Support** - Process profiles on multiple computers via pipeline or CSV input
- **Multiple Age Calculation Methods** - NTUSER.DAT, ProfilePath, Registry, LastLogon, LastLogoff
- **Active Session Detection** - Multi-method detection (quser, WMI, explorer.exe processes)
- **Flexible Filtering** - Include/exclude by username patterns, SID patterns, profile size
- **Profile State Detection** - Local, Roaming, Temporary, Mandatory, Corrupted
- **Disk Space Reporting** - Calculate and display profile sizes
- **Registry Hive Unloading** - Safely unload loaded hives before deletion
- **Retry Logic** - Configurable retries for locked files

### Enterprise Features

| Feature | Description |
| --------- | ------------- |
| `-Interactive` | Visual menu for manual profile selection with keyboard navigation |
| `-Test` | Validate prerequisites and connectivity without making changes |
| `-Preview` | Simulation mode with visual banners showing what WOULD be deleted |
| `-UseParallel` | Parallel processing for multiple computers with throttling |
| `-HtmlReport` | Generate professional HTML reports with CSS styling |
| Event Log Integration | Windows Event Log entries for auditing (Event IDs 1000-1012) |
| `-BackupPath` | ZIP compression backup before deletion |
| `-ConfigFile` | JSON configuration file support |
| Email Notifications | SMTP alerts for scheduled tasks |
| Progress Bars | Visual progress indicators for long operations |
| Age Color Coding | Color-coded output based on profile age |
| `-Detailed` | Per-folder size breakdown (Documents, Downloads, etc.) |
| Mass Deletion Safeguard | Warns on >50 profile deletions (requires "YES" confirmation) |
| `-FixCorruption` | Interactive corruption repair with admin approval |

## Requirements

- PowerShell 5.1 or later
- Administrative privileges on target computers
- Remote management enabled for remote computer processing (WinRM/PSRemoting)

## Installation

1. Download the script:

```powershell
# Clone the repository
git clone https://github.com/ArMaTeC/DelprofPS.git

# Or download directly
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/ArMaTeC/DelprofPS/main/DelprofPS.ps1" -OutFile "DelprofPS.ps1"
```

1. (Optional) Create a configuration file - see `DelprofPS.config.json` example

## Quick Start

### Basic Usage

```powershell
# List all profiles older than 30 days (dry run)
.\DelprofPS.ps1

# Preview what would be deleted
.\DelprofPS.ps1 -DaysInactive 60 -Preview

# Delete profiles older than 60 days
.\DelprofPS.ps1 -DaysInactive 60 -Delete

# Interactive selection mode
.\DelprofPS.ps1 -DaysInactive 90 -Interactive
```

### Advanced Usage

```powershell
# Enterprise deployment with full reporting
.\DelprofPS.ps1 -ComputerName (Import-Csv servers.csv).Name `
    -Delete -DaysInactive 90 `
    -LogPath "C:\Logs\delprof.log" `
    -OutputPath "C:\Logs\results.csv" `
    -HtmlReport "C:\Logs\report.html" `
    -BackupPath "C:\Backups\Profiles" `
    -UnloadHives -ShowSpace -Detailed

# Parallel processing for many computers
.\DelprofPS.ps1 -ComputerName (Get-Content servers.txt) `
    -UseParallel -ThrottleLimit 10 `
    -DaysInactive 120 -Delete

# Scheduled task with email notification
.\DelprofPS.ps1 -DaysInactive 60 -Delete -Exclude "*admin*" `
    -SmtpServer "mail.company.com" `
    -EmailTo "admin@company.com" `
    -EmailFrom "delprofps@server01" `
    -Quiet -LogPath "C:\Logs\delprof.log"
```

## Parameters

| Parameter | Description | Default |
| ----------- | ------------- | --------- |
| `ComputerName` | Target computer(s) | `$env:COMPUTERNAME` |
| `DaysInactive` | Minimum days of inactivity | 30 |
| `AgeCalculation` | Method: NTUSER_DAT, ProfilePath, Registry, LastLogon, LastLogoff | NTUSER_DAT |
| `Include` | Wildcard patterns to include | None |
| `Exclude` | Wildcard patterns to exclude | None |
| `Delete` | Actually delete profiles | False (dry-run) |
| `Force` | Skip confirmations | False |
| `IgnoreActiveSessions` | Allow deleting active sessions | False |
| `UnloadHives` | Unload registry hives | False |
| `MaxRetries` | Retry attempts for locked files | 3 |
| `RetryDelaySeconds` | Seconds between retries | 2 |
| `OutputPath` | CSV export path | None |
| `LogPath` | Log file path | None |
| `Quiet` | Suppress console output | False |
| `ShowSpace` | Display disk space | False |
| `IncludeSystemProfiles` | Include system profiles | False |
| `IncludeSpecialProfiles` | Include special accounts | False |
| `MinProfileSizeMB` | Minimum profile size | None |
| `MaxProfileSizeMB` | Maximum profile size | None |
| `IncludeCorrupted` | Include corrupted profiles | False |
| `FixCorruption` | Enable interactive corruption repair | False |
| `ProfileType` | Filter: Local, Roaming, Temporary, Mandatory, All | All |
| `Interactive` | Interactive selection mode | False |
| `Test` | Test connectivity only | False |
| `HtmlReport` | HTML report path | None |
| `BackupPath` | Backup directory path | None |
| `ConfigFile` | JSON config file path | None |
| `UseParallel` | Use parallel processing | False |
| `ThrottleLimit` | Max parallel threads | 5 |
| `SmtpServer` | SMTP server for email | None |
| `EmailTo` | Email recipient | None |
| `EmailFrom` | Email sender | delprofps@computername |
| `Detailed` | Show folder breakdown | False |
| `Preview` | Simulation mode | False |
| `VerifyIntegrity` | Verify script SHA256 hash before execution | False |
| `Credential` | PSCredential for remote computer authentication | Current identity |

## Configuration File

Create a `DelprofPS.config.json` file to store default settings:

```json
{
    "DaysInactive": 60,
    "Exclude": ["*admin*", "*service*"],
    "Include": [],
    "LogPath": "C:\\Logs\\DelprofPS.log",
    "OutputPath": "C:\\Logs\\DelprofPS_Results.csv",
    "HtmlReport": "C:\\Logs\\DelprofPS_Report.html",
    "BackupPath": "C:\\Backups\\Profiles",
    "SmtpServer": "mail.company.com",
    "EmailTo": "admin@company.com"
}
```

## Safety Features

Delprof2-PS includes multiple layers of protection:

1. **Default Dry-Run Mode** - Must use `-Delete` to actually remove profiles
2. **Active Session Protection** - Skips logged-in users (unless `-IgnoreActiveSessions`)
3. **System Profile Protection** - Excludes Default, Public, SYSTEM, etc.
4. **Registry Hive Unloading** - `-UnloadHives` safely unloads hives before deletion
5. **Backup Capability** - `-BackupPath` creates ZIP backups before deletion
6. **Mass Deletion Warning** - Warns on >50 profiles, requires "YES" confirmation
7. **Corruption Repair Safety** - `-FixCorruption` requires interactive admin approval for each fix
8. **Comprehensive Logging** - File and Windows Event Log integration
9. **ShouldProcess Support** - `-WhatIf` and `-Confirm` standard PowerShell parameters

## Event Log IDs

| Event ID | Description |
| ---------- | ------------- |
| 1000 | Script started |
| 1001 | HTML report generated |
| 1002 | Script completed |
| 1005 | Error - admin rights required |
| 1010 | Profile deleted |
| 1011 | Profile deletion failed |
| 1012 | Corruption repair action taken |

## Examples

### Example 1: Basic Dry Run

```powershell
# See what profiles exist and their ages
.\DelprofPS.ps1 -DaysInactive 30
```

### Example 2: Preview Mode

```powershell
# See what WOULD be deleted without actually deleting
.\DelprofPS.ps1 -DaysInactive 60 -Preview -ShowSpace
```

### Example 3: Interactive Selection

```powershell
# Use arrow keys to navigate, Space to toggle, Enter to confirm
.\DelprofPS.ps1 -DaysInactive 90 -Interactive
```

### Example 4: Delete with Backup

```powershell
# Backup profiles before deletion
.\DelprofPS.ps1 -Delete -DaysInactive 60 -BackupPath "D:\Backups"
```

### Example 5: Remote Computers

```powershell
# Process multiple computers
$computers = @("SERVER01", "SERVER02", "SERVER03")
.\DelprofPS.ps1 -ComputerName $computers -Delete -DaysInactive 90
```

### Example 6: Detailed Analysis

```powershell
# Show folder breakdowns and export to HTML
.\DelprofPS.ps1 -Detailed -ShowSpace -HtmlReport "C:\Reports\profiles.html"
```

### Example 7: Test Connectivity

```powershell
# Validate access to computers without processing
.\DelprofPS.ps1 -ComputerName (Get-Content servers.txt) -Test
```

### Example 8: Email Notification

```powershell
# Send email report after completion
.\DelprofPS.ps1 -Delete -DaysInactive 60 `
    -SmtpServer "smtp.office365.com" `
    -EmailTo "it-team@company.com" `
    -EmailFrom "delprofps@mgmt-server"
```

### Example 9: Use Configuration File

```powershell
# Load settings from JSON file
.\DelprofPS.ps1 -ConfigFile "C:\Config\delprof.json" -Delete
```

### Example 10: Parallel Processing

```powershell
# Process many computers in parallel
.\DelprofPS.ps1 -ComputerName (Get-Content 100-servers.txt) `
    -UseParallel -ThrottleLimit 20 `
    -Delete -DaysInactive 120
```

### Example 11: Corruption Repair Mode

```powershell
# Interactively fix corrupted profiles with full administrator control
# Options: Remove orphaned registry keys, Delete profile, Recreate NTUSER.DAT, Skip
.\DelprofPS.ps1 -FixCorruption -Interactive -IncludeCorrupted

# List corrupted profiles without making changes
.\DelprofPS.ps1 -IncludeCorrupted -Preview
```

---

## Output

### Console Output

The script provides color-coded console output:

- **Green** - Profiles < 30 days old
- **Yellow** - Profiles 30-90 days old  
- **Magenta** - Profiles 90-180 days old
- **Red** - Profiles > 180 days old
- **Yellow background** - Active sessions (protected)

### Summary Report

```text
================================================================================
 SUMMARY
================================================================================
 Computers processed: 5
 Profiles processed:  127
 Profiles deleted:    45
 Space freed:         12.5 GB
 Duration:            00:03:42

 TOP 5 LARGEST PROFILES:
  jdoe on SERVER01: 3.2 GB (185 days)
  smith on SERVER02: 2.8 GB (210 days)
  ...

 AGE BREAKDOWN:
  0-30 days: 12 profiles (2.1 GB)
  31-60 days: 25 profiles (4.5 GB)
  61-90 days: 30 profiles (3.2 GB)
  91-180 days: 35 profiles (5.1 GB)
  180+ days: 25 profiles (8.6 GB)
================================================================================
```

### HTML Report

Professional HTML reports include:

- Summary dashboard with statistics
- Detailed profile table with color-coded rows
- Profile type badges (Local, Roaming, Temporary, Mandatory)
- Status indicators (Deleted, Active, Error, Kept)
- Responsive design for mobile viewing

## Troubleshooting

### Common Issues

#### "Access Denied" errors

- Ensure running with administrative privileges
- Verify WinRM/PSRemoting is enabled on remote computers
- Check firewall rules for remote management

#### "Cannot enumerate profiles"

- Verify registry access permissions
- Check if target computer is online
- Test with `-Test` parameter first

#### Large deletions fail

- Use `-Force` to bypass confirmation (not recommended for production)
- Increase `-ThrottleLimit` for parallel processing
- Check available disk space for backups

### Debug Mode

```powershell
# Enable verbose output
.\DelprofPS.ps1 -Verbose -LogPath "C:\Logs\debug.log"
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Original [Delprof2](https://helgeklein.com/delprof2/) by Helge Klein
- PowerShell community for best practices and feedback

## Support

For issues, questions, or feature requests, please [open an issue](https://github.com/ArMaTeC/DelprofPS/issues).

---

**Disclaimer**: This tool modifies user data. Always test in non-production environments first. The authors are not responsible for data loss or system damage resulting from the use of this script.
