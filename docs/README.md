# DelprofPS

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://docs.microsoft.com/powershell/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Version](https://img.shields.io/badge/version-2.0.0-brightgreen.svg)](CHANGELOG.md)
[![CI](https://github.com/ArMaTeC/DelprofPS/actions/workflows/ci.yml/badge.svg)](https://github.com/ArMaTeC/DelprofPS/actions/workflows/ci.yml)

> **An enterprise-grade PowerShell replacement for Helge Klein's Delprof2 with a WPF GUI, advanced safety mechanisms, parallel processing, and comprehensive reporting.**

---

## Overview

DelprofPS is a feature-rich PowerShell tool that **exceeds** the capabilities of the original [Delprof2](https://helgeklein.com/delprof2/). It provides enterprise-grade user profile management with a modern graphical interface, comprehensive safety features, detailed reporting, and full PowerShell module support.

### Why Choose DelprofPS?

- **Modern WPF GUI** - Full graphical interface with profile browser, filters, and one-click actions
- **Safety First** - Multiple layers of protection prevent accidental deletions
- **Rich Reporting** - HTML reports, CSV export, and Windows Event Log integration
- **Corruption Repair** - Interactive mode to detect and fix corrupted profiles safely
- **Parallel Processing** - Process hundreds of computers efficiently with throttle control
- **Security Hardened** - Config validation, input sanitisation, log integrity hashing, and script self-verification
- **CI/CD Ready** - GitHub Actions pipeline with PSScriptAnalyzer linting and Pester tests

---

## Features

### Core Capabilities

- **Local and Remote Computer Support** - Process profiles on multiple computers via pipeline, arrays, or CSV input
- **Multiple Age Calculation Methods** - NTUSER.DAT, ProfilePath, Registry, LastLogon, LastLogoff
- **Active Session Detection** - Multi-method detection (quser, WMI, explorer.exe processes)
- **Flexible Filtering** - Include/exclude by username patterns, SID patterns, profile size thresholds
- **Profile State Detection** - Local, Roaming, Temporary, Mandatory, Corrupted
- **Disk Space Reporting** - Calculate and display profile sizes with human-readable formatting
- **Registry Hive Unloading** - Safely unload loaded hives before deletion
- **Retry Logic** - Configurable retries with delays for locked files

### Enterprise Features

| Feature | Description |
| --- | --- |
| `-UI` | **WPF graphical interface** with tabbed layout, profile DataGrid, and visual controls |
| `-Interactive` | Visual CLI menu for manual profile selection with keyboard navigation |
| `-Test` | Validate prerequisites and connectivity without making changes |
| `-Preview` | Simulation mode with visual banners showing what WOULD be deleted |
| `-UseParallel` | Parallel processing for multiple computers with throttle control |
| `-HtmlReport` | Generate professional HTML reports with CSS styling and badges |
| `-BackupPath` | ZIP compression backup before deletion |
| `-ConfigFile` | JSON configuration file support with schema validation |
| `-FixCorruption` | Interactive corruption repair with admin approval |
| `-VerifyIntegrity` | SHA256 hash verification of the script before execution |
| `-Detailed` | Per-folder size breakdown (Documents, Downloads, Desktop, AppData, etc.) |
| `-Credential` | PSCredential support for cross-domain remote authentication |
| Event Log | Windows Event Log entries for auditing (Event IDs 1000-1012) |
| Email Notifications | SMTP alerts with summary reports for scheduled tasks |
| Progress Bars | Visual progress indicators for long operations |
| Age Color Coding | Console output color-coded by profile age |
| Mass Deletion Safeguard | Warns on >50 profile deletions, requires "YES" confirmation |

### Graphical User Interface

Launch the GUI with the `-UI` parameter:

```powershell
.\DelprofPS.ps1 -UI
```

The WPF interface provides five tabs:

- **Connection** - Toggle between local and remote computer targeting
- **Profiles** - DataGrid listing all profiles with select/deselect all, refresh, and force-remove buttons
- **Filters** - Sliders and dropdowns for days inactive, age method, profile type, include/exclude patterns, and size filters
- **Actions** - Operation mode (Preview/Delete), option checkboxes, and parallel processing settings
- **Output & Reporting** - Backup path, log/CSV/HTML export paths, and email notification settings

The GUI also includes Load/Save Config buttons, a real-time output console, and a progress bar.

### Security Hardening

- **Config file schema validation** - JSON configs validated against expected types and ranges before applying
- **ComputerName input sanitisation** - Hostnames validated against RFC-compliant patterns to prevent injection
- **Log integrity hashing** - SHA256 hash appended to log files at end of each run for tamper detection
- **Script self-integrity verification** - `-VerifyIntegrity` checks the script's SHA256 hash at startup against `DelprofPS.sha256`

---

## Requirements

- PowerShell 5.1 or later
- Administrative privileges on target computers
- Remote management enabled for remote computer processing (WinRM/PSRemoting)
- .NET Framework 4.5+ (for WPF GUI)

## Installation

### Option 1: Clone the Repository

```powershell
git clone https://github.com/ArMaTeC/DelprofPS.git
```

### Option 2: Download Directly

```powershell
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/ArMaTeC/DelprofPS/main/delprofPS.ps1" -OutFile "DelprofPS.ps1"
```

### Option 3: Import as a PowerShell Module

```powershell
Import-Module .\DelprofPS.psd1

# The module exports the Show-DelprofPSGUI function
Show-DelprofPSGUI
```

### Optional Setup

- Copy and customise `DelprofPS.config.json` for your environment
- Generate the integrity hash: `(Get-FileHash .\delprofPS.ps1 -Algorithm SHA256).Hash | Out-File .\DelprofPS.sha256`

---

## Quick Start

### Basic Usage

```powershell
# List all profiles older than 30 days (dry run)
.\DelprofPS.ps1

# Launch the graphical interface
.\DelprofPS.ps1 -UI

# Preview what would be deleted
.\DelprofPS.ps1 -DaysInactive 60 -Preview

# Delete profiles older than 60 days
.\DelprofPS.ps1 -DaysInactive 60 -Delete

# Interactive CLI selection mode
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

# Verify script integrity before execution
.\DelprofPS.ps1 -VerifyIntegrity -DaysInactive 60 -Delete
```

---

## Demo Scripts

The `StartScripts/` folder contains 16 ready-to-run demo scripts and an interactive launcher:

```powershell
# Launch the interactive demo menu
.\StartScripts\00-Launcher.ps1
```

| Script | Description |
| --- | --- |
| `01-BasicListProfiles.ps1` | Basic profile listing (dry-run, default mode) |
| `02-TestMode.ps1` | Test / validate mode (check prerequisites) |
| `03-PreviewMode.ps1` | Preview mode (simulate what would be deleted) |
| `04-ShowSpaceDetailed.ps1` | Disk space and detailed folder breakdown |
| `05-AgeCalculationMethods.ps1` | Age calculation methods (NTUSER, Path, Registry, Logon) |
| `06-FilteringProfiles.ps1` | Include / exclude / size filtering |
| `07-HTMLReport.ps1` | Generate HTML report |
| `08-CSVExportAndLogging.ps1` | CSV export and log file output |
| `09-BackupAndDelete.ps1` | Backup profiles before deletion |
| `10-ConfigFileUsage.ps1` | Load settings from JSON config file |
| `11-CorruptionDetection.ps1` | Corrupted profile detection |
| `12-InteractiveMode.ps1` | Interactive profile selection |
| `13-DeleteProfiles.ps1` | Delete profiles (with safety prompts) |
| `14-GUIMode.ps1` | Graphical user interface (WPF GUI) |
| `15-ScheduledTask.ps1` | Scheduled task / quiet mode example |
| `16-RemoteComputers.ps1` | Remote computer targeting |

---

## Parameters

| Parameter | Description | Default |
| --- | --- | --- |
| `ComputerName` | Target computer(s). Accepts pipeline, arrays, CSV | `$env:COMPUTERNAME` |
| `DaysInactive` | Minimum days of inactivity (0-3650) | `30` |
| `AgeCalculation` | Method: NTUSER_DAT, ProfilePath, Registry, LastLogon, LastLogoff | `NTUSER_DAT` |
| `Include` | Wildcard patterns to include | None |
| `Exclude` | Wildcard patterns to exclude | None |
| `Delete` | Actually delete profiles (without this, runs in dry-run mode) | `False` |
| `Force` | Skip confirmations and ignore non-critical errors | `False` |
| `UI` | Launch the WPF graphical user interface | `False` |
| `Preview` | Simulation mode showing what would be deleted | `False` |
| `Interactive` | Interactive CLI selection mode with visual menu | `False` |
| `Test` | Test connectivity and prerequisites only | `False` |
| `IgnoreActiveSessions` | Allow deleting active sessions (**dangerous**) | `False` |
| `UnloadHives` | Unload registry hives before deletion | `False` |
| `MaxRetries` | Retry attempts for locked files (0-50) | `3` |
| `RetryDelaySeconds` | Seconds between retries (1-60) | `2` |
| `OutputPath` | CSV export path | None |
| `LogPath` | Log file path | None |
| `Quiet` | Suppress console output | `False` |
| `ShowSpace` | Display disk space per profile | `False` |
| `Detailed` | Show per-folder size breakdown | `False` |
| `IncludeSystemProfiles` | Include system profiles (Default, Public, etc.) | `False` |
| `IncludeSpecialProfiles` | Include special accounts (SYSTEM, NetworkService, etc.) | `False` |
| `MinProfileSizeMB` | Only include profiles larger than this size | None |
| `MaxProfileSizeMB` | Only include profiles smaller than this size | None |
| `IncludeCorrupted` | Include corrupted profiles in processing | `False` |
| `FixCorruption` | Enable interactive corruption repair mode | `False` |
| `ProfileType` | Filter: Local, Roaming, Temporary, Mandatory, All | `All` |
| `HtmlReport` | HTML report output path | None |
| `BackupPath` | Backup directory (profiles ZIPped before deletion) | None |
| `ConfigFile` | JSON configuration file path | None |
| `UseParallel` | Enable parallel processing for multiple computers | `False` |
| `ThrottleLimit` | Max parallel threads (1-100) | `5` |
| `SmtpServer` | SMTP server for email notifications | None |
| `EmailTo` | Email recipient address | None |
| `EmailFrom` | Email sender address | `delprofps@<computername>` |
| `VerifyIntegrity` | Verify script SHA256 hash before execution | `False` |
| `Credential` | PSCredential for remote computer authentication | Current identity |

---

## Configuration File

Create a `DelprofPS.config.json` file to store default settings. All fields are validated against expected types and ranges at load time.

```json
{
    "DaysInactive": 60,
    "AgeCalculation": "NTUSER_DAT",
    "Exclude": ["*admin*", "*service*", "Administrator*"],
    "Include": [],
    "ProfileType": "All",
    "MaxRetries": 3,
    "RetryDelaySeconds": 2,
    "MinProfileSizeMB": 0,
    "MaxProfileSizeMB": 0,
    "UnloadHives": true,
    "IncludeCorrupted": false,
    "IncludeSystemProfiles": false,
    "IncludeSpecialProfiles": false,
    "LogPath": "C:\\Logs\\DelprofPS.log",
    "OutputPath": "C:\\Logs\\DelprofPS_Results.csv",
    "HtmlReport": "C:\\Logs\\DelprofPS_Report.html",
    "BackupPath": "C:\\Backups\\Profiles",
    "SmtpServer": "mail.company.com",
    "EmailTo": "admin@company.com",
    "EmailFrom": "delprofps@server01"
}
```

---

## Safety Features

DelprofPS includes multiple layers of protection:

1. **Default Dry-Run Mode** - Must use `-Delete` to actually remove profiles
2. **Active Session Protection** - Skips logged-in users (unless `-IgnoreActiveSessions`)
3. **System Profile Protection** - Excludes Default, Public, SYSTEM, etc.
4. **Registry Hive Unloading** - `-UnloadHives` safely unloads hives before deletion
5. **Backup Capability** - `-BackupPath` creates ZIP backups before deletion
6. **Mass Deletion Warning** - Warns on >50 profiles, requires "YES" confirmation
7. **Corruption Repair Safety** - `-FixCorruption` requires interactive admin approval for each fix
8. **Comprehensive Logging** - File logging and Windows Event Log integration
9. **ShouldProcess Support** - `-WhatIf` and `-Confirm` standard PowerShell parameters
10. **Script Integrity Verification** - `-VerifyIntegrity` checks SHA256 hash at startup
11. **Config Schema Validation** - JSON config values validated before applying
12. **Input Sanitisation** - ComputerName values validated against RFC-compliant hostname patterns
13. **Log Integrity Hashing** - SHA256 hash appended to log files for tamper detection

---

## Event Log IDs

| Event ID | Level | Description |
| --- | --- | --- |
| 1000 | Information | Script started |
| 1001 | Information | HTML report generated |
| 1002 | Information | Script completed |
| 1005 | Error | Admin rights required |
| 1010 | Information | Profile deleted successfully |
| 1011 | Error | Profile deletion failed |
| 1012 | Information | Corruption repair action taken |

---

## Examples

### Example 1: Basic Dry Run

```powershell
# See what profiles exist and their ages
.\DelprofPS.ps1 -DaysInactive 30
```

### Example 2: Launch the GUI

```powershell
# Open the WPF graphical interface
.\DelprofPS.ps1 -UI
```

### Example 3: Preview Mode

```powershell
# See what WOULD be deleted without actually deleting
.\DelprofPS.ps1 -DaysInactive 60 -Preview -ShowSpace
```

### Example 4: Interactive Selection

```powershell
# Use arrow keys to navigate, Space to toggle, Enter to confirm
.\DelprofPS.ps1 -DaysInactive 90 -Interactive
```

### Example 5: Delete with Backup

```powershell
# Backup profiles before deletion
.\DelprofPS.ps1 -Delete -DaysInactive 60 -BackupPath "D:\Backups"
```

### Example 6: Remote Computers

```powershell
# Process multiple computers
$computers = @("SERVER01", "SERVER02", "SERVER03")
.\DelprofPS.ps1 -ComputerName $computers -Delete -DaysInactive 90
```

### Example 7: Remote with Credentials

```powershell
# Authenticate to remote computers with alternate credentials
$cred = Get-Credential
.\DelprofPS.ps1 -ComputerName "SERVER01" -Credential $cred -DaysInactive 60
```

### Example 8: Detailed Analysis

```powershell
# Show folder breakdowns and export to HTML
.\DelprofPS.ps1 -Detailed -ShowSpace -HtmlReport "C:\Reports\profiles.html"
```

### Example 9: Test Connectivity

```powershell
# Validate access to computers without processing
.\DelprofPS.ps1 -ComputerName (Get-Content servers.txt) -Test
```

### Example 10: Email Notification

```powershell
# Send email report after completion
.\DelprofPS.ps1 -Delete -DaysInactive 60 `
    -SmtpServer "smtp.office365.com" `
    -EmailTo "it-team@company.com" `
    -EmailFrom "delprofps@mgmt-server"
```

### Example 11: Use Configuration File

```powershell
# Load settings from JSON file
.\DelprofPS.ps1 -ConfigFile "C:\Config\delprof.json" -Delete
```

### Example 12: Parallel Processing

```powershell
# Process many computers in parallel
.\DelprofPS.ps1 -ComputerName (Get-Content 100-servers.txt) `
    -UseParallel -ThrottleLimit 20 `
    -Delete -DaysInactive 120
```

### Example 13: Corruption Repair Mode

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

---

## Building and Development

### Build System

The project includes `Build.ps1` for automated linting, testing, hashing, and packaging:

```powershell
# Run the full build pipeline (lint -> test -> hash -> package)
.\Build.ps1 -Task All

# Individual tasks
.\Build.ps1 -Task Lint       # Run PSScriptAnalyzer
.\Build.ps1 -Task Test       # Run Pester tests
.\Build.ps1 -Task Hash       # Generate SHA256 integrity hash
.\Build.ps1 -Task Package    # Create release ZIP in ./release/
```

### CI/CD Pipeline

GitHub Actions runs automatically on pushes to `main`/`develop` and on pull requests:

1. **PSScriptAnalyzer Lint** - Checks `delprofPS.ps1` for errors
2. **Pester Tests** - Runs the full test suite via `DelprofPS.Pester.Tests.ps1`
3. **Module Manifest Validation** - Verifies `DelprofPS.psd1` is valid
4. **Release Packaging** - Creates a release ZIP on pushes to `main`

### Development Dependencies

```powershell
Install-Module -Name PSScriptAnalyzer -Scope CurrentUser -Force
Install-Module -Name Pester -Scope CurrentUser -Force -MinimumVersion 5.0
```

### Running Tests

```powershell
# Quick test run
.\DelprofPS.Tests.ps1

# Pester framework tests
.\Build.ps1 -Task Test
```

---

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

#### Integrity verification fails

- Regenerate the hash file: `(Get-FileHash .\delprofPS.ps1 -Algorithm SHA256).Hash | Out-File .\DelprofPS.sha256`
- Use `-Force` with `-VerifyIntegrity` to bypass the check

### Debug Mode

```powershell
# Enable verbose output
.\DelprofPS.ps1 -Verbose -LogPath "C:\Logs\debug.log"
```

---

## Contributing

Contributions are welcome! See [CONTRIBUTING.md](CONTRIBUTING.md) for full guidelines, including:

- Code standards and PSScriptAnalyzer linting rules
- Commit message conventions (`feat:`, `fix:`, `docs:`, etc.)
- PR checklist and testing requirements

Quick start:

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Run `.\Build.ps1 -Task Lint` and `.\Build.ps1 -Task Test`
4. Commit your changes (`git commit -m 'feat: Add AmazingFeature'`)
5. Push to the branch (`git push origin feature/AmazingFeature`)
6. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Original [Delprof2](https://helgeklein.com/delprof2/) by Helge Klein
- PowerShell community for best practices and feedback

## Support

For issues, questions, or feature requests, please [open an issue](https://github.com/ArMaTeC/DelprofPS/issues).

---

**Disclaimer**: This tool modifies user data. Always test in non-production environments first. The authors are not responsible for data loss or system damage resulting from the use of this script.
