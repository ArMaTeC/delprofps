# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added - Quality & Security Hardening

#### Strategy & Conceptualisation
- **PowerShell module manifest** (`DelprofPS.psd1`) for PSGallery-ready distribution
- **CONTRIBUTING.md** with development guidelines, PR checklist, and commit conventions
- **PSScriptAnalyzer settings** (`PSScriptAnalyzerSettings.psd1`) for enforced code quality
- **Fixed repository URLs** in README (replaced placeholder `yourusername`)

#### Security & Data Integrity
- **Config file schema validation** - JSON configs validated against expected types and ranges before applying
- **ComputerName input sanitisation** - Hostnames validated against RFC-compliant patterns to prevent injection
- **Log integrity hashing** - SHA256 hash appended to log files at end of each run for tamper detection
- **Script self-integrity verification** - SHA256 hash check on script file at startup with optional `-VerifyIntegrity` parameter

#### DevOps & Infrastructure
- **GitHub Actions CI/CD pipeline** (`.github/workflows/ci.yml`) with PSScriptAnalyzer linting and Pester test execution
- **Build & release script** (`Build.ps1`) for automated versioning, linting, testing, and packaging
- **Pester test wrapper** (`DelprofPS.Pester.Tests.ps1`) bridging existing test suite into Pester framework

#### GUI Enhancements
- **Profiles tab** with DataGrid listing all current profiles (username, path, SID, size, last modified, status)
- **Force Remove Selected** button for targeted profile removal with confirmation dialog
- **Select All / Deselect All** buttons for batch selection
- **Refresh Profiles** button supporting local and remote computer scanning

## [2.0.0] - 2024

### Added - Major Enterprise Release

#### Core Features

- **Multi-computer support** - Process single or multiple computers via pipeline, arrays, or CSV import
- **Age calculation methods** - NTUSER.DAT, ProfilePath, Registry timestamps, LastLogon, LastLogoff
- **Active session detection** - Multi-method detection using quser, WMI, and explorer.exe process ownership
- **Flexible filtering** - Include/exclude by username patterns, SID patterns, profile size thresholds
- **Profile type detection** - Local, Roaming, Temporary, Mandatory, and Corrupted profile identification
- **Disk space reporting** - Calculate and display profile sizes with human-readable formatting
- **Registry hive unloading** - Safely unload loaded registry hives before profile deletion
- **Retry logic** - Configurable retry attempts with delays for locked files

#### Enterprise Features

- **Interactive mode (`-Interactive`)** - Visual menu system with keyboard navigation (arrow keys, space toggle, enter confirm)
- **Test mode (`-Test`)** - Validate prerequisites and connectivity without making changes
- **Preview mode (`-Preview`)** - Simulation mode with visual banners showing what WOULD be deleted
- **Parallel processing (`-UseParallel`)** - Multi-threaded processing for large computer lists with throttle control
- **HTML reporting (`-HtmlReport`)** - Professional styled HTML reports with CSS, badges, and color-coded tables
- **Windows Event Log integration** - Audit trail with specific Event IDs (1000-1011)
- **Profile backup (`-BackupPath`)** - Automatic ZIP compression of profiles before deletion
- **JSON configuration (`-ConfigFile`)** - External configuration file support for reusable settings
- **Email notifications** - SMTP alerts with summary reports for scheduled task runs
- **Progress bars** - Visual progress indicators for multi-computer operations
- **Age-based color coding** - Console output color-coded by profile age (green→yellow→magenta→red)
- **Detailed view (`-Detailed`)** - Per-folder size breakdown (Documents, Downloads, Desktop, AppData, Pictures, Videos, Music)
- **Mass deletion safeguard** - Automatic warning when >50 profiles would be deleted (requires "YES" confirmation)
- **Dry-run preview** - Summary shows "Would delete" counts when not in delete mode
- **Top 5 largest profiles** - Automatic identification of largest profiles in summary
- **Age breakdown analysis** - Profiles grouped by age ranges with totals

#### Safety Features

- Default dry-run mode (must use `-Delete` to actually remove profiles)
- Active session protection (skips logged-in users unless `-IgnoreActiveSessions`)
- System profile protection (excludes Default, Public, SYSTEM, etc.)
- ShouldProcess support (`-WhatIf` and `-Confirm`)
- Comprehensive logging to file and Windows Event Log
- Registry backup and safety checks

### Security

- Protected SID list for system accounts
- Admin privilege validation
- Connection testing before operations

## [1.0.0] - Initial Release

### Added

- Basic profile enumeration from registry
- Simple age-based filtering
- Profile deletion with registry cleanup
- Basic logging support

---

## Event Log Reference

| Event ID | Level | Description |
| ---------- | ------- | ------------- |
| 1000 | Information | Script started |
| 1001 | Information | HTML report generated |
| 1002 | Information | Script completed |
| 1005 | Error | Admin rights required |
| 1010 | Information | Profile deleted successfully |
| 1011 | Error | Profile deletion failed |

---

## Migration Guide

### From Delprof2

Delprof2-PS is a drop-in replacement with enhanced capabilities:

| Delprof2 | Delprof2-PS Equivalent |
| ---------- | ---------------------- |
| `/c` | `-ComputerName` |
| `/d` | `-DaysInactive` |
| `/q` | `-Quiet` |
| `/i` | `-Include` |
| `/x` | `-Exclude` |
| `/l` | `-LogPath` |
| `/u` | `-UnloadHives` |

### New Capabilities Not in Delprof2

- Interactive mode (`-Interactive`)
- HTML reporting (`-HtmlReport`)
- Email notifications (`-SmtpServer`, `-EmailTo`)
- Profile backup (`-BackupPath`)
- Parallel processing (`-UseParallel`)
- JSON configuration (`-ConfigFile`)
- Preview mode (`-Preview`)
- Test mode (`-Test`)

---

## Planned Features

### Version 2.1.0 (Planned)

- [ ] Azure AD / Entra ID profile support
- [ ] Cloud profile synchronization detection
- [ ] Integration with Microsoft Endpoint Manager
- [ ] Power BI report templates

### Version 2.2.0 (Planned)

- [ ] GUI version (WinForms/WPF)
- [ ] REST API endpoint
- [ ] Web-based dashboard
- [ ] Automated scheduling service

### Version 3.0.0 (Future)

- [ ] Cross-platform support (PowerShell Core on Linux)
- [ ] Container deployment options
- [ ] Integration with monitoring systems (SCOM, Nagios)

---

## Contributors

- **Karl Lawrence** - Original author and maintainer
- Community contributors - See [CONTRIBUTORS.md](CONTRIBUTORS.md)

---

## Feedback

Have a feature request or found a bug? Please [open an issue](https://github.com/yourusername/Delprof2-PS/issues).
