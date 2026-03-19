# Contributing to DelprofPS

Thank you for your interest in contributing to DelprofPS! This document provides guidelines and instructions for contributing.

## Getting Started

1. **Fork** the repository on GitHub
2. **Clone** your fork locally:
   ```bash
   git clone https://github.com/YOUR-USERNAME/DelprofPS.git
   ```
3. Create a **feature branch**:
   ```bash
   git checkout -b feature/your-feature-name
   ```

## Development Requirements

- PowerShell 5.1 or later
- PSScriptAnalyzer (for linting)
- Administrator privileges (for running tests)

### Installing Development Dependencies

```powershell
Install-Module -Name PSScriptAnalyzer -Scope CurrentUser -Force
Install-Module -Name Pester -Scope CurrentUser -Force -MinimumVersion 5.0
```

## Code Standards

### Style

- Follow the [PowerShell Practice and Style Guide](https://poshcode.gitbook.io/powershell-practice-and-style/)
- Use **PascalCase** for function names and parameters
- Use **camelCase** for local variables
- All public functions must include comment-based help (`.SYNOPSIS`, `.DESCRIPTION`, `.PARAMETER`, `.EXAMPLE`)

### Linting

Run PSScriptAnalyzer before submitting:

```powershell
Invoke-ScriptAnalyzer -Path .\delprofPS.ps1 -Settings .\PSScriptAnalyzerSettings.psd1
```

All code must pass with **zero errors**. Warnings should be addressed where practical.

### Testing

- Run the existing test suite before submitting changes:
  ```powershell
  .\DelprofPS.Tests.ps1
  ```
- Add tests for any new functionality
- Tests must pass on a clean Windows 10/11 or Server 2016+ system

## Making Changes

### Commit Messages

Use clear, descriptive commit messages:

```
feat: Add Azure AD profile detection
fix: Resolve registry hive unload failure on Server 2022
docs: Update parameter documentation for -HtmlReport
test: Add age calculation edge case tests
```

Prefix with: `feat:`, `fix:`, `docs:`, `test:`, `refactor:`, `chore:`

### What to Change

- **delprofPS.ps1** — Main script (core logic, functions, embedded GUI)
- **DelprofPS-GUI.ps1** — Standalone GUI file (keep in sync with embedded copy)
- **DelprofPS.Tests.ps1** — Test suite
- **StartScripts/** — Example/demo scripts

> **Important:** The GUI is defined in both `DelprofPS-GUI.ps1` and embedded within `delprofPS.ps1`. Changes to one **must** be mirrored in the other.

## Pull Request Process

1. Ensure your code passes PSScriptAnalyzer with no errors
2. Run the full test suite and confirm all tests pass
3. Update `CHANGELOG.md` with your changes under an `[Unreleased]` section
4. Update `README.md` if you've added new parameters or features
5. Submit a pull request with a clear description of changes

### PR Checklist

- [ ] Code passes `Invoke-ScriptAnalyzer` with no errors
- [ ] All existing tests pass
- [ ] New tests added for new functionality
- [ ] CHANGELOG.md updated
- [ ] README.md updated (if applicable)
- [ ] GUI changes mirrored in both files
- [ ] Comment-based help added for new functions

## Reporting Issues

When reporting bugs, please include:

- PowerShell version (`$PSVersionTable.PSVersion`)
- Operating system and version
- Steps to reproduce
- Expected vs actual behaviour
- Relevant log output (use `-Verbose -LogPath debug.log`)

## Security Vulnerabilities

If you discover a security vulnerability, please **do not** open a public issue. Instead, email the maintainer directly at the address in the repository.

## License

By contributing, you agree that your contributions will be licensed under the [MIT License](LICENSE).
