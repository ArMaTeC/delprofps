#Requires -Version 5.1
<#
.SYNOPSIS
    Build and release script for DelprofPS.

.DESCRIPTION
    Automates linting, testing, versioning, packaging, and hash generation for DelprofPS releases.

.PARAMETER Task
    The build task to run: Lint, Test, Package, Hash, All (default).

.EXAMPLE
    .\Build.ps1 -Task All

.EXAMPLE
    .\Build.ps1 -Task Package
#>
param(
    [ValidateSet('Lint', 'Test', 'Package', 'Hash', 'All')]
    [string]$Task = 'All'
)

$ErrorActionPreference = 'Stop'
$script:ProjectRoot = Split-Path $PSScriptRoot -Parent
$script:Version = (Import-PowerShellDataFile "$script:ProjectRoot\src\DelprofPS.psd1").ModuleVersion
$script:ReleaseDir = Join-Path $script:ProjectRoot 'release'
$script:PassedSteps = 0
$script:FailedSteps = 0

function Write-BuildLog {
    param([string]$Message, [string]$Level = 'INFO')
    $color = switch ($Level) {
        'INFO' { 'Cyan' }
        'SUCCESS' { 'Green' }
        'WARNING' { 'Yellow' }
        'ERROR' { 'Red' }
        default { 'White' }
    }
    Write-Host "[BUILD] [$Level] $Message" -ForegroundColor $color
}

function Invoke-Lint {
    Write-BuildLog "Running PSScriptAnalyzer..." 'INFO'

    $analyzerInstalled = Get-Module -ListAvailable -Name PSScriptAnalyzer
    if (-not $analyzerInstalled) {
        Write-BuildLog "Installing PSScriptAnalyzer..." 'WARNING'
        Install-Module -Name PSScriptAnalyzer -Force -Scope CurrentUser
    }

    $settingsPath = Join-Path $script:ProjectRoot 'config\PSScriptAnalyzerSettings.psd1'
    $scripts = @('delprofPS.ps1')
    $totalErrors = 0

    foreach ($scriptFile in $scripts) {
        $path = Join-Path $script:ProjectRoot $scriptFile
        if (Test-Path $path) {
            $results = Invoke-ScriptAnalyzer -Path $path -Settings $settingsPath -Severity Error
            if ($results) {
                Write-BuildLog "$scriptFile : $($results.Count) error(s)" 'ERROR'
                $results | Format-Table RuleName, Line, Message -AutoSize
                $totalErrors += $results.Count
            }
            else {
                Write-BuildLog "$scriptFile : PASSED" 'SUCCESS'
            }
        }
    }

    if ($totalErrors -gt 0) {
        $script:FailedSteps++
        Write-BuildLog "Lint failed with $totalErrors error(s)" 'ERROR'
        return $false
    }
    $script:PassedSteps++
    Write-BuildLog "Lint passed" 'SUCCESS'
    return $true
}

function Invoke-Test {
    Write-BuildLog "Running tests..." 'INFO'

    $pesterInstalled = Get-Module -ListAvailable -Name Pester | Where-Object { $_.Version -ge '5.0' }
    if (-not $pesterInstalled) {
        Write-BuildLog "Installing Pester 5+..." 'WARNING'
        Install-Module -Name Pester -Force -Scope CurrentUser -MinimumVersion 5.0
    }

    $testFile = Join-Path $script:ProjectRoot 'tests\DelprofPS.Pester.Tests.ps1'
    if (Test-Path $testFile) {
        try {
            $config = New-PesterConfiguration
            $config.Run.Path = $testFile
            $config.Output.Verbosity = 'Detailed'
            $pesterResult = Invoke-Pester -Configuration $config

            if ($pesterResult.FailedCount -gt 0) {
                $script:FailedSteps++
                Write-BuildLog "Tests failed: $($pesterResult.FailedCount) failure(s)" 'ERROR'
                return $false
            }
            $script:PassedSteps++
            Write-BuildLog "Tests passed: $($pesterResult.PassedCount) passed, $($pesterResult.FailedCount) failed" 'SUCCESS'
            return $true
        }
        catch {
            $script:FailedSteps++
            Write-BuildLog "Test execution error: $($_.Exception.Message)" 'ERROR'
            return $false
        }
    }
    else {
        Write-BuildLog "Pester test file not found, skipping" 'WARNING'
        return $true
    }
}

function Invoke-Package {
    Write-BuildLog "Creating release package..." 'INFO'

    if (-not (Test-Path $script:ReleaseDir)) {
        New-Item -ItemType Directory -Path $script:ReleaseDir -Force | Out-Null
    }

    $packageName = "DelprofPS-v$script:Version.zip"
    $packagePath = Join-Path $script:ReleaseDir $packageName

    $filesToInclude = @(
        'delprofPS.ps1',
        'src\DelprofPS.psd1',
        'config\DelprofPS.config.json',
        'tests\DelprofPS.Tests.ps1',
        'tests\DelprofPS.Pester.Tests.ps1',
        'config\PSScriptAnalyzerSettings.psd1',
        'docs\README.md',
        'docs\CHANGELOG.md',
        'docs\CONTRIBUTING.md',
        'docs\LICENSE'
    )

    $tempDir = Join-Path $env:TEMP "DelprofPS-Package-$(Get-Date -Format 'yyyyMMddHHmmss')"
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

    foreach ($file in $filesToInclude) {
        $sourcePath = Join-Path $script:ProjectRoot $file
        if (Test-Path $sourcePath) {
            Copy-Item -Path $sourcePath -Destination $tempDir
        }
        else {
            Write-BuildLog "File not found for packaging: $file" 'WARNING'
        }
    }

    # Include examples folder
    $examplesDir = Join-Path $script:ProjectRoot 'examples'
    if (Test-Path $examplesDir) {
        Copy-Item -Path $examplesDir -Destination (Join-Path $tempDir 'examples') -Recurse
    }

    # Generate integrity hash
    $mainScript = Join-Path $tempDir 'delprofPS.ps1'
    if (Test-Path $mainScript) {
        $hash = (Get-FileHash -Path $mainScript -Algorithm SHA256).Hash
        $hash | Out-File (Join-Path $tempDir 'DelprofPS.sha256') -Encoding UTF8
        Write-BuildLog "SHA256 hash: $hash" 'INFO'
    }

    # Create ZIP
    if (Test-Path $packagePath) { Remove-Item $packagePath -Force }
    Compress-Archive -Path "$tempDir\*" -DestinationPath $packagePath -CompressionLevel Optimal

    # Cleanup temp
    Remove-Item $tempDir -Recurse -Force

    $script:PassedSteps++
    Write-BuildLog "Package created: $packagePath ($('{0:N2}' -f ((Get-Item $packagePath).Length / 1KB)) KB)" 'SUCCESS'
    return $true
}

function Invoke-Hash {
    Write-BuildLog "Generating script integrity hash..." 'INFO'

    $mainScript = Join-Path $script:ProjectRoot 'delprofPS.ps1'
    if (Test-Path $mainScript) {
        $hash = (Get-FileHash -Path $mainScript -Algorithm SHA256).Hash
        $hashFile = Join-Path $script:ProjectRoot 'assets\DelprofPS.sha256'
        $hash | Out-File $hashFile -Encoding UTF8
        Write-BuildLog "SHA256: $hash" 'INFO'
        Write-BuildLog "Hash written to assets\DelprofPS.sha256" 'SUCCESS'
        $script:PassedSteps++
        return $true
    }
    else {
        Write-BuildLog "Main script not found" 'ERROR'
        $script:FailedSteps++
        return $false
    }
}

# Main execution
Write-Host ""
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "  DelprofPS Build System v$script:Version" -ForegroundColor Cyan
Write-Host "  Task: $Task" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

$success = $true

switch ($Task) {
    'Lint' { $success = Invoke-Lint }
    'Test' { $success = Invoke-Test }
    'Package' { $success = Invoke-Package }
    'Hash' { $success = Invoke-Hash }
    'All' {
        $success = Invoke-Lint
        if ($success) { $success = Invoke-Test }
        if ($success) { $success = Invoke-Hash }
        if ($success) { $success = Invoke-Package }
    }
}

Write-Host ""
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "  Build Summary" -ForegroundColor Cyan
Write-Host "  Passed: $script:PassedSteps | Failed: $script:FailedSteps" -ForegroundColor $(if ($script:FailedSteps -gt 0) { 'Red' } else { 'Green' })
Write-Host "=====================================" -ForegroundColor Cyan

if (-not $success) {
    exit 1
}
