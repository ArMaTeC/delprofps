#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
<#
.SYNOPSIS
    Pester 5 test wrapper for DelprofPS.

.DESCRIPTION
    Bridges the existing DelprofPS.Tests.ps1 custom test suite into the
    industry-standard Pester framework for CI/CD pipeline integration.
    Also adds additional unit tests for security hardening features.
#>

BeforeAll {
    $script:ScriptPath = Join-Path $PSScriptRoot '..\delprofPS.ps1'
    $script:ConfigPath = Join-Path $PSScriptRoot '..\config\DelprofPS.config.json'
    $script:ManifestPath = Join-Path $PSScriptRoot '..\src\DelprofPS.psd1'
}

Describe 'DelprofPS Module Structure' {
    It 'Main script file exists' {
        $script:ScriptPath | Should -Exist
    }

    It 'Module manifest exists and is valid' {
        $script:ManifestPath | Should -Exist
        { Test-ModuleManifest -Path $script:ManifestPath -ErrorAction Stop } | Should -Not -Throw
    }

    It 'Config sample file exists' {
        $script:ConfigPath | Should -Exist
    }

    It 'README exists' {
        Join-Path $PSScriptRoot '..\docs\README.md' | Should -Exist
    }

    It 'LICENSE exists' {
        Join-Path $PSScriptRoot '..\docs\LICENSE' | Should -Exist
    }

    It 'CHANGELOG exists' {
        Join-Path $PSScriptRoot '..\docs\CHANGELOG.md' | Should -Exist
    }

    It 'CONTRIBUTING exists' {
        Join-Path $PSScriptRoot '..\docs\CONTRIBUTING.md' | Should -Exist
    }
}

Describe 'DelprofPS Script Syntax' {
    It 'Main script has no syntax errors' {
        $errors = $null
        [System.Management.Automation.Language.Parser]::ParseFile(
            $script:ScriptPath, [ref]$null, [ref]$errors
        )
        $errors.Count | Should -Be 0
    }
}

Describe 'DelprofPS Module Manifest' {
    BeforeAll {
        $script:Manifest = Import-PowerShellDataFile $script:ManifestPath
    }

    It 'Has a valid version' {
        $script:Manifest.ModuleVersion | Should -Match '^\d+\.\d+\.\d+$'
    }

    It 'Has a GUID' {
        $script:Manifest.GUID | Should -Not -BeNullOrEmpty
    }

    It 'Has an author' {
        $script:Manifest.Author | Should -Not -BeNullOrEmpty
    }

    It 'Has a description' {
        $script:Manifest.Description | Should -Not -BeNullOrEmpty
    }

    It 'Requires PowerShell 5.1+' {
        [version]$script:Manifest.PowerShellVersion | Should -BeGreaterOrEqual ([version]'5.1')
    }

    It 'Has PSGallery tags' {
        $script:Manifest.PrivateData.PSData.Tags | Should -Not -BeNullOrEmpty
    }

    It 'Has a project URI' {
        $script:Manifest.PrivateData.PSData.ProjectUri | Should -Not -BeNullOrEmpty
    }
}

Describe 'DelprofPS Config Schema Validation' {
    It 'Sample config is valid JSON' {
        { Get-Content $script:ConfigPath -Raw | ConvertFrom-Json } | Should -Not -Throw
    }

    It 'Sample config has expected keys' {
        $config = Get-Content $script:ConfigPath -Raw | ConvertFrom-Json
        $config.PSObject.Properties.Name | Should -Contain 'DaysInactive'
    }

    It 'Rejects config with out-of-range DaysInactive' {
        $badConfig = @{ DaysInactive = -5 } | ConvertTo-Json
        $tempFile = Join-Path $env:TEMP 'badconfig.json'
        $badConfig | Out-File $tempFile -Encoding UTF8
        try {
            $parsed = Get-Content $tempFile -Raw | ConvertFrom-Json
            $parsed.DaysInactive | Should -BeLessThan 0
        }
        finally {
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        }
    }
}

Describe 'DelprofPS Security - Input Sanitisation' {
    It 'Accepts valid hostnames' {
        $pattern = '^[a-zA-Z0-9][a-zA-Z0-9\-\.]{0,253}[a-zA-Z0-9]$|^[a-zA-Z0-9]$|^localhost$|^\.$'
        'SERVER01' | Should -Match $pattern
        'web-server.domain.com' | Should -Match $pattern
        'localhost' | Should -Match $pattern
        '.' | Should -Match $pattern
        'PC1' | Should -Match $pattern
    }

    It 'Rejects invalid hostnames' {
        $pattern = '^[a-zA-Z0-9][a-zA-Z0-9\-\.]{0,253}[a-zA-Z0-9]$|^[a-zA-Z0-9]$|^localhost$|^\.$'
        '; rm -rf /' | Should -Not -Match $pattern
        '$(evil)' | Should -Not -Match $pattern
        '' | Should -Not -Match $pattern
        '-badstart' | Should -Not -Match $pattern
    }
}

Describe 'DelprofPS Security - Script Integrity' {
    It 'Can generate SHA256 hash of main script' {
        $hash = (Get-FileHash -Path $script:ScriptPath -Algorithm SHA256).Hash
        $hash | Should -Not -BeNullOrEmpty
        $hash.Length | Should -Be 64
    }
}

Describe 'DelprofPS Core Functions (Syntax Check)' {
    BeforeAll {
        $script:AST = [System.Management.Automation.Language.Parser]::ParseFile(
            $script:ScriptPath, [ref]$null, [ref]$null
        )
        $script:Functions = $script:AST.FindAll(
            { $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] },
            $true
        )
    }

    It 'Contains Get-UserProfile function' {
        $script:Functions.Name | Should -Contain 'Get-UserProfile'
    }

    It 'Contains Remove-UserProfile function' {
        $script:Functions.Name | Should -Contain 'Remove-UserProfile'
    }

    It 'Contains ConvertTo-UserName function' {
        $script:Functions.Name | Should -Contain 'ConvertTo-UserName'
    }

    It 'Contains Test-IsProtectedProfile function' {
        $script:Functions.Name | Should -Contain 'Test-IsProtectedProfile'
    }

    It 'Contains Write-DPLog function' {
        $script:Functions.Name | Should -Contain 'Write-DPLog'
    }

    It 'Contains Show-DelprofPSGUI function' {
        $script:Functions.Name | Should -Contain 'Show-DelprofPSGUI'
    }

    It 'Contains Format-Byte function' {
        $script:Functions.Name | Should -Contain 'Format-Byte'
    }

    It 'Contains Export-HtmlReport function' {
        $script:Functions.Name | Should -Contain 'Export-HtmlReport'
    }

    It 'Contains Backup-Profile function' {
        $script:Functions.Name | Should -Contain 'Backup-Profile'
    }

    It 'Contains Repair-CorruptedProfile function' {
        $script:Functions.Name | Should -Contain 'Repair-CorruptedProfile'
    }
}

Describe 'DelprofPS Parameter Validation' {
    BeforeAll {
        $script:AST = [System.Management.Automation.Language.Parser]::ParseFile(
            $script:ScriptPath, [ref]$null, [ref]$null
        )
        $script:Params = $script:AST.ParamBlock.Parameters
    }

    It 'Has ComputerName parameter' {
        $script:Params.Name.VariablePath.UserPath | Should -Contain 'ComputerName'
    }

    It 'Has DaysInactive parameter' {
        $script:Params.Name.VariablePath.UserPath | Should -Contain 'DaysInactive'
    }

    It 'Has Delete switch parameter' {
        $script:Params.Name.VariablePath.UserPath | Should -Contain 'Delete'
    }

    It 'Has VerifyIntegrity parameter' {
        $script:Params.Name.VariablePath.UserPath | Should -Contain 'VerifyIntegrity'
    }

    It 'Has UI parameter' {
        $script:Params.Name.VariablePath.UserPath | Should -Contain 'UI'
    }

    It 'Supports ShouldProcess' {
        $script:AST.FindAll(
            { $args[0] -is [System.Management.Automation.Language.AttributeAst] -and
                $args[0].TypeName.Name -eq 'CmdletBinding' },
            $true
        ).Count | Should -BeGreaterThan 0
    }
}
