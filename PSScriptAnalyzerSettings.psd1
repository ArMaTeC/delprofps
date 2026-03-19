@{
    # PSScriptAnalyzer settings for DelprofPS
    # Run: Invoke-ScriptAnalyzer -Path . -Settings .\PSScriptAnalyzerSettings.psd1 -Recurse

    Severity = @('Error', 'Warning')

    IncludeRules = @(
        'PSAvoidUsingCmdletAliases',
        'PSAvoidUsingPositionalParameters',
        'PSAvoidUsingWriteHost',
        'PSAvoidGlobalVars',
        'PSUseDeclaredVarsMoreThanAssignments',
        'PSUseShouldProcessForStateChangingFunctions',
        'PSAvoidUsingPlainTextForPassword',
        'PSAvoidUsingConvertToSecureStringWithPlainText',
        'PSUsePSCredentialType',
        'PSAvoidUsingInvokeExpression',
        'PSAvoidUsingWMICmdlet',
        'PSUseApprovedVerbs',
        'PSUseSingularNouns',
        'PSMissingModuleManifestField',
        'PSUseOutputTypeCorrectly',
        'PSProvideCommentHelp',
        'PSUseConsistentWhitespace',
        'PSUseConsistentIndentation',
        'PSAlignAssignmentStatement'
    )

    ExcludeRules = @(
        # Write-Host is intentionally used for coloured console output
        'PSAvoidUsingWriteHost',
        # WMI cmdlets used for remote registry access (CIM not available for StdRegProv)
        'PSAvoidUsingWMICmdlet',
        # PSUseConsistentIndentation produces false positives after the embedded XAML
        # here-string (~300 lines of XML) in delprofPS.ps1. The indent tracker loses
        # sync and flags all subsequent code. Manual review confirms correct indentation.
        'PSUseConsistentIndentation',
        # Write-Host "text" with positional parameter is standard PowerShell idiom
        'PSAvoidUsingPositionalParameters'
    )

    Rules = @{
        # Kept for reference - re-enable when PSScriptAnalyzer improves
        # here-string indent tracking (see ExcludeRules note above)
        PSUseConsistentIndentation = @{
            Enable = $true
            IndentationSize = 4
            Kind = 'space'
            PipelineIndentation = 'IncreaseIndentationForFirstPipeline'
        }

        PSUseConsistentWhitespace = @{
            Enable = $true
            CheckOpenBrace = $true
            CheckOpenParen = $true
            CheckOperator = $true
            CheckSeparator = $true
            CheckInnerBrace = $true
            CheckPipeForRedundantWhitespace = $true
        }

        PSAlignAssignmentStatement = @{
            Enable = $true
            CheckHashtable = $false
        }

        PSProvideCommentHelp = @{
            Enable = $true
            ExportedOnly = $true
            BlockComment = $true
            VSCodeSnippetCorrection = $false
            Placement = 'begin'
        }
    }
}
