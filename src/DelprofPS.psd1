@{
    # Module manifest for DelprofPS
    RootModule = '..\delprofPS.ps1'
    ModuleVersion = '2.0.0'
    GUID = 'a3e4b8c1-2d5f-4e6a-9b7c-8d1e0f2a3b4c'
    Author = 'Karl Lawrence'
    CompanyName = 'GCI Network Solutions'
    Copyright = '(c) 2026 Karl Lawrence. All rights reserved.'
    Description = 'Enterprise-grade PowerShell replacement for Delprof2. Provides advanced user profile management with safety mechanisms, reporting, parallel processing, and a modern GUI.'

    # Minimum version of PowerShell required
    PowerShellVersion = '5.1'

    # Functions to export
    FunctionsToExport = @()

    # Cmdlets to export
    CmdletsToExport = @()

    # Variables to export
    VariablesToExport = @()

    # Aliases to export
    AliasesToExport = @()

    # Required assemblies
    RequiredAssemblies = @()

    # Files to package with this module
    FileList = @()

    # Private data / PSGallery metadata
    PrivateData = @{
        PSData = @{
            Tags = @('Profile', 'Management', 'Delprof2', 'UserProfile', 'Cleanup', 'Enterprise', 'GUI', 'Windows')
            LicenseUri = 'https://github.com/ArMaTeC/DelprofPS/blob/main/LICENSE'
            ProjectUri = 'https://github.com/ArMaTeC/DelprofPS'
            ReleaseNotes = @'
v2.0.0 - Major Enterprise Release
- Multi-computer support with parallel processing
- WPF GUI with profile browser and force-removal
- Multiple age calculation methods
- Interactive mode, Preview mode, Test mode
- HTML reporting, CSV export, Event Log integration
- Profile backup, corruption repair, email notifications
'@
        }
    }
}
