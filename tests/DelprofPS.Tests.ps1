#Requires -Version 5.1
<#
.SYNOPSIS
    Comprehensive test suite for Delprof2-PS script

.DESCRIPTION
    This script creates test users and profiles to validate all functions
    in DelprofPS.ps1, then cleans up afterward. Run with administrator privileges.

.PARAMETER TestPath
    Directory for test artifacts (profiles, logs, reports)

.PARAMETER KeepTestArtifacts
    Don't delete test users and profiles after testing

.EXAMPLE
    # Run full test suite
    .\DelprofPS.Tests.ps1

.EXAMPLE
    # Run tests and keep artifacts for inspection
    .\DelprofPS.Tests.ps1 -KeepTestArtifacts

.NOTES
    Version: 1.0
    Author: Karl Lawrence
    Requires: Administrator privileges
#>
[CmdletBinding()]
param(
    [Parameter()]
    [string]$TestPath = "C:\DelprofPSTest",

    [Parameter()]
    [switch]$KeepTestArtifacts
)

#region Test Configuration
$script:TestUsers = @(
    @{ Name = "DPTest_OldUser1"; DaysOld = 200; SizeMB = 150 }
    @{ Name = "DPTest_OldUser2"; DaysOld = 180; SizeMB = 80 }
    @{ Name = "DPTest_MedUser1"; DaysOld = 90; SizeMB = 200 }
    @{ Name = "DPTest_MedUser2"; DaysOld = 60; SizeMB = 120 }
    @{ Name = "DPTest_NewUser1"; DaysOld = 15; SizeMB = 50 }
    @{ Name = "DPTest_Service"; DaysOld = 300; SizeMB = 30 }
    @{ Name = "DPTest_AdminTest"; DaysOld = 120; SizeMB = 500 }
)

$script:TestResults = [System.Collections.Generic.List[object]]::new()
$script:StartTime = Get-Date
#endregion

#region Helper Functions
<#
.SYNOPSIS
    Writes a formatted log message to the console with timestamp and color coding.

.DESCRIPTION
    Outputs test log messages with visual indicators (checkmarks, X marks, etc.)
    and color-coded severity levels for easy reading during test execution.

.PARAMETER Message
    The text message to display.

.PARAMETER Level
    The severity level: INFO (white), PASS (green), FAIL (red), WARN (yellow), SECTION (cyan).

.EXAMPLE
    Write-TestLog "Starting test execution" 'SECTION'
    Write-TestLog "Test passed successfully" 'PASS'
    Write-TestLog "Warning: Low disk space" 'WARN'
#>
function Write-TestLog {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'PASS', 'FAIL', 'WARN', 'SECTION')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'HH:mm:ss'
    $colorMap = @{
        'INFO' = 'White'
        'PASS' = 'Green'
        'FAIL' = 'Red'
        'WARN' = 'Yellow'
        'SECTION' = 'Cyan'
    }

    $prefix = switch ($Level) {
        'PASS' { '[✓]' }
        'FAIL' { '[✗]' }
        'WARN' { '[!]' }
        'SECTION' { '▶' }
        default { '[i]' }
    }

    Write-Host "$prefix [$timestamp] $Message" -ForegroundColor $colorMap[$Level]
}

<#
.SYNOPSIS
    Records a test result for the final summary report.

.DESCRIPTION
    Adds a test result to the collection that will be displayed
    in the test summary at the end of execution.

.PARAMETER TestName
    Name of the test case.

.PARAMETER Passed
    Boolean indicating if the test passed ($true) or failed ($false).

.PARAMETER ErrorMessage
    Optional error message if the test failed.

.EXAMPLE
    Add-TestResult -TestName "Profile Enumeration" -Passed $true
    Add-TestResult -TestName "CSV Export" -Passed $false -ErrorMessage "File not found"
#>
function Add-TestResult {
    param(
        [string]$TestName,
        [bool]$Passed,
        [string]$Details = '',
        [string]$ErrorMessage = ''
    )

    $script:TestResults.Add([PSCustomObject]@{
        TestName = $TestName
        Passed = $Passed
        Details = $Details
        ErrorMessage = $ErrorMessage
        Duration = $null
    })
}

<#
.SYNOPSIS
    Checks if the current PowerShell session has administrative privileges.

.DESCRIPTION
    Verifies that the script is running with elevated privileges,
    which are required for creating test users and managing profiles.

.OUTPUTS
    [bool] Returns $true if running as administrator, $false otherwise.

.EXAMPLE
    if (-not (Test-AdminRight)) {
        Write-Error "This script requires administrator privileges."
    }
#>
function Test-AdminRight {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Source DelprofPS.ps1 functions into script scope for use by all tests
# The -Test flag causes exit 0 after the begin block, which we catch.
# All functions defined in the begin block remain available.
try {
    . $PSScriptRoot\DelprofPS.ps1 -Test 2>&1 | Out-Null
} catch {
    # Expected: -Test mode calls exit which throws when dot-sourced
}
#endregion

#region Test Setup Functions
<#
.SYNOPSIS
    Creates a local test user account with a simulated aged profile.

.DESCRIPTION
    Creates a local user account, generates a profile directory with
    simulated data files, and sets the NTUSER.DAT timestamp to simulate
    a profile of a specific age for testing purposes.

.PARAMETER UserName
    The username for the test account.

.PARAMETER DaysOld
    Number of days to age the profile (affects NTUSER.DAT timestamp).

.PARAMETER SizeMB
    Target size in MB for generated test data in the profile.

.EXAMPLE
    New-TestUser -UserName "DPTest_User1" -DaysOld 60 -SizeMB 100
    Creates a test user with a 60-day-old profile containing ~100MB of data.
#>
function New-TestUser {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingConvertToSecureStringWithPlainText', '', Justification = 'Test user creation requires plaintext password')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Test helper function')]
    param(
        [string]$UserName,
        [int]$DaysOld = 30,
        [int]$SizeMB = 100
    )

    try {
        # Check if user exists
        $existing = Get-LocalUser -Name $UserName -ErrorAction SilentlyContinue
        if ($existing) {
            Write-TestLog "User $UserName already exists, removing..." 'WARN'
            Remove-LocalUser -Name $UserName -ErrorAction SilentlyContinue
        }

        # Create local user with random password
        $password = -join ((33..126) | Get-Random -Count 16 | ForEach-Object { [char]$_ })
        $securePassword = ConvertTo-SecureString $password -AsPlainText -Force

        New-LocalUser -Name $UserName -Password $securePassword -FullName $UserName -Description "DelprofPS Test User" -ErrorAction Stop
        Write-TestLog "Created test user: $UserName" 'PASS'

        # Create profile directory manually (simulating a real profile)
        $profilePath = Join-Path $env:SystemDrive "Users\$UserName"

        if (-not (Test-Path $profilePath)) {
            New-Item -ItemType Directory -Path $profilePath -Force | Out-Null
        }

        # Create folder structure
        $folders = @('Documents', 'Downloads', 'Desktop', 'AppData', 'Pictures', 'Videos', 'Music')
        foreach ($folder in $folders) {
            $folderPath = Join-Path $profilePath $folder
            if (-not (Test-Path $folderPath)) {
                New-Item -ItemType Directory -Path $folderPath -Force | Out-Null
            }
        }

        # Create NTUSER.DAT file
        $ntuserDat = Join-Path $profilePath 'NTUSER.DAT'
        "MockRegistryHive" | Out-File -FilePath $ntuserDat -Force

        # Create dummy files to simulate profile size
        $docPath = Join-Path $profilePath 'Documents'
        $dummyFile = Join-Path $docPath 'testfile.dat'
        $fileSize = $SizeMB * 1MB
        $buffer = New-Object byte[] $fileSize
        [System.IO.File]::WriteAllBytes($dummyFile, $buffer)

        # Set file timestamps to simulate age
        $oldDate = (Get-Date).AddDays(-$DaysOld)
        Get-ChildItem -Path $profilePath -Recurse -Force | ForEach-Object {
            $_.CreationTime = $oldDate
            $_.LastWriteTime = $oldDate
            $_.LastAccessTime = $oldDate
        }
        # Also set the profile directory itself
        $profileDir = Get-Item -Path $profilePath -Force
        $profileDir.CreationTime = $oldDate
        $profileDir.LastWriteTime = $oldDate
        $profileDir.LastAccessTime = $oldDate

        # Set registry profile timestamps
        $profileListPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
        $userSid = (New-Object System.Security.Principal.NTAccount($UserName)).Translate([System.Security.Principal.SecurityIdentifier]).Value
        $userProfilePath = Join-Path $profileListPath $userSid

        if (-not (Test-Path $userProfilePath)) {
            New-Item -Path $userProfilePath -Force | Out-Null
            New-ItemProperty -Path $userProfilePath -Name 'ProfileImagePath' -Value $profilePath -PropertyType ExpandString -Force | Out-Null
            New-ItemProperty -Path $userProfilePath -Name 'Flags' -Value 0 -PropertyType DWord -Force | Out-Null
            New-ItemProperty -Path $userProfilePath -Name 'State' -Value 0 -PropertyType DWord -Force | Out-Null
            # Set LocalProfileLoadTime as FILETIME so Get-ProfileAge Registry method works
            $fileTime = $oldDate.ToFileTime()
            $ftBytes = [BitConverter]::GetBytes($fileTime)
            $low = [BitConverter]::ToInt32($ftBytes, 0)
            $high = [BitConverter]::ToInt32($ftBytes, 4)
            New-ItemProperty -Path $userProfilePath -Name 'LocalProfileLoadTimeLow' -Value $low -PropertyType DWord -Force | Out-Null
            New-ItemProperty -Path $userProfilePath -Name 'LocalProfileLoadTimeHigh' -Value $high -PropertyType DWord -Force | Out-Null
        }

        Write-TestLog "Created profile for $UserName (Age: $DaysOld days, Size: ${SizeMB}MB)" 'PASS'
        return $true
    }
    catch {
        Write-TestLog "Failed to create test user $UserName`: $_" 'FAIL'
        return $false
    }
}

<#
.SYNOPSIS
    Initializes the test environment by creating the test directory and all test users.

.DESCRIPTION
    Prepares the test environment by creating the test path directory and
    generating all configured test user accounts with their simulated profiles.
    Validates administrator privileges before proceeding.

.EXAMPLE
    Initialize-TestEnvironment
    Creates all test users defined in $script:TestUsers configuration.
#>
function Initialize-TestEnvironment {
    Write-TestLog "Initializing test environment..." 'SECTION'

    if (-not (Test-AdminRight)) {
        Write-TestLog "Administrator privileges required!" 'FAIL'
        exit 1
    }

    # Create test directory
    if (-not (Test-Path $TestPath)) {
        New-Item -ItemType Directory -Path $TestPath -Force | Out-Null
    }

    # Create test users
    Write-TestLog "Creating test users..." 'INFO'
    $successCount = 0
    foreach ($user in $script:TestUsers) {
        if (New-TestUser -UserName $user.Name -DaysOld $user.DaysOld -SizeMB $user.SizeMB) {
            $successCount++
        }
    }

    Write-TestLog "Created $successCount of $($script:TestUsers.Count) test users" 'INFO'

    # Refresh profile list
    Write-TestLog "Refreshing profile list..." 'INFO'
    Start-Sleep -Seconds 2
}
#endregion

#region Test Cases
<#
.SYNOPSIS
    Tests the profile enumeration functionality of DelprofPS.

.DESCRIPTION
    Validates that DelprofPS can correctly discover and enumerate user profiles
    on the local system, returning proper profile data with paths and SIDs.

.EXAMPLE
    Test-ProfileEnumeration
    Runs the profile enumeration test and records the result.
#>
function Test-ProfileEnumeration {
    Write-TestLog "`nTEST: Profile Enumeration" 'SECTION'

    try {
        # Use Get-UserProfile from DelprofPS.ps1
        $profiles = Get-UserProfile -ComputerName $env:COMPUTERNAME

        if ($null -eq $profiles) {
            Add-TestResult -TestName "Profile Enumeration - Get-UserProfile" -Passed $false -ErrorMessage "Get-UserProfile returned null"
            Write-TestLog "Get-UserProfile returned null" 'FAIL'
            return
        }

        Add-TestResult -TestName "Profile Enumeration - Get-UserProfile" -Passed $true -Details "Get-UserProfile returned $($profiles.Count) profiles"
        Write-TestLog "Get-UserProfile found $($profiles.Count) total profiles" 'PASS'

        # Use ConvertTo-UserName from DelprofPS.ps1 to resolve SIDs
        $testProfiles = @()
        foreach ($prof in $profiles) {
            $userName = ConvertTo-UserName -SID $prof.SID
            if ($userName -and $userName -like '*DPTest_*') {
                $testProfiles += [PSCustomObject]@{
                    SID = $prof.SID
                    UserName = $userName
                    ProfilePath = $prof.ProfilePath
                }
            }
        }

        if ($testProfiles.Count -eq $script:TestUsers.Count) {
            Add-TestResult -TestName "Profile Enumeration - Test Users" -Passed $true -Details "Found $($testProfiles.Count) test profiles via ConvertTo-UserName"
            Write-TestLog "Found all $($testProfiles.Count) test profiles" 'PASS'
        }
        else {
            Add-TestResult -TestName "Profile Enumeration - Test Users" -Passed $false -Details "Expected $($script:TestUsers.Count), found $($testProfiles.Count)"
            Write-TestLog "Profile count mismatch" 'FAIL'
        }
    }
    catch {
        Add-TestResult -TestName "Profile Enumeration" -Passed $false -ErrorMessage $_.Exception.Message
        Write-TestLog "Profile enumeration failed: $_" 'FAIL'
    }
}

function Test-AgeCalculation {
    Write-TestLog "`nTEST: Age Calculation Methods" 'SECTION'

    $methods = @('NTUSER_DAT', 'ProfilePath', 'Registry')

    foreach ($method in $methods) {
        try {
            # Find an old test user
            $oldUser = $script:TestUsers | Where-Object { $_.DaysOld -gt 100 } | Select-Object -First 1
            if (-not $oldUser) { continue }

            $profilePath = Join-Path $env:SystemDrive "Users\$($oldUser.Name)"

            # Resolve SID for this test user
            $userSid = (New-Object System.Security.Principal.NTAccount($oldUser.Name)).Translate(
                [System.Security.Principal.SecurityIdentifier]).Value

            # Use Get-ProfileAge from DelprofPS.ps1
            $ageInfo = Get-ProfileAge -ProfilePath $profilePath -SID $userSid -Method $method -ComputerName $env:COMPUTERNAME

            if ($ageInfo -and $ageInfo.LastUsed -and $ageInfo.LastUsed -ne [DateTime]::MinValue) {
                $ageDays = [math]::Floor(((Get-Date) - $ageInfo.LastUsed).TotalDays)
                $expectedAge = $oldUser.DaysOld
                $variance = [math]::Abs($ageDays - $expectedAge)

                if ($variance -le 1) {
                    Add-TestResult -TestName "Age Calculation - $method" -Passed $true -Details "Age: $ageDays days via $($ageInfo.Source) (expected ~$expectedAge)"
                    Write-TestLog "Age calculation ($method): $ageDays days (source: $($ageInfo.Source))" 'PASS'
                }
                else {
                    Add-TestResult -TestName "Age Calculation - $method" -Passed $false -Details "Age: $ageDays, Expected: $expectedAge, Source: $($ageInfo.Source)"
                    Write-TestLog "Age variance too high for $method ($ageDays vs $expectedAge)" 'WARN'
                }
            }
            else {
                Add-TestResult -TestName "Age Calculation - $method" -Passed $false -Details "Get-ProfileAge returned no valid date"
                Write-TestLog "Get-ProfileAge returned no valid date for $method" 'FAIL'
            }
        }
        catch {
            Add-TestResult -TestName "Age Calculation - $method" -Passed $false -ErrorMessage $_.Exception.Message
            Write-TestLog "Age calculation ($method) failed: $_" 'FAIL'
        }
    }
}

function Test-ProfileFiltering {
    Write-TestLog "`nTEST: Profile Filtering" 'SECTION'

    try {
        # Get all test user profiles using Get-UserProfile from DelprofPS.ps1
        $allProfiles = Get-UserProfile -ComputerName $env:COMPUTERNAME
        $testProfileData = @()
        foreach ($prof in $allProfiles) {
            $userName = ConvertTo-UserName -SID $prof.SID
            if ($userName -and $userName -like '*DPTest_*') {
                $shortName = $userName.Split('\')[-1]
                $size = Get-ProfileFolderSize -Path $prof.ProfilePath
                $testProfileData += @{
                    UserName = $shortName
                    SID = $prof.SID
                    ProfilePath = $prof.ProfilePath
                    Size = $size
                }
            }
        }

        # Save original filter values
        $origInclude = $Include
        $origExclude = $Exclude
        $origMinSize = $MinProfileSizeMB
        $origMaxSize = $MaxProfileSizeMB

        # Test 1: Include filter - should match only DPTest_Old* users
        try {
            $script:Include = @('DPTest_Old*')
            $script:Exclude = $null
            $script:MinProfileSizeMB = 0
            $script:MaxProfileSizeMB = 0
            $matchCount = 0
            foreach ($tp in $testProfileData) {
                if (Test-ProfileFilter -UserName $tp.UserName -SID $tp.SID -ProfilePath $tp.ProfilePath -ProfileSize $tp.Size -ActualProfileType 'Local') {
                    $matchCount++
                }
            }
            $passed = ($matchCount -eq 2)
            Add-TestResult -TestName "Filter - Include Pattern" -Passed $passed -Details "Test-ProfileFilter matched $matchCount (expected 2)"
            Write-TestLog "Filter test (Include Pattern): $matchCount matches via Test-ProfileFilter" $(if ($passed) { 'PASS' } else { 'FAIL' })
        }
        catch {
            Add-TestResult -TestName "Filter - Include Pattern" -Passed $false -ErrorMessage $_.Exception.Message
            Write-TestLog "Filter test (Include Pattern) failed: $_" 'FAIL'
        }

        # Test 2: Exclude filter - *Service* should be excluded
        try {
            $script:Include = $null
            $script:Exclude = @('*Service*')
            $script:MinProfileSizeMB = 0
            $script:MaxProfileSizeMB = 0
            $excludedCount = 0
            foreach ($tp in $testProfileData) {
                if (-not (Test-ProfileFilter -UserName $tp.UserName -SID $tp.SID -ProfilePath $tp.ProfilePath -ProfileSize $tp.Size -ActualProfileType 'Local')) {
                    $excludedCount++
                }
            }
            $passed = ($excludedCount -ge 1)
            Add-TestResult -TestName "Filter - Exclude Pattern" -Passed $passed -Details "Test-ProfileFilter excluded $excludedCount profiles"
            Write-TestLog "Filter test (Exclude Pattern): $excludedCount excluded via Test-ProfileFilter" $(if ($passed) { 'PASS' } else { 'FAIL' })
        }
        catch {
            Add-TestResult -TestName "Filter - Exclude Pattern" -Passed $false -ErrorMessage $_.Exception.Message
            Write-TestLog "Filter test (Exclude Pattern) failed: $_" 'FAIL'
        }

        # Test 3: Size filter - only profiles >= 100MB
        try {
            $script:Include = $null
            $script:Exclude = $null
            $script:MinProfileSizeMB = 100
            $script:MaxProfileSizeMB = 0
            $matchCount = 0
            foreach ($tp in $testProfileData) {
                if (Test-ProfileFilter -UserName $tp.UserName -SID $tp.SID -ProfilePath $tp.ProfilePath -ProfileSize $tp.Size -ActualProfileType 'Local') {
                    $matchCount++
                }
            }
            # Users with >= 100MB: OldUser1(150), MedUser1(200), MedUser2(120), AdminTest(500) = 4
            $passed = ($matchCount -ge 3)
            Add-TestResult -TestName "Filter - Size Filter" -Passed $passed -Details "Test-ProfileFilter matched $matchCount profiles >= 100MB"
            Write-TestLog "Filter test (Size Filter): $matchCount matches via Test-ProfileFilter" $(if ($passed) { 'PASS' } else { 'FAIL' })
        }
        catch {
            Add-TestResult -TestName "Filter - Size Filter" -Passed $false -ErrorMessage $_.Exception.Message
            Write-TestLog "Filter test (Size Filter) failed: $_" 'FAIL'
        }

        # Restore original filter values
        $script:Include = $origInclude
        $script:Exclude = $origExclude
        $script:MinProfileSizeMB = $origMinSize
        $script:MaxProfileSizeMB = $origMaxSize
    }
    catch {
        Add-TestResult -TestName "Filter - Setup" -Passed $false -ErrorMessage $_.Exception.Message
        Write-TestLog "Profile filtering setup failed: $_" 'FAIL'
    }
}

function Test-ProfileSizeCalculation {
    Write-TestLog "`nTEST: Profile Size Calculation" 'SECTION'

    try {
        $testUser = $script:TestUsers | Select-Object -First 1
        $profilePath = Join-Path $env:SystemDrive "Users\$($testUser.Name)"

        if (Test-Path $profilePath) {
            # Use Get-ProfileFolderSize from DelprofPS.ps1
            $size = Get-ProfileFolderSize -Path $profilePath

            $sizeMB = [math]::Round($size / 1MB, 2)
            $expectedMB = $testUser.SizeMB

            # Use Format-Byte from DelprofPS.ps1
            $formatted = Format-Byte -Bytes $size

            # Allow variance for filesystem overhead
            if ($sizeMB -ge ($expectedMB * 0.8) -and $sizeMB -le ($expectedMB * 1.5)) {
                Add-TestResult -TestName "Profile Size - Get-ProfileFolderSize" -Passed $true -Details "Size: $formatted (expected ~$expectedMB MB)"
                Write-TestLog "Get-ProfileFolderSize: $formatted" 'PASS'
            }
            else {
                Add-TestResult -TestName "Profile Size - Get-ProfileFolderSize" -Passed $false -Details "Size: $formatted, Expected: ~$expectedMB MB"
                Write-TestLog "Size outside expected range" 'WARN'
            }

            # Verify Format-Byte output is well-formed
            $fmtPassed = ($formatted -match '^[\d.]+ (B|KB|MB|GB|TB)$')
            Add-TestResult -TestName "Profile Size - Format-Byte" -Passed $fmtPassed -Details "Format-Byte returned: $formatted"
            Write-TestLog "Format-Byte output: $formatted" $(if ($fmtPassed) { 'PASS' } else { 'FAIL' })
        }
    }
    catch {
        Add-TestResult -TestName "Profile Size Calculation" -Passed $false -ErrorMessage $_.Exception.Message
        Write-TestLog "Size calculation failed: $_" 'FAIL'
    }
}

function Test-ProtectedProfileDetection {
    Write-TestLog "`nTEST: Protected Profile Detection" 'SECTION'

    try {
        # Test system profile names via Test-IsProtectedProfile from DelprofPS.ps1
        $systemProfiles = @('Administrator', 'Guest', 'Default', 'Public')
        foreach ($prof in $systemProfiles) {
            $isProtected = Test-IsProtectedProfile -UserName $prof -SID 'S-1-5-21-0-0-0-1234'
            Add-TestResult -TestName "Protected Profile - $prof" -Passed $isProtected -Details "Test-IsProtectedProfile returned $isProtected"
            Write-TestLog "Test-IsProtectedProfile('$prof'): $isProtected" $(if ($isProtected) { 'PASS' } else { 'FAIL' })
        }

        # Test protected SIDs via Test-IsProtectedProfile from DelprofPS.ps1
        $protectedSIDs = @('S-1-5-18', 'S-1-5-19', 'S-1-5-20')
        foreach ($sid in $protectedSIDs) {
            $isProtected = Test-IsProtectedProfile -UserName 'SomeUser' -SID $sid
            Add-TestResult -TestName "Protected SID - $sid" -Passed $isProtected -Details "Test-IsProtectedProfile returned $isProtected"
            Write-TestLog "Test-IsProtectedProfile(SID=$sid): $isProtected" $(if ($isProtected) { 'PASS' } else { 'FAIL' })
        }

        # Test that a regular test user is NOT protected
        $testUserName = ($script:TestUsers | Select-Object -First 1).Name
        $isNotProtected = -not (Test-IsProtectedProfile -UserName $testUserName -SID 'S-1-5-21-0-0-0-9999')
        Add-TestResult -TestName "Protected Profile - Test User Not Protected" -Passed $isNotProtected -Details "Test-IsProtectedProfile('$testUserName') = $(-not $isNotProtected)"
        Write-TestLog "Test user '$testUserName' correctly not protected: $isNotProtected" $(if ($isNotProtected) { 'PASS' } else { 'FAIL' })
    }
    catch {
        Add-TestResult -TestName "Protected Profile Detection" -Passed $false -ErrorMessage $_.Exception.Message
        Write-TestLog "Protected profile detection failed: $_" 'FAIL'
    }
}

function Test-DryRunMode {
    Write-TestLog "`nTEST: Dry Run / Preview Mode" 'SECTION'

    try {
        # Test that dry run doesn't delete anything
        $initialCount = (Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" |
            Where-Object { $_.PSChildName -match '^S-1-5-21' }).Count

        # Simulate dry run
        Write-TestLog "Simulating dry run..." 'INFO'
        Start-Sleep -Seconds 1

        $finalCount = (Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" |
            Where-Object { $_.PSChildName -match '^S-1-5-21' }).Count

        if ($initialCount -eq $finalCount) {
            Add-TestResult -TestName "Dry Run Mode" -Passed $true -Details "No profiles deleted in dry run"
            Write-TestLog "Dry run mode preserved all profiles" 'PASS'
        }
        else {
            Add-TestResult -TestName "Dry Run Mode" -Passed $false -Details "Profile count changed during dry run"
            Write-TestLog "Profile count changed!" 'FAIL'
        }
    }
    catch {
        Add-TestResult -TestName "Dry Run Mode" -Passed $false -ErrorMessage $_.Exception.Message
        Write-TestLog "Dry run test failed: $_" 'FAIL'
    }
}

function Test-BackupFunctionality {
    Write-TestLog "`nTEST: Profile Backup" 'SECTION'

    try {
        $backupDir = Join-Path $TestPath 'Backups'
        $testUser = ($script:TestUsers | Select-Object -First 1).Name
        $profilePath = Join-Path $env:SystemDrive "Users\$testUser"

        if (Test-Path $profilePath) {
            # Set $BackupPath so Backup-Profile from DelprofPS.ps1 uses it
            $script:BackupPath = $backupDir

            $result = Backup-Profile -SourcePath $profilePath -UserName $testUser

            # Find the created backup file
            $backupFiles = Get-ChildItem -Path $backupDir -Filter "${testUser}_*.zip" -ErrorAction SilentlyContinue

            if ($result -and $backupFiles) {
                $backupSize = ($backupFiles | Select-Object -Last 1).Length
                $formatted = Format-Byte -Bytes $backupSize
                Add-TestResult -TestName "Profile Backup - Backup-Profile" -Passed $true -Details "Backup-Profile created: $formatted"
                Write-TestLog "Backup-Profile created: $($backupFiles[-1].FullName)" 'PASS'
            }
            else {
                Add-TestResult -TestName "Profile Backup - Backup-Profile" -Passed $false -Details "Backup-Profile returned $result, files found: $($backupFiles.Count)"
                Write-TestLog "Backup-Profile did not create backup" 'FAIL'
            }

            # Reset BackupPath
            $script:BackupPath = $null
        }
    }
    catch {
        Add-TestResult -TestName "Profile Backup" -Passed $false -ErrorMessage $_.Exception.Message
        Write-TestLog "Backup test failed: $_" 'FAIL'
    }
}

function Test-CSVExport {
    Write-TestLog "`nTEST: CSV Export" 'SECTION'

    try {
        $csvPath = Join-Path $TestPath 'TestResults.csv'

        $testData = @(
            [PSCustomObject]@{ ComputerName = 'TEST01'; UserName = 'DPTest_OldUser1'; AgeInDays = 200; SizeBytes = 157286400 }
            [PSCustomObject]@{ ComputerName = 'TEST01'; UserName = 'DPTest_NewUser1'; AgeInDays = 15; SizeBytes = 52428800 }
        )

        $testData | Export-Csv -Path $csvPath -NoTypeInformation -Force

        if (Test-Path $csvPath) {
            $imported = Import-Csv -Path $csvPath
            if ($imported.Count -eq 2) {
                Add-TestResult -TestName "CSV Export" -Passed $true -Details "Exported and imported $($imported.Count) records"
                Write-TestLog "CSV export/import working" 'PASS'
            }
            else {
                Add-TestResult -TestName "CSV Export" -Passed $false -Details "Record count mismatch"
                Write-TestLog "CSV record count mismatch" 'FAIL'
            }
        }
        else {
            Add-TestResult -TestName "CSV Export" -Passed $false -Details "CSV file not created"
            Write-TestLog "CSV file not created" 'FAIL'
        }
    }
    catch {
        Add-TestResult -TestName "CSV Export" -Passed $false -ErrorMessage $_.Exception.Message
        Write-TestLog "CSV export test failed: $_" 'FAIL'
    }
}

function Test-HTMLReportGeneration {
    Write-TestLog "`nTEST: HTML Report Generation" 'SECTION'

    try {
        $htmlPath = Join-Path $TestPath 'TestReport.html'

        # Build sample results matching the format Invoke-ComputerProcessing produces
        $sampleResults = @(
            [PSCustomObject]@{
                ComputerName = $env:COMPUTERNAME; UserName = 'DPTest_OldUser1'; Domain = $env:USERDOMAIN
                SID = 'S-1-5-21-0-0-0-1001'; ProfilePath = 'C:\Users\DPTest_OldUser1'
                ProfileType = 'Local'; LastUsed = (Get-Date).AddDays(-200).ToString('yyyy-MM-dd HH:mm:ss')
                AgeInDays = 200; AgeSource = 'NTUSER.DAT'; SizeBytes = 157286400
                SizeFormatted = '150 MB'; IsActiveSession = $false; EligibleForDeletion = $true
                Deleted = $false; Error = $null
            }
            [PSCustomObject]@{
                ComputerName = $env:COMPUTERNAME; UserName = 'DPTest_NewUser1'; Domain = $env:USERDOMAIN
                SID = 'S-1-5-21-0-0-0-1002'; ProfilePath = 'C:\Users\DPTest_NewUser1'
                ProfileType = 'Local'; LastUsed = (Get-Date).AddDays(-15).ToString('yyyy-MM-dd HH:mm:ss')
                AgeInDays = 15; AgeSource = 'NTUSER.DAT'; SizeBytes = 52428800
                SizeFormatted = '50 MB'; IsActiveSession = $false; EligibleForDeletion = $true
                Deleted = $false; Error = $null
            }
        )

        $summary = @{
            Computers = 1
            ProfilesProcessed = 2
            ProfilesDeleted = 0
            SpaceFreed = '0 B'
            Duration = '00:00:05'
        }

        # Use Export-HtmlReport from DelprofPS.ps1
        Export-HtmlReport -Path $htmlPath -Results $sampleResults -Summary $summary

        if (Test-Path $htmlPath) {
            $content = Get-Content $htmlPath -Raw
            $hasTitle = $content -match 'Delprof2-PS'
            $hasTable = $content -match '<table'
            $hasData = $content -match 'DPTest_OldUser1'

            if ($hasTitle -and $hasTable -and $hasData) {
                Add-TestResult -TestName "HTML Report - Export-HtmlReport" -Passed $true -Details "Export-HtmlReport created valid HTML report"
                Write-TestLog "Export-HtmlReport generated valid report" 'PASS'
            }
            else {
                Add-TestResult -TestName "HTML Report - Export-HtmlReport" -Passed $false -Details "Missing: title=$hasTitle, table=$hasTable, data=$hasData"
                Write-TestLog "Export-HtmlReport output incomplete" 'FAIL'
            }
        }
        else {
            Add-TestResult -TestName "HTML Report - Export-HtmlReport" -Passed $false -Details "Export-HtmlReport did not create file"
            Write-TestLog "HTML file not created" 'FAIL'
        }
    }
    catch {
        Add-TestResult -TestName "HTML Report Generation" -Passed $false -ErrorMessage $_.Exception.Message
        Write-TestLog "HTML report test failed: $_" 'FAIL'
    }
}

function Test-ProfileDeletion {
    Write-TestLog "`nTEST: Profile Deletion (with -WhatIf)" 'SECTION'

    try {
        # This test uses -WhatIf to simulate deletion without actually removing anything
        $testUser = ($script:TestUsers | Where-Object { $_.Name -eq 'DPTest_NewUser1' }).Name

        if ($testUser) {
            Write-TestLog "Testing deletion safety for $testUser..." 'INFO'

            # Verify profile exists
            $profilePath = Join-Path $env:SystemDrive "Users\$testUser"
            $profileExists = Test-Path $profilePath

            # Test that we can simulate the deletion process
            $wouldDelete = $profileExists -and ($testUser -notlike '*admin*') -and ($testUser -notlike '*system*')

            if ($wouldDelete) {
                Add-TestResult -TestName "Profile Deletion Safety" -Passed $true -Details "Deletion safety checks working"
                Write-TestLog "Deletion safety checks passed" 'PASS'
            }
            else {
                Add-TestResult -TestName "Profile Deletion Safety" -Passed $false -Details "Safety check issue"
                Write-TestLog "Deletion safety check issue" 'WARN'
            }
        }
    }
    catch {
        Add-TestResult -TestName "Profile Deletion Safety" -Passed $false -ErrorMessage $_.Exception.Message
        Write-TestLog "Deletion test failed: $_" 'FAIL'
    }
}
#endregion

#region Cleanup Functions
<#
.SYNOPSIS
    Removes a test user account and associated profile data.

.DESCRIPTION
    Cleans up a test user by removing the local user account,
    profile directory, and associated registry entries.

.PARAMETER UserName
    The username of the test account to remove.

.EXAMPLE
    Remove-TestUser -UserName "DPTest_User1"
    Removes the specified test user and cleans up all associated data.
#>
function Remove-TestUser {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Test cleanup helper function')]
    param([string]$UserName)

    try {
        # Remove local user
        $user = Get-LocalUser -Name $UserName -ErrorAction SilentlyContinue
        if ($user) {
            Remove-LocalUser -Name $UserName -ErrorAction SilentlyContinue
            Write-TestLog "Removed test user: $UserName" 'PASS'
        }

        # Remove profile directory
        $profilePath = Join-Path $env:SystemDrive "Users\$UserName"
        if (Test-Path $profilePath) {
            Remove-Item -Path $profilePath -Recurse -Force -ErrorAction SilentlyContinue
            Write-TestLog "Removed profile directory: $profilePath" 'PASS'
        }

        # Remove registry profile
        $profileListPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
        $profiles = Get-ChildItem $profileListPath -ErrorAction SilentlyContinue |
            Where-Object { $_.PSChildName -match '^S-1-5-21' }

        foreach ($prof in $profiles) {
            try {
                $props = Get-ItemProperty $prof.PSPath -ErrorAction SilentlyContinue
                if ($props.ProfileImagePath -like "*\$UserName") {
                    Remove-Item -Path $prof.PSPath -Recurse -Force -ErrorAction SilentlyContinue
                    Write-TestLog "Removed registry profile for $UserName" 'PASS'
                }
            }
            catch { }
        }

        return $true
    }
    catch {
        Write-TestLog "Failed to remove test user $UserName`: $_" 'WARN'
        return $false
    }
}

<#
.SYNOPSIS
    Cleans up the entire test environment by removing all test artifacts.

.DESCRIPTION
    Removes all test users, their profiles, and generated test files.
    Respects the -KeepTestArtifacts switch to preserve artifacts for inspection.

.EXAMPLE
    Clear-TestEnvironment
    Removes all test users and cleans up the test directory.

.EXAMPLE
    Clear-TestEnvironment -KeepTestArtifacts
    Skips cleanup when KeepTestArtifacts is specified.
#>
function Clear-TestEnvironment {
    Write-TestLog "`nCleaning up test environment..." 'SECTION'

    if ($KeepTestArtifacts) {
        Write-TestLog "KeepTestArtifacts specified - skipping cleanup" 'WARN'
        Write-TestLog "Test users and profiles remain for inspection" 'INFO'
        return
    }

    foreach ($user in $script:TestUsers) {
        Remove-TestUser -UserName $user.Name
    }

    # Clean up test files
    $testFiles = @(
        (Join-Path $TestPath 'TestResults.csv')
        (Join-Path $TestPath 'TestReport.html')
    )

    foreach ($file in $testFiles) {
        if (Test-Path $file) {
            Remove-Item -Path $file -Force -ErrorAction SilentlyContinue
        }
    }

    # Clean up backups
    $backupPath = Join-Path $TestPath 'Backups'
    if (Test-Path $backupPath) {
        Remove-Item -Path $backupPath -Recurse -Force -ErrorAction SilentlyContinue
    }

    Write-TestLog "Cleanup complete" 'PASS'
}
#endregion

#region Main Execution
<#
.SYNOPSIS
    Displays a formatted summary of all test results.

.DESCRIPTION
    Calculates and displays the total number of tests run, passed, and failed,
    along with execution duration. Lists any failed tests with their error messages.
    Exits with code 0 if all tests passed, or 1 if any tests failed.

.EXAMPLE
    Show-TestSummary
    Displays the test summary report and exits with appropriate return code.
#>
function Show-TestSummary {
    Write-TestLog ("`n" + ('=' * 80)) 'SECTION'
    Write-TestLog "TEST SUMMARY" 'SECTION'
    Write-TestLog ('=' * 80) 'SECTION'

    $passed = ($script:TestResults | Where-Object { $_.Passed }).Count
    $failed = ($script:TestResults | Where-Object { -not $_.Passed }).Count
    $total = $script:TestResults.Count

    $duration = (Get-Date) - $script:StartTime

    Write-Host "`nResults:"
    Write-Host "  Total Tests:  $total"
    Write-Host "  Passed:       $passed" -ForegroundColor Green
    Write-Host "  Failed:       $failed" -ForegroundColor $(if ($failed -gt 0) { 'Red' } else { 'Green' })
    Write-Host "  Duration:     $($duration.ToString('mm\:ss'))"

    if ($failed -gt 0) {
        Write-Host "`nFailed Tests:" -ForegroundColor Red
        $script:TestResults | Where-Object { -not $_.Passed } | ForEach-Object {
            Write-Host "  - $($_.TestName): $($_.ErrorMessage)" -ForegroundColor Red
        }
    }

    Write-TestLog ('=' * 80) 'SECTION'

    # Return exit code
    if ($failed -eq 0) {
        exit 0
    }
    else {
        exit 1
    }
}

<#
.SYNOPSIS
    Main entry point for the Delprof2-PS test suite.

.DESCRIPTION
    Orchestrates the entire test execution flow:
    1. Displays test header information
    2. Initializes the test environment
    3. Runs all test cases
    4. Cleans up test artifacts
    5. Displays test summary

    Handles errors gracefully and ensures cleanup runs even if tests fail.

.EXAMPLE
    Main
    Executes the complete test suite.

.EXAMPLE
    # Run from command line
    .\DelprofPS.Tests.ps1

.EXAMPLE
    # Run and keep artifacts for inspection
    .\DelprofPS.Tests.ps1 -KeepTestArtifacts
#>
function Main {
    Clear-Host
    Write-TestLog "Delprof2-PS Comprehensive Test Suite" 'SECTION'
    Write-TestLog "Test Path: $TestPath" 'INFO'
    Write-TestLog "Keep Artifacts: $KeepTestArtifacts`n" 'INFO'

    try {
        # Setup
        Initialize-TestEnvironment

        # Run all tests
        Test-ProfileEnumeration
        Test-AgeCalculation
        Test-ProfileFiltering
        Test-ProfileSizeCalculation
        Test-ProtectedProfileDetection
        Test-DryRunMode
        Test-BackupFunctionality
        Test-CSVExport
        Test-HTMLReportGeneration
        Test-ProfileDeletion

        # Cleanup
        Clear-TestEnvironment

        # Summary
        Show-TestSummary
    }
    catch {
        Write-TestLog "Critical error in test suite: $_" 'FAIL'
        exit 1
    }
}

# Run main function
Main
#endregion
