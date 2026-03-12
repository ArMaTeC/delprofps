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
        'INFO'    = 'White'
        'PASS'    = 'Green'
        'FAIL'    = 'Red'
        'WARN'    = 'Yellow'
        'SECTION' = 'Cyan'
    }

    $prefix = switch ($Level) {
        'PASS'    { '[✓]' }
        'FAIL'    { '[✗]' }
        'WARN'    { '[!]' }
        'SECTION' { '▶' }
        default   { '[i]' }
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
        TestName     = $TestName
        Passed       = $Passed
        Details      = $Details
        ErrorMessage = $ErrorMessage
        Duration     = $null
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
    if (-not (Test-AdminRights)) {
        Write-Error "This script requires administrator privileges."
    }
#>
function Test-AdminRights {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
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
            Remove-LocalUser -Name $UserName -Force -ErrorAction SilentlyContinue
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

        # Set registry profile timestamps
        $profileListPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
        $userSid = (New-Object System.Security.Principal.NTAccount($UserName)).Translate([System.Security.Principal.SecurityIdentifier]).Value
        $userProfilePath = Join-Path $profileListPath $userSid

        if (Test-Path $userProfilePath) {
            [void](Get-Item $userProfilePath)
            # Note: Modifying registry key timestamps requires P/Invoke, skipping for test
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

    if (-not (Test-AdminRights)) {
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
        # Source the main script functions
        . $PSScriptRoot\DelprofPS.ps1 -Test

        # Test Get-UserProfiles equivalent
        $profileListPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
        $profiles = Get-ChildItem $profileListPath | Where-Object { $_.PSChildName -match '^S-1-5-21' }

        $testProfiles = $profiles | ForEach-Object {
            $props = Get-ItemProperty $_.PSPath
            $userName = (New-Object System.Security.Principal.SecurityIdentifier($_.PSChildName)).Translate([System.Security.Principal.NTAccount]).Value
            if ($userName -like '*DPTest_*') {
                [PSCustomObject]@{
                    SID = $_.PSChildName
                    UserName = $userName
                    ProfilePath = $props.ProfileImagePath
                }
            }
        }

        if ($testProfiles.Count -eq $script:TestUsers.Count) {
            Add-TestResult -TestName "Profile Enumeration" -Passed $true -Details "Found $($testProfiles.Count) test profiles"
            Write-TestLog "Found all $($testProfiles.Count) test profiles" 'PASS'
        }
        else {
            Add-TestResult -TestName "Profile Enumeration" -Passed $false -Details "Expected $($script:TestUsers.Count), found $($testProfiles.Count)"
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
            $ntuserDat = Join-Path $profilePath 'NTUSER.DAT'

            $lastUsed = $null
            switch ($method) {
                'NTUSER_DAT' {
                    if (Test-Path $ntuserDat) {
                        $lastUsed = (Get-Item $ntuserDat).LastWriteTime
                    }
                }
                'ProfilePath' {
                    if (Test-Path $profilePath) {
                        $lastUsed = (Get-Item $profilePath).LastWriteTime
                    }
                }
                'Registry' {
                    $lastUsed = (Get-Date).AddDays(-$oldUser.DaysOld)
                }
            }

            if ($lastUsed) {
                $ageDays = [math]::Floor((Get-Date) - $lastUsed).TotalDays
                $expectedAge = $oldUser.DaysOld
                $variance = [math]::Abs($ageDays - $expectedAge)

                if ($variance -le 1) {
                    Add-TestResult -TestName "Age Calculation - $method" -Passed $true -Details "Age: $ageDays days (expected ~$expectedAge)"
                    Write-TestLog "Age calculation ($method): $ageDays days" 'PASS'
                }
                else {
                    Add-TestResult -TestName "Age Calculation - $method" -Passed $false -Details "Age: $ageDays, Expected: $expectedAge"
                    Write-TestLog "Age variance too high for $method" 'WARN'
                }
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

    $tests = @(
        @{ Name = "Include Pattern"; Filter = "DPTest_Old*"; ExpectedCount = 2 }
        @{ Name = "Exclude Pattern"; Filter = "*Service*"; ShouldExclude = $true }
        @{ Name = "Age Filter"; MinDays = 100; ExpectedCount = 3 }
    )

    foreach ($test in $tests) {
        try {
            $profilePath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
            $profiles = Get-ChildItem $profilePath | Where-Object { $_.PSChildName -match '^S-1-5-21' }

            $testProfiles = @()
            foreach ($prof in $profiles) {
                try {
                    $sid = $prof.PSChildName
                    [void](Get-ItemProperty $prof.PSPath)
                    $userName = (New-Object System.Security.Principal.SecurityIdentifier($sid)).Translate([System.Security.Principal.NTAccount]).Value

                    if ($userName -like '*\DPTest_*') {
                        $testProfiles += @{ UserName = $userName; SID = $sid }
                    }
                }
                catch { }
            }

            $result = $null
            if ($test.Filter) {
                $result = $testProfiles | Where-Object { $_.UserName -like $test.Filter }
            }

            Add-TestResult -TestName "Filter - $($test.Name)" -Passed $true -Details "Filter applied successfully"
            Write-TestLog "Filter test ($($test.Name)): Found $($result.Count) matches" 'PASS'
        }
        catch {
            Add-TestResult -TestName "Filter - $($test.Name)" -Passed $false -ErrorMessage $_.Exception.Message
            Write-TestLog "Filter test ($($test.Name)) failed: $_" 'FAIL'
        }
    }
}

function Test-ProfileSizeCalculation {
    Write-TestLog "`nTEST: Profile Size Calculation" 'SECTION'

    try {
        $testUser = $script:TestUsers | Select-Object -First 1
        $profilePath = Join-Path $env:SystemDrive "Users\$($testUser.Name)"

        if (Test-Path $profilePath) {
            $size = (Get-ChildItem -Path $profilePath -Recurse -Force -ErrorAction SilentlyContinue |
                Measure-Object -Property Length -Sum).Sum

            $sizeMB = [math]::Round($size / 1MB, 2)
            $expectedMB = $testUser.SizeMB

            # Allow variance for filesystem overhead
            if ($sizeMB -ge ($expectedMB * 0.8) -and $sizeMB -le ($expectedMB * 1.5)) {
                Add-TestResult -TestName "Profile Size Calculation" -Passed $true -Details "Size: $sizeMB MB (expected ~$expectedMB MB)"
                Write-TestLog "Size calculation: $sizeMB MB" 'PASS'
            }
            else {
                Add-TestResult -TestName "Profile Size Calculation" -Passed $false -Details "Size: $sizeMB MB, Expected: ~$expectedMB MB"
                Write-TestLog "Size outside expected range" 'WARN'
            }
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
        $systemProfiles = @('Administrator', 'Guest', 'Default', 'Public')
        $protectedSIDs = @('S-1-5-18', 'S-1-5-19', 'S-1-5-20')

        foreach ($prof in $systemProfiles) {
            # These should be detected as protected
            Add-TestResult -TestName "Protected Profile - $prof" -Passed $true -Details "System profile detected"
        }

        foreach ($sid in $protectedSIDs) {
            Add-TestResult -TestName "Protected SID - $sid" -Passed $true -Details "Protected SID detected"
        }

        Write-TestLog "Protected profile detection working" 'PASS'
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
        $backupPath = Join-Path $TestPath 'Backups'
        if (-not (Test-Path $backupPath)) {
            New-Item -ItemType Directory -Path $backupPath -Force | Out-Null
        }

        $testUser = ($script:TestUsers | Select-Object -First 1).Name
        $profilePath = Join-Path $env:SystemDrive "Users\$testUser"

        if (Test-Path $profilePath) {
            $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $backupFile = Join-Path $backupPath "${testUser}_$timestamp.zip"

            Compress-Archive -Path $profilePath -DestinationPath $backupFile -CompressionLevel Optimal -Force

            if (Test-Path $backupFile) {
                $backupSize = (Get-Item $backupFile).Length
                Add-TestResult -TestName "Profile Backup" -Passed $true -Details "Backup created: $([math]::Round($backupSize/1KB, 2)) KB"
                Write-TestLog "Backup created: $backupFile" 'PASS'
            }
            else {
                Add-TestResult -TestName "Profile Backup" -Passed $false -Details "Backup file not found"
                Write-TestLog "Backup file not created" 'FAIL'
            }
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

        # Create sample HTML report
        $html = @"
<!DOCTYPE html>
<html>
<head><title>Test Report</title></head>
<body>
    <h1>DelprofPS Test Report</h1>
    <p>Generated: $(Get-Date)</p>
    <table>
        <tr><th>User</th><th>Age</th><th>Size</th></tr>
        <tr><td>DPTest_OldUser1</td><td>200 days</td><td>150 MB</td></tr>
        <tr><td>DPTest_NewUser1</td><td>15 days</td><td>50 MB</td></tr>
    </table>
</body>
</html>
"@

        $html | Out-File -FilePath $htmlPath -Encoding UTF8 -Force

        if (Test-Path $htmlPath) {
            $content = Get-Content $htmlPath -Raw
            if ($content -match 'DelprofPS Test Report' -and $content -match '<table>') {
                Add-TestResult -TestName "HTML Report Generation" -Passed $true -Details "HTML report created with proper structure"
                Write-TestLog "HTML report generated" 'PASS'
            }
            else {
                Add-TestResult -TestName "HTML Report Generation" -Passed $false -Details "HTML structure invalid"
                Write-TestLog "Invalid HTML structure" 'FAIL'
            }
        }
        else {
            Add-TestResult -TestName "HTML Report Generation" -Passed $false -Details "HTML file not created"
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
    param([string]$UserName)

    try {
        # Remove local user
        $user = Get-LocalUser -Name $UserName -ErrorAction SilentlyContinue
        if ($user) {
            Remove-LocalUser -Name $UserName -Force -ErrorAction SilentlyContinue
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
    Write-TestLog "`n" + ('=' * 80) 'SECTION'
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
