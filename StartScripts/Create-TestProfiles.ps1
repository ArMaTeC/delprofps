<#
.SYNOPSIS
    Creates local test user profiles for testing DelprofPS.

.DESCRIPTION
    Creates multiple local user accounts with varying profile ages, sizes, and types
    to exercise DelprofPS filtering, age calculation, and removal logic.

    Profiles created:
      - TestUser_Old1..3    : Inactive 60-90 days (should be caught by default 30-day filter)
      - TestUser_Recent1..2 : Active within last 7 days (should be skipped)
      - TestUser_Large1     : Profile with ~50 MB dummy data (tests size display/filtering)
      - TestUser_Small1     : Profile with minimal data
      - TestUser_Admin1     : Matches *admin* exclude pattern (should be excluded)
      - TestUser_Service1   : Matches *service* exclude pattern (should be excluded)

    Run with -Remove to clean up all test profiles created by this script.

.PARAMETER Remove
    Remove all test profiles and local accounts created by this script.

.PARAMETER Password
    Password for the test accounts. Defaults to 'T3stP@ss!2026'.

.EXAMPLE
    # Create test profiles (requires elevation)
    .\Create-TestProfiles.ps1

.EXAMPLE
    # Remove all test profiles
    .\Create-TestProfiles.ps1 -Remove

.NOTES
    Requires Administrator privileges.
    Only creates LOCAL accounts — safe for domain-joined machines.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$Remove,
    [string]$Password = 'T3stP@ss!2026'
)

#Requires -RunAsAdministrator

# ── Configuration ──────────────────────────────────────────────────────────────
$testUsers = @(
    @{ Name = 'TestUser_Old1';     DaysOld = 90;  SizeMB = 5;  Desc = 'Old profile (90 days)' }
    @{ Name = 'TestUser_Old2';     DaysOld = 60;  SizeMB = 3;  Desc = 'Old profile (60 days)' }
    @{ Name = 'TestUser_Old3';     DaysOld = 45;  SizeMB = 2;  Desc = 'Old profile (45 days)' }
    @{ Name = 'TestUser_Recent1';  DaysOld = 3;   SizeMB = 10; Desc = 'Recent profile (3 days)' }
    @{ Name = 'TestUser_Recent2';  DaysOld = 0;   SizeMB = 1;  Desc = 'Active today' }
    @{ Name = 'TestUser_Large1';   DaysOld = 40;  SizeMB = 50; Desc = 'Large old profile (50 MB)' }
    @{ Name = 'TestUser_Small1';   DaysOld = 35;  SizeMB = 0;  Desc = 'Minimal old profile' }
    @{ Name = 'TestUser_Admin1';   DaysOld = 120; SizeMB = 2;  Desc = 'Matches *admin* exclude' }
    @{ Name = 'TestUser_Service1'; DaysOld = 120; SizeMB = 1;  Desc = 'Matches *service* exclude' }
)

$securePass = ConvertTo-SecureString $Password -AsPlainText -Force

# ── Remove Mode ────────────────────────────────────────────────────────────────
if ($Remove) {
    Write-Host "`n=== Removing Test Profiles ===" -ForegroundColor Yellow
    foreach ($user in $testUsers) {
        $name = $user.Name
        Write-Host "  [$name] " -NoNewline

        # Remove local account (also removes profile via -DeleteProfile in newer Windows)
        try {
            $account = Get-LocalUser -Name $name -ErrorAction SilentlyContinue
            if ($account) {
                # Get profile path from registry
                $sid = (New-Object System.Security.Principal.NTAccount($name)).Translate(
                    [System.Security.Principal.SecurityIdentifier]).Value
                $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid"
                $profilePath = $null
                if (Test-Path $regPath) {
                    $profilePath = (Get-ItemProperty $regPath -ErrorAction SilentlyContinue).ProfileImagePath
                }

                Remove-LocalUser -Name $name -ErrorAction Stop
                Write-Host "account removed" -ForegroundColor Green -NoNewline

                # Clean up profile folder if it still exists
                if ($profilePath -and (Test-Path $profilePath)) {
                    Remove-Item $profilePath -Recurse -Force -ErrorAction SilentlyContinue
                    Write-Host ", folder deleted" -ForegroundColor Green
                }
                else {
                    Write-Host "" 
                }

                # Clean up orphaned registry key
                if (Test-Path $regPath) {
                    Remove-Item $regPath -Force -ErrorAction SilentlyContinue
                }
            }
            else {
                Write-Host "not found (skipped)" -ForegroundColor DarkGray
            }
        }
        catch {
            Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    Write-Host "`nDone.`n" -ForegroundColor Cyan
    return
}

# ── Create Mode ────────────────────────────────────────────────────────────────
Write-Host "`n=== Creating Test Profiles for DelprofPS ===" -ForegroundColor Cyan
Write-Host "Password: $Password"
Write-Host "Profiles: $($testUsers.Count)`n"

foreach ($user in $testUsers) {
    $name = $user.Name
    $daysOld = $user.DaysOld
    $sizeMB = $user.SizeMB
    $desc = $user.Desc

    Write-Host "[$name] $desc" -ForegroundColor White

    # 1. Create local user account
    $existing = Get-LocalUser -Name $name -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  Account already exists - skipping creation" -ForegroundColor DarkGray
    }
    else {
        try {
            New-LocalUser -Name $name -Password $securePass -Description "DelprofPS test: $desc" `
                -PasswordNeverExpires -UserMayNotChangePassword -ErrorAction Stop | Out-Null
            Write-Host "  Account created" -ForegroundColor Green
        }
        catch {
            Write-Host "  FAILED to create account: $($_.Exception.Message)" -ForegroundColor Red
            continue
        }
    }

    # 2. Force profile creation by running a process as the user
    #    (This creates the profile folder, NTUSER.DAT, etc.)
    $profilePath = "C:\Users\$name"
    if (-not (Test-Path $profilePath)) {
        Write-Host "  Creating profile (first logon simulation)..." -NoNewline
        try {
            # Start a process as the user to trigger profile creation
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = "$env:SystemRoot\System32\cmd.exe"
            $psi.Arguments = '/c echo profile_init'
            $psi.UserName = $name
            $psi.Domain = $env:COMPUTERNAME
            $psi.Password = $securePass
            $psi.UseShellExecute = $false
            $psi.LoadUserProfile = $true
            $psi.CreateNoWindow = $true
            $psi.WorkingDirectory = "$env:SystemRoot\Temp"
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true

            $proc = [System.Diagnostics.Process]::Start($psi)
            $proc.WaitForExit(15000) | Out-Null
            if (-not $proc.HasExited) { $proc.Kill() }
            
            # Wait briefly for profile folder to appear
            Start-Sleep -Seconds 2
            
            if (Test-Path $profilePath) {
                Write-Host " OK" -ForegroundColor Green
            }
            else {
                Write-Host " profile folder not created (may need manual logon)" -ForegroundColor Yellow
                continue
            }
        }
        catch {
            Write-Host " ERROR: $($_.Exception.Message)" -ForegroundColor Red
            continue
        }
    }
    else {
        Write-Host "  Profile folder exists" -ForegroundColor DarkGray
    }

    # 3. Populate profile with dummy data to reach target size
    if ($sizeMB -gt 0) {
        $dataDir = Join-Path $profilePath 'Documents\TestData'
        if (-not (Test-Path $dataDir)) { New-Item -Path $dataDir -ItemType Directory -Force | Out-Null }
        
        $currentSize = 0
        if (Test-Path $dataDir) {
            $currentSize = [math]::Round(((Get-ChildItem $dataDir -Recurse -File -ErrorAction SilentlyContinue | 
                Measure-Object -Property Length -Sum).Sum / 1MB), 1)
        }
        
        $remainingMB = $sizeMB - $currentSize
        if ($remainingMB -gt 0.5) {
            Write-Host "  Adding ~${remainingMB} MB of test data..." -NoNewline
            $chunkSizeMB = [math]::Min($remainingMB, 10)
            $chunks = [math]::Ceiling($remainingMB / $chunkSizeMB)
            for ($i = 1; $i -le $chunks; $i++) {
                $thisChunk = [math]::Min($chunkSizeMB, $remainingMB - (($i - 1) * $chunkSizeMB))
                $filePath = Join-Path $dataDir "testdata_$i.bin"
                if (-not (Test-Path $filePath)) {
                    $bytes = [byte[]]::new([int]($thisChunk * 1MB))
                    [System.Random]::new().NextBytes($bytes)
                    [System.IO.File]::WriteAllBytes($filePath, $bytes)
                }
            }
            Write-Host " done" -ForegroundColor Green
        }
        else {
            Write-Host "  Data already at target size" -ForegroundColor DarkGray
        }
    }

    # 4. Set NTUSER.DAT and profile folder timestamps to simulate age
    if ($daysOld -gt 0) {
        $targetDate = (Get-Date).AddDays(-$daysOld)
        Write-Host "  Setting timestamps to $($targetDate.ToString('yyyy-MM-dd'))..." -NoNewline
        try {
            # Set profile folder timestamps
            $dir = Get-Item $profilePath -Force
            $dir.LastWriteTime = $targetDate
            $dir.LastAccessTime = $targetDate

            # Set NTUSER.DAT timestamp (this is what NTUSER_DAT age method reads)
            $ntuser = Join-Path $profilePath 'NTUSER.DAT'
            if (Test-Path $ntuser) {
                $f = Get-Item $ntuser -Force
                $f.LastWriteTime = $targetDate
                $f.LastAccessTime = $targetDate
            }

            # Also set ntuser.dat.LOG files if present
            Get-ChildItem $profilePath -Filter 'ntuser*' -Force -ErrorAction SilentlyContinue | ForEach-Object {
                $_.LastWriteTime = $targetDate
                $_.LastAccessTime = $targetDate
            }

            Write-Host " done" -ForegroundColor Green
        }
        catch {
            Write-Host " ERROR: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# ── Summary ────────────────────────────────────────────────────────────────────
Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host "Created $($testUsers.Count) test profiles:`n"
Write-Host ("{0,-22} {1,-8} {2,-8} {3}" -f 'Username', 'Age', 'Size', 'Expected DelprofPS Behavior')
Write-Host ("{0,-22} {1,-8} {2,-8} {3}" -f '--------', '---', '----', '----------------------------')
foreach ($user in $testUsers) {
    $age = "$($user.DaysOld)d"
    $size = if ($user.SizeMB -gt 0) { "$($user.SizeMB) MB" } else { "~0 MB" }
    $behavior = if ($user.Name -match 'admin|service') {
        "EXCLUDED (pattern match)"
    }
    elseif ($user.DaysOld -ge 30) {
        "WOULD DELETE (inactive > 30d)"
    }
    else {
        "KEPT (active within 30d)"
    }
    Write-Host ("{0,-22} {1,-8} {2,-8} {3}" -f $user.Name, $age, $size, $behavior)
}

Write-Host "`nTest with DelprofPS:" -ForegroundColor Yellow
Write-Host "  Preview:  .\DelprofPS.ps1 -Preview -DaysInactive 30 -Exclude 'Administrator*,*admin*,*service*' -ShowSpace"
Write-Host "  GUI:      .\DelprofPS.ps1 -UI"
Write-Host "  Cleanup:  .\StartScripts\Create-TestProfiles.ps1 -Remove`n"
