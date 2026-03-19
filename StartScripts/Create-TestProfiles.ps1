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
    Only creates LOCAL accounts - safe for domain-joined machines.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$Remove,
    [SecureString]$Password = (ConvertTo-SecureString 'T3stP@ss!2026' -AsPlainText -Force)
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

$securePass = $Password

# ── Remove Mode ────────────────────────────────────────────────────────────────
if ($Remove) {
    Write-Host "`n=== Removing Test Profiles ===" -ForegroundColor Yellow
    foreach ($user in $testUsers) {
        $name = $user.Name
        Write-Host "  [$name] " -NoNewline

        try {
            $account = Get-LocalUser -Name $name -ErrorAction SilentlyContinue
            if (-not $account) {
                Write-Host "not found (skipped)" -ForegroundColor DarkGray
                continue
            }
            
            # Resolve SID for profile lookup
            $sid = $null
            try {
                $sid = (New-Object System.Security.Principal.NTAccount($name)).Translate(
                    [System.Security.Principal.SecurityIdentifier]).Value
            } catch {}
            
            # Remove the local user account first
            Remove-LocalUser -Name $name -ErrorAction Stop
            Write-Host "account removed" -ForegroundColor Green -NoNewline
            
            # Use Win32_UserProfile to properly delete the profile (folder + registry)
            $profileDeleted = $false
            if ($sid) {
                $userProfile = Get-CimInstance -ClassName Win32_UserProfile -Filter "SID='$sid'" -ErrorAction SilentlyContinue
                if ($userProfile) {
                    try {
                        # Unload hive first if loaded
                        $hivePath = "Registry::HKEY_USERS\$sid"
                        if (Test-Path $hivePath) {
                            [gc]::Collect(); [gc]::WaitForPendingFinalizers()
                            $null = & reg.exe unload "HKU\$sid" 2>&1
                        }
                        Remove-CimInstance -InputObject $userProfile -ErrorAction Stop
                        $profileDeleted = $true
                        Write-Host ", profile deleted (WMI)" -ForegroundColor Green
                    }
                    catch {
                        Write-Host ", WMI delete failed: $($_.Exception.Message)" -ForegroundColor Yellow -NoNewline
                    }
                }
            }
            
            # Fallback: force-remove folder and registry key if WMI didn't handle it
            $profilePath = "C:\Users\$name"
            if (-not $profileDeleted -and (Test-Path $profilePath)) {
                # Unload hive if still loaded
                if ($sid) {
                    $hivePath = "Registry::HKEY_USERS\$sid"
                    if (Test-Path $hivePath) {
                        [gc]::Collect(); [gc]::WaitForPendingFinalizers()
                        $null = & reg.exe unload "HKU\$sid" 2>&1
                        Start-Sleep -Milliseconds 500
                    }
                }
                Remove-Item $profilePath -Recurse -Force -ErrorAction Stop
                Write-Host ", folder deleted (fallback)" -ForegroundColor Green -NoNewline
            }
            
            # Clean up orphaned registry key
            if ($sid) {
                $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid"
                if (Test-Path $regPath) {
                    Remove-Item $regPath -Force -ErrorAction SilentlyContinue
                    Write-Host ", registry cleaned" -ForegroundColor Green -NoNewline
                }
            }
            
            Write-Host ""
        }
        catch {
            Write-Host " ERROR: $($_.Exception.Message)" -ForegroundColor Red
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
    #    (This creates the profile folder, NTUSER.DAT, registry entry, etc.)
    $profilePath = "C:\Users\$name"
    
    # Check if profile is registered in the registry (not just folder existing)
    $needsLogon = $false
    if (-not (Test-Path $profilePath)) {
        $needsLogon = $true
    }
    else {
        # Folder exists - check if profile is actually registered in the registry
        $sid = $null
        try {
            $sid = (New-Object System.Security.Principal.NTAccount($name)).Translate(
                [System.Security.Principal.SecurityIdentifier]).Value
        } catch {}
        $regKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid"
        if (-not $sid -or -not (Test-Path $regKey)) {
            Write-Host "  Profile folder exists but NOT registered in registry - re-creating..." -ForegroundColor Yellow
            Remove-Item $profilePath -Recurse -Force -ErrorAction SilentlyContinue
            $needsLogon = $true
        }
        else {
            Write-Host "  Profile registered OK" -ForegroundColor DarkGray
        }
    }
    
    if ($needsLogon) {
        Write-Host "  Creating profile (first logon simulation)..." -NoNewline
        try {
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

    # 4. Unload any loaded registry hives for this user
    #    (Prevents WMI Win32_LoggedOnUser from reporting them as "active sessions")
    $sid = $null
    try {
        $sid = (New-Object System.Security.Principal.NTAccount($name)).Translate(
            [System.Security.Principal.SecurityIdentifier]).Value
    } catch {}
    if ($sid) {
        # Check if hive is loaded under HKU
        $hivePath = "Registry::HKEY_USERS\$sid"
        if (Test-Path $hivePath) {
            Write-Host "  Unloading registry hive..." -NoNewline
            # Close any handles then unload
            [gc]::Collect()
            [gc]::WaitForPendingFinalizers()
            $null = & reg.exe unload "HKU\$sid" 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Host " done" -ForegroundColor Green
            }
            else {
                Write-Host " skipped (hive in use - will clear on reboot)" -ForegroundColor DarkGray
            }
        }
    }

    # 5. Set NTUSER.DAT and profile folder timestamps to simulate age
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
