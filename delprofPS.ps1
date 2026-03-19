#Requires -Version 5.1
<#
.SYNOPSIS
    Delprof2-PS v2.0 - Enterprise-grade PowerShell replacement for Delprof2 with advanced features.

.DESCRIPTION
    A comprehensive user profile management tool that exceeds Delprof2 capabilities:
    
    CORE FEATURES:
    - Local and remote computer profile management
    - Multiple age calculation methods (NTUSER.DAT, ProfilePath, Registry, LastLogon)
    - Active session detection (prevents deleting logged-in users)
    - Flexible filtering (include/exclude by username, SID, or pattern)
    - Profile state detection (local, roaming, temporary, mandatory, corrupted)
    - Disk space calculation and reporting
    - Registry hive unloading before deletion
    - Comprehensive logging and CSV export
    
    ENTERPRISE FEATURES:
    - Interactive profile selection mode with visual menu
    - Test/Validate mode to check prerequisites without changes
    - Parallel processing for multiple computers
    - HTML report generation with professional styling
    - Windows Event Log integration for audit trails
    - Profile backup before deletion
    - JSON configuration file support
    - Email notifications for scheduled runs
    - Progress bars for long operations
    - Age-based color coding in output

.PARAMETER ComputerName
    Target computer(s). Defaults to localhost. Accepts multiple, pipeline input, and CSV files.

.PARAMETER DaysInactive
    Minimum days of inactivity for profile deletion. Default: 30

.PARAMETER AgeCalculation
    Method to determine profile age: NTUSER_DAT (default), ProfilePath, Registry, LastLogon, LastLogoff

.PARAMETER Include
    Wildcard pattern for usernames to include (e.g., "user*", "*admin*")

.PARAMETER Exclude
    Wildcard pattern for usernames to exclude (e.g., "Administrator*", "*service*")

.PARAMETER Delete
    Actually delete profiles. Without this, runs in dry-run/list mode.

.PARAMETER Force
    Skip confirmation prompts and ignore non-critical errors.

.PARAMETER UI
    Launches the graphical user interface for visual profile management.
    Provides a modern WPF-based GUI with access to all script functionality.

.PARAMETER Preview
    Shows what would be deleted without actually deleting. Performs a dry run
    that displays all profiles that match the criteria and would be removed.
    Combine with -Delete to see preview before actual deletion.

.PARAMETER IgnoreActiveSessions
    Allow deletion of profiles with active user sessions (DANGEROUS).

.PARAMETER UnloadHives
    Unload loaded registry hives before deletion (recommended).

.PARAMETER MaxRetries
    Number of retry attempts for locked files. Default: 3

.PARAMETER RetryDelaySeconds
    Seconds between retry attempts. Default: 2

.PARAMETER OutputPath
    Export results to CSV file.

.PARAMETER LogPath
    Write detailed log to file.

.PARAMETER Quiet
    Suppress console output (useful for scheduled tasks).

.PARAMETER ShowSpace
    Display disk space used by each profile.

.PARAMETER IncludeSystemProfiles
    Include system profiles (Default, Public, etc.) - EXTREME CAUTION.

.PARAMETER IncludeSpecialProfiles
    Include special accounts (SYSTEM, NetworkService, LocalService).

.PARAMETER MinProfileSizeMB
    Only consider profiles larger than specified MB.

.PARAMETER MaxProfileSizeMB
    Only consider profiles smaller than specified MB.

.PARAMETER IncludeCorrupted
    Include corrupted profiles in processing.

.PARAMETER FixCorruption
    Enable interactive corruption repair mode. Presents options to fix corrupted profiles:
    - Remove orphaned registry keys (for missing profile paths)
    - Delete corrupted profiles entirely
    - Recreate NTUSER.DAT from default template
    - Skip and continue. Requires -Interactive for full control.

.PARAMETER ProfileType
    Filter by profile type: Local, Roaming, Temporary, Mandatory, or All (default).

.PARAMETER Interactive
    Enable interactive mode for manual profile selection with visual menu.

.PARAMETER Test
    Test mode - validate prerequisites and connectivity without making changes.

.PARAMETER HtmlReport
    Generate professional HTML report at specified path.

.PARAMETER BackupPath
    Backup profiles to ZIP files before deletion (specify directory path).

.PARAMETER ConfigFile
    Load settings from JSON configuration file.

.PARAMETER UseParallel
    Use parallel processing for multiple computers.

.PARAMETER ThrottleLimit
    Maximum parallel threads when using -UseParallel. Default: 5

.PARAMETER SmtpServer
    SMTP server for email notifications.

.PARAMETER EmailTo
    Email recipient address for notifications.

.PARAMETER EmailFrom
    Email sender address. Default: delprofps@computername

.PARAMETER Detailed
    Show detailed folder breakdown for each profile (Documents, Downloads, Desktop, etc.)

.PARAMETER VerifyIntegrity
    Verify the script's SHA256 hash against DelprofPS.sha256 before execution.
    Exits with error if hash mismatch detected (unless -Force is also specified).
    Generate the hash file with: (Get-FileHash .\delprofPS.ps1 -Algorithm SHA256).Hash | Out-File .\DelprofPS.sha256

.PARAMETER Credential
    PSCredential object for authenticating to remote computers.
    Useful for cross-domain scenarios or when the current identity lacks access.
    If not specified, the current user's identity is used.

.EXAMPLE
    # List all profiles older than 30 days on local computer (dry run)
    .\DelprofPS.ps1

.EXAMPLE
    # Delete profiles older than 60 days, excluding administrators
    .\DelprofPS.ps1 -DaysInactive 60 -Delete -Exclude "*admin*"

.EXAMPLE
    # Interactive mode - select profiles visually before deletion
    .\DelprofPS.ps1 -DaysInactive 90 -Interactive

.EXAMPLE
    # Test connectivity to remote computers without making changes
    .\DelprofPS.ps1 -ComputerName SERVER1,SERVER2,SERVER3 -Test

.EXAMPLE
    # List profiles on remote computers showing disk space with progress bar
    .\DelprofPS.ps1 -ComputerName SERVER1,SERVER2 -ShowSpace

.EXAMPLE
    # Enterprise deployment with full reporting
    .\DelprofPS.ps1 -ComputerName (Import-Csv servers.csv).Name `
        -Delete -DaysInactive 90 `
        -LogPath "C:\Logs\delprof.log" `
        -OutputPath "C:\Logs\results.csv" `
        -HtmlReport "C:\Logs\report.html" `
        -BackupPath "C:\Backups\Profiles" `
        -UnloadHives -ShowSpace

.EXAMPLE
    # Scheduled task with email notification
    .\DelprofPS.ps1 -DaysInactive 60 -Delete -Exclude "*admin*" `
        -SmtpServer "mail.company.com" `
        -EmailTo "admin@company.com" `
        -EmailFrom "delprofps@server01" `
        -Quiet -LogPath "C:\Logs\delprof.log"

.EXAMPLE
    # Use JSON configuration file
    .\DelprofPS.ps1 -ConfigFile "C:\Config\delprof.json" -Delete

.EXAMPLE
    # Parallel processing for many computers
    .\DelprofPS.ps1 -ComputerName (Get-Content servers.txt) `
        -UseParallel -ThrottleLimit 10 `
        -DaysInactive 120 -Delete

.EXAMPLE
    # Show detailed folder breakdown for each profile
    .\DelprofPS.ps1 -DaysInactive 60 -Detailed -ShowSpace

.EXAMPLE
    # Process with all safety checks disabled (EXTREME CAUTION)
    .\DelprofPS.ps1 -Delete -Force -IgnoreActiveSessions -IncludeSystemProfiles

.NOTES
    Version:        2.0.0
    Author:         Karl Lawrence
    Creation Date:  2024
    
    REQUIREMENTS:
    - PowerShell 5.1 or later
    - Administrative privileges on target computers
    - Remote management enabled for remote computer processing
    
    SAFETY FEATURES:
    - Default dry-run mode (must use -Delete to actually remove profiles)
    - Active session protection (skips logged-in users unless -IgnoreActiveSessions)
    - System profile protection (excludes Default, Public, SYSTEM, etc.)
    - Registry hive unloading before deletion (-UnloadHives)
    - Profile backup capability (-BackupPath)
    - Comprehensive logging for audit trails
    - Windows Event Log integration
    
    EVENT LOG IDs:
    - 1000: Script started
    - 1001: HTML report generated
    - 1002: Script completed
    - 1005: Error - admin rights required
    - 1010: Profile deleted
    - 1012: Corruption repair action taken
    
.LINK
    Original Delprof2: https://helgeklein.com/delprof2/
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High', DefaultParameterSetName = 'List')]
param (
    [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [Alias('CN', 'MachineName', 'Server')]
    [string[]]$ComputerName = $env:COMPUTERNAME,

    [Parameter()]
    [Alias('Age', 'Days')]
    [ValidateRange(0, 3650)]
    [int]$DaysInactive = 30,

    [Parameter()]
    [ValidateSet('NTUSER_DAT', 'ProfilePath', 'Registry', 'LastLogon', 'LastLogoff')]
    [string]$AgeCalculation = 'NTUSER_DAT',

    [Parameter()]
    [string[]]$Include,

    [Parameter()]
    [string[]]$Exclude,

    [Parameter(ParameterSetName = 'Delete')]
    [switch]$Delete,

    [Parameter()]
    [switch]$Force,

    [Parameter(ParameterSetName = 'UI')]
    [switch]$UI,

    [Parameter(ParameterSetName = 'Preview')]
    [switch]$Preview,

    [Parameter()]
    [switch]$IgnoreActiveSessions,

    [Parameter()]
    [switch]$UnloadHives,

    [Parameter()]
    [ValidateRange(0, 50)]
    [int]$MaxRetries = 3,

    [Parameter()]
    [ValidateRange(1, 60)]
    [int]$RetryDelaySeconds = 2,

    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [string]$LogPath,

    [Parameter()]
    [switch]$Quiet,

    [Parameter()]
    [switch]$ShowSpace,

    [Parameter()]
    [switch]$IncludeSystemProfiles,

    [Parameter()]
    [switch]$IncludeSpecialProfiles,

    [Parameter()]
    [long]$MinProfileSizeMB,

    [Parameter()]
    [long]$MaxProfileSizeMB,

    [Parameter()]
    [switch]$IncludeCorrupted,

    [Parameter()]
    [switch]$FixCorruption,

    [Parameter()]
    [ValidateSet('Local', 'Roaming', 'Temporary', 'Mandatory', 'All')]
    [string]$ProfileType = 'All',

    [Parameter(ParameterSetName = 'Interactive')]
    [switch]$Interactive,

    [Parameter(ParameterSetName = 'Test')]
    [switch]$Test,

    [Parameter()]
    [string]$HtmlReport,

    [Parameter()]
    [string]$BackupPath,

    [Parameter()]
    [string]$ConfigFile,

    [Parameter()]
    [switch]$UseParallel,

    [Parameter()]
    [ValidateRange(1, 100)]
    [int]$ThrottleLimit = 5,

    [Parameter()]
    [string]$SmtpServer,

    [Parameter()]
    [string]$EmailTo,

    [Parameter()]
    [string]$EmailFrom = "delprofps@$env:COMPUTERNAME",

    [Parameter()]
    [switch]$Detailed,

    [Parameter()]
    [switch]$VerifyIntegrity,

    [Parameter()]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()]
    $Credential = [System.Management.Automation.PSCredential]::Empty
)

begin {
    #region GUI Function
    function Show-DelprofPSGUI {
        <#
        .SYNOPSIS
            Displays the Delprof2-PS graphical user interface.
        
        .DESCRIPTION
            Launches a modern WPF-based GUI providing visual access to all Delprof2-PS
            functionality including computer selection, filtering, actions, and real-time output.
        #>
        
        Add-Type -AssemblyName PresentationFramework
        Add-Type -AssemblyName PresentationCore
        Add-Type -AssemblyName WindowsBase
        Add-Type -AssemblyName System.Windows.Forms
        
        # XAML UI Definition
        [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Delprof2-PS Profile Manager" Height="750" Width="1000"
        WindowStartupLocation="CenterScreen" Background="#F5F5F5">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Background" Value="#007ACC"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
        <Style TargetType="GroupBox">
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Padding" Value="10"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#CCCCCC"/>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="200"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#007ACC" Padding="15">
            <StackPanel>
                <TextBlock Text="Delprof2-PS Profile Manager" Foreground="White" FontSize="24" FontWeight="Bold"/>
                <TextBlock Text="Enterprise User Profile Management" Foreground="#E0E0E0" FontSize="12"/>
            </StackPanel>
        </Border>
        
        <!-- Main Content -->
        <TabControl Grid.Row="1" Margin="10">
            <!-- Connection Tab -->
            <TabItem Header="Connection">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                        <RadioButton x:Name="rbLocalComputer" Content="Local Computer" IsChecked="True" Margin="0,0,20,0"/>
                        <RadioButton x:Name="rbRemoteComputers" Content="Remote Computers"/>
                    </StackPanel>
                    <TextBlock Grid.Row="1" Text="Computer Names (one per line or comma-separated):" Margin="0,0,0,5"/>
                    <TextBox Grid.Row="2" x:Name="txtComputerList" AcceptsReturn="True" VerticalScrollBarVisibility="Auto"
                             TextWrapping="Wrap" FontFamily="Consolas" FontSize="12"/>
                </Grid>
            </TabItem>
            
            <!-- Profiles Tab -->
            <TabItem Header="Profiles">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                        <Button x:Name="btnRefreshProfiles" Content="Refresh Profiles" Background="#17A2B8" Padding="10,6" Margin="0,0,5,0"/>
                        <Button x:Name="btnSelectAll" Content="Select All" Background="#6C757D" Padding="10,6" Margin="0,0,5,0"/>
                        <Button x:Name="btnDeselectAll" Content="Deselect All" Background="#6C757D" Padding="10,6" Margin="0,0,5,0"/>
                        <Button x:Name="btnForceRemove" Content="Force Remove Selected" Background="#DC3545" Padding="10,6" Margin="10,0,0,0"/>
                        <TextBlock x:Name="txtProfileCount" Text="" VerticalAlignment="Center" Margin="15,0,0,0" FontStyle="Italic" Foreground="#555555"/>
                    </StackPanel>
                    <DataGrid Grid.Row="1" x:Name="dgProfiles" AutoGenerateColumns="False" CanUserAddRows="False"
                              CanUserDeleteRows="False" IsReadOnly="False" SelectionMode="Extended"
                              HeadersVisibility="Column" GridLinesVisibility="Horizontal"
                              AlternatingRowBackground="#F9F9F9" RowBackground="White"
                              BorderBrush="#CCCCCC" BorderThickness="1"
                              VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto">
                        <DataGrid.Columns>
                            <DataGridCheckBoxColumn Header="Select" Binding="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Width="55"/>
                            <DataGridTextColumn Header="Username" Binding="{Binding UserName}" Width="150" IsReadOnly="True"/>
                            <DataGridTextColumn Header="Profile Path" Binding="{Binding ProfilePath}" Width="250" IsReadOnly="True"/>
                            <DataGridTextColumn Header="SID" Binding="{Binding SID}" Width="280" IsReadOnly="True"/>
                            <DataGridTextColumn Header="Size (MB)" Binding="{Binding SizeMB}" Width="80" IsReadOnly="True"/>
                            <DataGridTextColumn Header="Last Modified" Binding="{Binding LastModified}" Width="140" IsReadOnly="True"/>
                            <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="80" IsReadOnly="True"/>
                        </DataGrid.Columns>
                    </DataGrid>
                    <TextBlock Grid.Row="2" Text="Select profiles above and click 'Force Remove Selected' to delete them regardless of age filters." 
                               Foreground="#888888" FontSize="11" Margin="0,8,0,0" TextWrapping="Wrap"/>
                </Grid>
            </TabItem>
            
            <!-- Filters Tab -->
            <TabItem Header="Filters">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="10">
                        <GroupBox Header="Age Settings">
                            <Grid Margin="5">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Row="0" Grid.Column="0" Text="Days Inactive:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                                <Slider Grid.Row="0" Grid.Column="1" x:Name="sldDaysInactive" Minimum="0" Maximum="365" Value="30" TickFrequency="1"/>
                                <TextBlock Grid.Row="0" Grid.Column="2" x:Name="txtDaysValue" Text="30 days" Width="60" Margin="10,0,0,0"/>
                                
                                <TextBlock Grid.Row="1" Grid.Column="0" Text="Age Method:" VerticalAlignment="Center" Margin="0,10,10,0"/>
                                <ComboBox Grid.Row="1" Grid.Column="1" x:Name="cmbAgeMethod" SelectedIndex="0">
                                    <ComboBoxItem Content="NTUSER.DAT modified"/>
                                    <ComboBoxItem Content="Profile Path modified"/>
                                    <ComboBoxItem Content="Registry value"/>
                                    <ComboBoxItem Content="Last Logon"/>
                                    <ComboBoxItem Content="Last Logoff"/>
                                </ComboBox>
                                
                                <TextBlock Grid.Row="2" Grid.Column="0" Text="Profile Type:" VerticalAlignment="Center" Margin="0,10,10,0"/>
                                <ComboBox Grid.Row="2" Grid.Column="1" x:Name="cmbProfileType" SelectedIndex="0">
                                    <ComboBoxItem Content="All Profiles"/>
                                    <ComboBoxItem Content="Local Only"/>
                                    <ComboBoxItem Content="Roaming Only"/>
                                    <ComboBoxItem Content="Temporary Only"/>
                                    <ComboBoxItem Content="Mandatory Only"/>
                                </ComboBox>
                            </Grid>
                        </GroupBox>
                        
                        <GroupBox Header="Include/Exclude Patterns">
                            <Grid Margin="5">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <TextBlock Grid.Row="0" Text="Include Pattern (e.g., user*, admin*):" Margin="0,0,0,5"/>
                                <TextBox Grid.Row="1" x:Name="txtInclude" Margin="0,0,0,10"/>
                                <TextBlock Grid.Row="2" Text="Exclude Pattern (e.g., Administrator*, *service*):" Margin="0,0,0,5"/>
                                <TextBox Grid.Row="3" x:Name="txtExclude" Text="Administrator*, *admin*, *service*"/>
                            </Grid>
                        </GroupBox>
                        
                        <GroupBox Header="Size Filters (MB)">
                            <StackPanel Orientation="Horizontal" Margin="5">
                                <TextBlock Text="Min Size:" VerticalAlignment="Center" Margin="0,0,5,0"/>
                                <TextBox x:Name="txtMinSize" Width="80" Margin="0,0,20,0"/>
                                <TextBlock Text="Max Size:" VerticalAlignment="Center" Margin="0,0,5,0"/>
                                <TextBox x:Name="txtMaxSize" Width="80"/>
                            </StackPanel>
                        </GroupBox>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            
            <!-- Actions Tab -->
            <TabItem Header="Actions">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="10">
                        <GroupBox Header="Operation Mode">
                            <StackPanel Margin="5">
                                <RadioButton x:Name="rbModePreview" Content="Preview Only (no changes)" IsChecked="True" Margin="0,2"/>
                                <RadioButton x:Name="rbModeDelete" Content="Delete Profiles" Margin="0,2"/>
                            </StackPanel>
                        </GroupBox>
                        
                        <GroupBox Header="Options">
                            <UniformGrid Columns="2" Margin="5">
                                <CheckBox x:Name="chkIncludeCorrupted" Content="Include Corrupted" Margin="5"/>
                                <CheckBox x:Name="chkFixCorruption" Content="Fix Corruption (Interactive)" Margin="5"/>
                                <CheckBox x:Name="chkShowSpace" Content="Show Disk Space" IsChecked="True" Margin="5"/>
                                <CheckBox x:Name="chkDetailed" Content="Detailed Folder Breakdown" Margin="5"/>
                                <CheckBox x:Name="chkUnloadHives" Content="Unload Registry Hives" IsChecked="True" Margin="5"/>
                                <CheckBox x:Name="chkIgnoreActive" Content="Ignore Active Sessions (DANGER)" Margin="5"/>
                                <CheckBox x:Name="chkIncludeSystem" Content="Include System Profiles" Margin="5"/>
                                <CheckBox x:Name="chkIncludeSpecial" Content="Include Special Profiles" Margin="5"/>
                                <CheckBox x:Name="chkForce" Content="Force (Skip Confirmations)" Margin="5"/>
                                <CheckBox x:Name="chkInteractive" Content="Interactive Mode" Margin="5"/>
                                <CheckBox x:Name="chkQuiet" Content="Quiet Mode" Margin="5"/>
                                <CheckBox x:Name="chkTestMode" Content="Test Mode (Validate Only)" Margin="5"/>
                            </UniformGrid>
                        </GroupBox>
                        
                        <GroupBox Header="Parallel Processing">
                            <Grid Margin="5">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <CheckBox Grid.Row="0" x:Name="chkUseParallel" Content="Enable Parallel Processing" Margin="0,0,0,10"/>
                                <StackPanel Grid.Row="1" Orientation="Horizontal">
                                    <TextBlock Text="Throttle Limit:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                                    <Slider x:Name="sldThrottle" Minimum="1" Maximum="20" Value="5" Width="200" TickFrequency="1"/>
                                    <TextBlock x:Name="txtThrottleValue" Text="5" Width="30" Margin="10,0,0,0"/>
                                </StackPanel>
                            </Grid>
                        </GroupBox>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            
            <!-- Output Tab -->
            <TabItem Header="Output &amp; Reporting">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="10">
                        <GroupBox Header="Backup Settings">
                            <Grid Margin="5">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <CheckBox Grid.Column="0" x:Name="chkBackup" Content="Enable Backup" Margin="0,0,10,0"/>
                                <TextBox Grid.Column="1" x:Name="txtBackupPath" IsEnabled="{Binding ElementName=chkBackup, Path=IsChecked}"/>
                                <Button Grid.Column="2" x:Name="btnBrowseBackup" Content="Browse..." Width="80" Margin="10,0,0,0"/>
                            </Grid>
                        </GroupBox>
                        
                        <GroupBox Header="Log &amp; Export Paths">
                            <Grid Margin="5">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                
                                <CheckBox Grid.Row="0" Grid.Column="0" x:Name="chkLogPath" Content="Log File:" Margin="0,5"/>
                                <TextBox Grid.Row="0" Grid.Column="1" x:Name="txtLogPath" IsEnabled="{Binding ElementName=chkLogPath, Path=IsChecked}"/>
                                <Button Grid.Row="0" Grid.Column="2" x:Name="btnBrowseLog" Content="Browse..." Width="80" Margin="10,0,0,0"/>
                                
                                <CheckBox Grid.Row="1" Grid.Column="0" x:Name="chkOutputCSV" Content="CSV Output:" Margin="0,5"/>
                                <TextBox Grid.Row="1" Grid.Column="1" x:Name="txtOutputPath" IsEnabled="{Binding ElementName=chkOutputCSV, Path=IsChecked}"/>
                                <Button Grid.Row="1" Grid.Column="2" x:Name="btnBrowseCSV" Content="Browse..." Width="80" Margin="10,0,0,0"/>
                                
                                <CheckBox Grid.Row="2" Grid.Column="0" x:Name="chkHtmlReport" Content="HTML Report:" Margin="0,5"/>
                                <TextBox Grid.Row="2" Grid.Column="1" x:Name="txtHtmlPath" IsEnabled="{Binding ElementName=chkHtmlReport, Path=IsChecked}"/>
                                <Button Grid.Row="2" Grid.Column="2" x:Name="btnBrowseHtml" Content="Browse..." Width="80" Margin="10,0,0,0"/>
                            </Grid>
                        </GroupBox>
                        
                        <GroupBox Header="Email Notifications">
                            <Grid Margin="5">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Row="0" Grid.Column="0" Text="SMTP Server:" Width="100" VerticalAlignment="Center"/>
                                <TextBox Grid.Row="0" Grid.Column="1" x:Name="txtSmtpServer"/>
                                <TextBlock Grid.Row="1" Grid.Column="0" Text="To:" Width="100" VerticalAlignment="Center" Margin="0,5,0,0"/>
                                <TextBox Grid.Row="1" Grid.Column="1" x:Name="txtEmailTo" Margin="0,5,0,0"/>
                                <TextBlock Grid.Row="2" Grid.Column="0" Text="From:" Width="100" VerticalAlignment="Center" Margin="0,5,0,0"/>
                                <TextBox Grid.Row="2" Grid.Column="1" x:Name="txtEmailFrom" Text="delprofps@localhost" Margin="0,5,0,0"/>
                            </Grid>
                        </GroupBox>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
        </TabControl>
        
        <!-- Action Buttons -->
        <Border Grid.Row="2" Background="#E8E8E8" Padding="10">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                <Button x:Name="btnLoadConfig" Content="Load Config" Background="#6C757D"/>
                <Button x:Name="btnSaveConfig" Content="Save Config" Background="#6C757D"/>
                <Button x:Name="btnClear" Content="Clear Output" Background="#FFC107" Foreground="Black"/>
                <Button x:Name="btnStop" Content="Stop" Background="#DC3545" IsEnabled="False"/>
                <Button x:Name="btnRun" Content="RUN" Width="120" FontSize="14" Background="#28A745"/>
            </StackPanel>
        </Border>
        
        <!-- Output Console -->
        <Border Grid.Row="3" Background="#1E1E1E" Margin="10,0">
            <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto">
                <TextBox x:Name="txtOutput" Background="Transparent" Foreground="#D4D4D4"
                         FontFamily="Consolas" FontSize="11" IsReadOnly="True"
                         TextWrapping="Wrap" BorderThickness="0" Padding="5"/>
            </ScrollViewer>
        </Border>
        
        <!-- Progress Bar -->
        <ProgressBar Grid.Row="4" x:Name="progressBar" Height="20" Margin="10"
                     IsIndeterminate="False" Visibility="Collapsed"/>
    </Grid>
</Window>
"@
        
        # Load XAML
        $reader = New-Object System.Xml.XmlNodeReader $xaml
        $window = [Windows.Markup.XamlReader]::Load($reader)
        
        # Get controls reference
        $controls = @{}
        $nsMgr = New-Object System.Xml.XmlNamespaceManager($xaml.NameTable)
        $nsMgr.AddNamespace('x', 'http://schemas.microsoft.com/winfx/2006/xaml')
        $xaml.SelectNodes("//*[@x:Name]", $nsMgr) | ForEach-Object {
            $name = $_.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml')
            $controls[$name] = $window.FindName($name)
        }
        
        # Script-level variables for GUI state
        $script:guiState = @{ Running = $false; StopRequested = $false }
        $script:scriptRoot = $PSScriptRoot
        
        # Stop any stale timers from a previous GUI session in the same PS process
        if ($script:activeRefreshTimer) {
            try { $script:activeRefreshTimer.Stop() } catch {}
            $script:activeRefreshTimer = $null
        }
        if ($script:activeRunTimer) {
            try { $script:activeRunTimer.Stop() } catch {}
            $script:activeRunTimer = $null
        }
        
        # Initialise index counters defensively (prevents null-index if stale timer fires)
        $script:refreshInfoIndex = 0
        $script:lastOutputIndex = 0
        $script:lastInfoIndex = 0
        $script:lastWarningIndex = 0
        $script:lastErrorIndex = 0
        
        # Output function for GUI
        $script:WriteGuiOutput = {
            param([string]$Text, [string]$Color = "White")
            $timestamp = Get-Date -Format "HH:mm:ss"
            $coloredText = "[$timestamp] $Text"
            $controls['txtOutput'].AppendText("$coloredText`r`n")
            $controls['txtOutput'].ScrollToEnd()
        }
        
        # Default log path to script directory
        $defaultLogPath = Join-Path $PSScriptRoot "DelprofPS_$(Get-Date -Format 'yyyyMMdd').log"
        $controls['txtLogPath'].Text = $defaultLogPath
        $controls['chkLogPath'].IsChecked = $true
        
        # Reusable config loader — applies a parsed JSON config object to the GUI controls
        $script:ApplyConfig = {
            param([psobject]$config, [string]$source)
            try {
                # Connection
                if ($null -ne $config.RemoteComputers) { $controls['rbRemoteComputers'].IsChecked = [bool]$config.RemoteComputers; $controls['rbLocalComputer'].IsChecked = -not [bool]$config.RemoteComputers }
                if ($config.ComputerList) { $controls['txtComputerList'].Text = $config.ComputerList }
                
                # Age / Type filters
                if ($null -ne $config.DaysInactive) { $controls['sldDaysInactive'].Value = $config.DaysInactive }
                if ($null -ne $config.AgeMethod) { $controls['cmbAgeMethod'].SelectedIndex = [int]$config.AgeMethod }
                if ($null -ne $config.ProfileType) { $controls['cmbProfileType'].SelectedIndex = [int]$config.ProfileType }
                
                # Pattern filters
                if ($config.Exclude) { $controls['txtExclude'].Text = ($config.Exclude -join ', ') }
                if ($config.Include) { $controls['txtInclude'].Text = ($config.Include -join ', ') }
                
                # Size filters
                if ($config.MinSize) { $controls['txtMinSize'].Text = $config.MinSize }
                if ($config.MaxSize) { $controls['txtMaxSize'].Text = $config.MaxSize }
                
                # Operation mode
                if ($null -ne $config.DeleteMode) { $controls['rbModeDelete'].IsChecked = [bool]$config.DeleteMode; $controls['rbModePreview'].IsChecked = -not [bool]$config.DeleteMode }
                
                # Checkboxes (Actions tab)
                if ($null -ne $config.IncludeCorrupted) { $controls['chkIncludeCorrupted'].IsChecked = [bool]$config.IncludeCorrupted }
                if ($null -ne $config.FixCorruption) { $controls['chkFixCorruption'].IsChecked = [bool]$config.FixCorruption }
                if ($null -ne $config.ShowSpace) { $controls['chkShowSpace'].IsChecked = [bool]$config.ShowSpace }
                if ($null -ne $config.Detailed) { $controls['chkDetailed'].IsChecked = [bool]$config.Detailed }
                if ($null -ne $config.UnloadHives) { $controls['chkUnloadHives'].IsChecked = [bool]$config.UnloadHives }
                if ($null -ne $config.IgnoreActive) { $controls['chkIgnoreActive'].IsChecked = [bool]$config.IgnoreActive }
                if ($null -ne $config.IncludeSystem) { $controls['chkIncludeSystem'].IsChecked = [bool]$config.IncludeSystem }
                if ($null -ne $config.IncludeSpecial) { $controls['chkIncludeSpecial'].IsChecked = [bool]$config.IncludeSpecial }
                if ($null -ne $config.Force) { $controls['chkForce'].IsChecked = [bool]$config.Force }
                if ($null -ne $config.Interactive) { $controls['chkInteractive'].IsChecked = [bool]$config.Interactive }
                if ($null -ne $config.Quiet) { $controls['chkQuiet'].IsChecked = [bool]$config.Quiet }
                if ($null -ne $config.TestMode) { $controls['chkTestMode'].IsChecked = [bool]$config.TestMode }
                
                # Parallel
                if ($null -ne $config.UseParallel) { $controls['chkUseParallel'].IsChecked = [bool]$config.UseParallel }
                if ($null -ne $config.ThrottleLimit) { $controls['sldThrottle'].Value = $config.ThrottleLimit }
                
                # Output paths
                if ($config.LogPath) { $controls['txtLogPath'].Text = $config.LogPath; $controls['chkLogPath'].IsChecked = $true }
                if ($config.OutputPath) { $controls['txtOutputPath'].Text = $config.OutputPath; $controls['chkOutputCSV'].IsChecked = $true }
                if ($config.HtmlReport) { $controls['txtHtmlPath'].Text = $config.HtmlReport; $controls['chkHtmlReport'].IsChecked = $true }
                if ($config.BackupPath) { $controls['txtBackupPath'].Text = $config.BackupPath; $controls['chkBackup'].IsChecked = $true }
                
                # Email
                if ($config.SmtpServer) { $controls['txtSmtpServer'].Text = $config.SmtpServer }
                if ($config.EmailTo) { $controls['txtEmailTo'].Text = $config.EmailTo }
                if ($config.EmailFrom) { $controls['txtEmailFrom'].Text = $config.EmailFrom }
                
                & $script:WriteGuiOutput -Text "[Config] Loaded from $source" -Color 'Green'
                & $script:WriteGuiOutput -Text "[Config] Applied: DaysInactive=$($controls['sldDaysInactive'].Value), Exclude=$($controls['txtExclude'].Text), Include=$($controls['txtInclude'].Text)" -Color 'Gray'
            }
            catch {
                & $script:WriteGuiOutput -Text "[Config] ERROR applying config: $($_.Exception.Message)" -Color 'Red'
            }
        }
        
        & $script:WriteGuiOutput -Text '=== Delprof2-PS GUI Initialized ===' -Color 'Cyan'
        & $script:WriteGuiOutput -Text "PowerShell Version: $($PSVersionTable.PSVersion)" -Color 'Gray'
        & $script:WriteGuiOutput -Text "Running as: $([Security.Principal.WindowsIdentity]::GetCurrent().Name)" -Color 'Gray'
        & $script:WriteGuiOutput -Text "Computer: $env:COMPUTERNAME" -Color 'Gray'
        
        # Auto-load config if exactly one .config.json file exists next to the script
        $configFiles = @(Get-ChildItem -Path $PSScriptRoot -Filter '*.config.json' -File -ErrorAction SilentlyContinue)
        if ($configFiles.Count -eq 1) {
            & $script:WriteGuiOutput -Text "[Config] Found config file: $($configFiles[0].Name) — auto-loading..." -Color 'Cyan'
            try {
                $autoConfig = Get-Content $configFiles[0].FullName -Raw | ConvertFrom-Json
                & $script:ApplyConfig -config $autoConfig -source $configFiles[0].Name
            }
            catch {
                & $script:WriteGuiOutput -Text "[Config] ERROR: Failed to auto-load config: $($_.Exception.Message)" -Color 'Red'
            }
        }
        elseif ($configFiles.Count -gt 1) {
            & $script:WriteGuiOutput -Text "[Config] Multiple config files found ($($configFiles.Count)) — use Load Config to choose one" -Color 'Yellow'
        }
        
        & $script:WriteGuiOutput -Text 'Ready. Use the tabs above to configure, then click RUN or Refresh Profiles.' -Color 'Green'
        
        # Event Handler: Throttle Slider
        $controls['sldThrottle'].Add_ValueChanged({
            $controls['txtThrottleValue'].Text = $controls['sldThrottle'].Value
            & $script:WriteGuiOutput -Text "[Setting] Throttle limit changed to $($controls['sldThrottle'].Value)" -Color 'Gray'
        })
        
        # Event Handler: Days Slider
        $controls['sldDaysInactive'].Add_ValueChanged({
            $controls['txtDaysValue'].Text = "$($controls['sldDaysInactive'].Value) days"
            & $script:WriteGuiOutput -Text "[Setting] Days inactive changed to $($controls['sldDaysInactive'].Value)" -Color 'Gray'
        })
        
        # Profile list data source
        $script:profileList = [System.Collections.ObjectModel.ObservableCollection[PSObject]]::new()
        $controls['dgProfiles'].ItemsSource = $script:profileList
        
        # Helper: Enumerate profiles on a computer (runs in background runspace)
        $script:RefreshProfileList = {
            param([string]$TargetComputer)
            
            # Disable the refresh button and show progress
            $controls['btnRefreshProfiles'].IsEnabled = $false
            $controls['txtProfileCount'].Text = "Scanning..."
            $controls['progressBar'].Visibility = "Visible"
            $controls['progressBar'].IsIndeterminate = $true
            
            $profileListRef = $script:profileList
            $controlsRef = $controls
            
            # Clear existing items on the UI thread
            $script:profileList.Clear()
            
            # Build a background runspace to scan without freezing the GUI
            $runspace = [runspacefactory]::CreateRunspace()
            $runspace.ApartmentState = 'STA'
            $runspace.Open()
            $runspace.SessionStateProxy.SetVariable('TargetComputer', $TargetComputer)
            $runspace.SessionStateProxy.SetVariable('localComputerName', $env:COMPUTERNAME)
            
            $ps = [powershell]::Create().AddScript({
                $profileListPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
                $isLocal = ($TargetComputer -eq $localComputerName -or $TargetComputer -eq 'localhost' -or $TargetComputer -eq '.')
                $systemNames = @('Default', 'Default User', 'Public', 'SYSTEM', 'LocalService', 'NetworkService', 'systemprofile')
                $results = [System.Collections.Generic.List[object]]::new()
                $errorMessages = [System.Collections.Generic.List[string]]::new()
                
                Write-Host "[Scan] Starting profile scan on $TargetComputer (local=$isLocal)..."
                
                try {
                    if ($isLocal) {
                        Write-Host "[Scan] Reading local registry ProfileList..."
                        $profileKeys = Get-ChildItem $profileListPath -ErrorAction Stop |
                            Where-Object { $_.PSChildName -match '^S-1-5-21' }
                        
                        $totalKeys = @($profileKeys).Count
                        Write-Host "[Scan] Found $totalKeys profile SIDs to process"
                        $current = 0
                        
                        foreach ($key in $profileKeys) {
                            $current++
                            try {
                                $props = Get-ItemProperty $key.PSPath
                                $sid = $key.PSChildName
                                $profilePath = $props.ProfileImagePath
                                
                                Write-Host "[Scan] ($current/$totalKeys) Processing SID $sid..."
                                
                                $userName = $null
                                try {
                                    $secId = New-Object System.Security.Principal.SecurityIdentifier($sid)
                                    $ntAccount = $secId.Translate([System.Security.Principal.NTAccount])
                                    $userName = $ntAccount.Value
                                } catch {
                                    $userName = "(Unresolvable)"
                                }
                                
                                Write-Host "[Scan]   User: $userName | Path: $profilePath"
                                
                                $sizeMB = "N/A"
                                $lastMod = "N/A"
                                if ($profilePath -and (Test-Path $profilePath)) {
                                    Write-Host "[Scan]   Calculating folder size..."
                                    try {
                                        $dirInfo = Get-Item $profilePath -ErrorAction SilentlyContinue
                                        $lastMod = $dirInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm")
                                        $ntUserDat = Join-Path $profilePath "NTUSER.DAT"
                                        if (Test-Path $ntUserDat) {
                                            $lastMod = (Get-Item $ntUserDat -Force -ErrorAction SilentlyContinue).LastWriteTime.ToString("yyyy-MM-dd HH:mm")
                                        }
                                        # Use robocopy /L (list-only) for fast native size calculation, /XJ skips junctions
                                        $totalSize = 0
                                        try {
                                            $roboOut = & robocopy $profilePath 'C:\RobocopyNull' /L /E /BYTES /NJH /NC /NDL /NFL /XJ /R:0 /W:0 2>&1
                                            $roboText = ($roboOut | Out-String)
                                            if ($roboText -match 'Bytes\s*:\s*(\d+)') {
                                                $totalSize = [long]$Matches[1]
                                            }
                                        } catch {
                                            $totalSize = 0
                                        }
                                        $sizeMB = [math]::Round($totalSize / 1MB, 1)
                                        Write-Host "[Scan]   Size: $sizeMB MB | Last modified: $lastMod"
                                    } catch {
                                        $sizeMB = "Error"
                                        Write-Host "[Scan]   Size calculation error: $_"
                                    }
                                }
                                else {
                                    Write-Host "[Scan]   Path not found or missing"
                                }
                                
                                $status = "OK"
                                $shortName = if ($userName) { ($userName -split '\\')[-1] } else { "" }
                                if ($systemNames -contains $shortName) {
                                    $status = "System"
                                }
                                elseif (-not $profilePath -or -not (Test-Path $profilePath)) {
                                    $status = "Orphaned"
                                }
                                
                                Write-Host "[Scan]   Status: $status"
                                
                                $results.Add([PSCustomObject]@{
                                    Selected = $false
                                    UserName = $userName
                                    ProfilePath = $profilePath
                                    SID = $sid
                                    SizeMB = $sizeMB
                                    LastModified = $lastMod
                                    Status = $status
                                })
                            }
                            catch {
                                $errorMessages.Add("Error reading profile $($key.PSChildName): $_")
                                Write-Host "[Scan]   ERROR: $_"
                            }
                        }
                    }
                    else {
                        Write-Host "[Scan] Connecting to remote registry on $TargetComputer via WMI..."
                        $regProv = Get-WmiObject -ComputerName $TargetComputer -Class StdRegProv -Namespace 'root\default' -ErrorAction Stop
                        $enumResult = $regProv.EnumKey(2147483650, 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList')
                        $sids = $enumResult.sNames | Where-Object { $_ -match '^S-1-5-21' }
                        
                        $totalSids = @($sids).Count
                        Write-Host "[Scan] Found $totalSids profile SIDs on remote $TargetComputer"
                        $current = 0
                        
                        foreach ($sid in $sids) {
                            $current++
                            Write-Host "[Scan] ($current/$totalSids) Processing remote SID $sid..."
                            try {
                                $pathResult = $regProv.GetStringValue(2147483650, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid", 'ProfileImagePath')
                                $profilePath = $pathResult.sValue
                                if ($profilePath) { $profilePath = $profilePath -replace '%SystemDrive%', 'C:' }
                                
                                $userName = $null
                                try {
                                    $secId = New-Object System.Security.Principal.SecurityIdentifier($sid)
                                    $ntAccount = $secId.Translate([System.Security.Principal.NTAccount])
                                    $userName = $ntAccount.Value
                                } catch {
                                    $userName = "(Unresolvable)"
                                }
                                
                                Write-Host "[Scan]   User: $userName | Path: $profilePath"
                                
                                $results.Add([PSCustomObject]@{
                                    Selected = $false
                                    UserName = $userName
                                    ProfilePath = $profilePath
                                    SID = $sid
                                    SizeMB = "Remote"
                                    LastModified = "Remote"
                                    Status = "OK"
                                })
                            }
                            catch {
                                $errorMessages.Add("Error reading remote profile $sid`: $_")
                                Write-Host "[Scan]   ERROR: $_"
                            }
                        }
                    }
                }
                catch {
                    $errorMessages.Add("Failed to enumerate profiles: $($_.Exception.Message)")
                    Write-Host "[Scan] FATAL: $($_.Exception.Message)"
                }
                
                Write-Host "[Scan] Scan complete. Found $($results.Count) profiles."
                return @{ Results = $results; Errors = $errorMessages; Computer = $TargetComputer }
            })
            
            $ps.Runspace = $runspace
            $asyncResult = $ps.BeginInvoke()
            
            # Mutable state hashtable — reference type persists across ticks inside .GetNewClosure()
            $pollState = @{ InfoIdx = 0 }
            
            # Timer to poll for progress and completion without blocking the UI
            $timer = New-Object System.Windows.Threading.DispatcherTimer
            $script:activeRefreshTimer = $timer
            $timer.Interval = [TimeSpan]::FromMilliseconds(150)
            $timer.Add_Tick({
                try {
                    # Guard against stale timer from a previous GUI session
                    if ($null -eq $ps -or $null -eq $asyncResult) { $timer.Stop(); return }
                    
                    # Drain Information stream for live progress (write directly via $controlsRef to avoid scope issues)
                    while ($pollState.InfoIdx -lt $ps.Streams.Information.Count) {
                        $msg = "$($ps.Streams.Information[$pollState.InfoIdx].MessageData)"
                        $controlsRef['txtOutput'].AppendText("[$(Get-Date -Format 'HH:mm:ss')] $msg`r`n")
                        $controlsRef['txtOutput'].ScrollToEnd()
                        $pollState.InfoIdx++
                    }
                    
                    # Update scanning status with count
                    $scannedSoFar = $ps.Streams.Information | Where-Object { "$($_.MessageData)" -match '^\[Scan\] \(\d+' }
                    if ($scannedSoFar) {
                        $controlsRef['txtProfileCount'].Text = "Scanning... ($(@($scannedSoFar).Count) processed)"
                    }
                    
                    if ($asyncResult.IsCompleted) {
                        $timer.Stop()
                        $script:activeRefreshTimer = $null
                        try {
                            $output = $ps.EndInvoke($asyncResult)
                            $data = $output | Select-Object -Last 1
                            
                            if ($data.Errors) {
                                foreach ($err in $data.Errors) {
                                    $controlsRef['txtOutput'].AppendText("[$(Get-Date -Format 'HH:mm:ss')] [ERROR] $err`r`n")
                                    $controlsRef['txtOutput'].ScrollToEnd()
                                }
                            }
                            if ($data.Results) {
                                foreach ($item in $data.Results) {
                                    $profileListRef.Add($item)
                                }
                                $controlsRef['txtProfileCount'].Text = "$($data.Results.Count) profiles found"
                                $controlsRef['txtOutput'].AppendText("[$(Get-Date -Format 'HH:mm:ss')] Found $($data.Results.Count) profiles on $($data.Computer)`r`n")
                                $controlsRef['txtOutput'].ScrollToEnd()
                            }
                            else {
                                $controlsRef['txtProfileCount'].Text = "0 profiles found"
                                $controlsRef['txtOutput'].AppendText("[$(Get-Date -Format 'HH:mm:ss')] No profiles returned`r`n")
                                $controlsRef['txtOutput'].ScrollToEnd()
                            }
                        }
                        catch {
                            $controlsRef['txtProfileCount'].Text = "Error scanning profiles"
                            $controlsRef['txtOutput'].AppendText("[$(Get-Date -Format 'HH:mm:ss')] Scan error: $($_.Exception.Message)`r`n")
                            $controlsRef['txtOutput'].ScrollToEnd()
                        }
                        finally {
                            $ps.Dispose()
                            $runspace.Close()
                            $runspace.Dispose()
                            $controlsRef['btnRefreshProfiles'].IsEnabled = $true
                            $controlsRef['progressBar'].Visibility = "Collapsed"
                        }
                    }
                }
                catch {
                    # Log the error visibly, then clean up
                    try {
                        $controlsRef['txtOutput'].AppendText("[$(Get-Date -Format 'HH:mm:ss')] [Timer Error] $($_.Exception.Message)`r`n")
                        $controlsRef['txtOutput'].ScrollToEnd()
                    } catch {}
                    try { $timer.Stop() } catch {}
                    $script:activeRefreshTimer = $null
                    $controlsRef['btnRefreshProfiles'].IsEnabled = $true
                    $controlsRef['progressBar'].Visibility = "Collapsed"
                }
            }.GetNewClosure())
            $timer.Start()
        }
        
        # Event Handler: Refresh Profiles
        $controls['btnRefreshProfiles'].Add_Click({
            & $script:WriteGuiOutput -Text '[Button] Refresh Profiles clicked' -Color 'Cyan'
            $targetComputer = $env:COMPUTERNAME
            if ($controls['rbRemoteComputers'].IsChecked -and $controls['txtComputerList'].Text) {
                $computers = $controls['txtComputerList'].Text -split "[\r\n,]+" | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                if ($computers.Count -gt 0) { $targetComputer = $computers[0] }
            }
            & $script:RefreshProfileList -TargetComputer $targetComputer
        })
        
        # Event Handler: Select All
        $controls['btnSelectAll'].Add_Click({
            & $script:WriteGuiOutput -Text "[Button] Select All clicked - selecting $($script:profileList.Count) profiles" -Color 'Cyan'
            foreach ($item in $script:profileList) {
                $item.Selected = $true
            }
            $controls['dgProfiles'].Items.Refresh()
            & $script:WriteGuiOutput -Text "[Select] All $($script:profileList.Count) profiles selected" -Color 'Gray'
        })
        
        # Event Handler: Deselect All
        $controls['btnDeselectAll'].Add_Click({
            & $script:WriteGuiOutput -Text '[Button] Deselect All clicked' -Color 'Cyan'
            foreach ($item in $script:profileList) {
                $item.Selected = $false
            }
            $controls['dgProfiles'].Items.Refresh()
            & $script:WriteGuiOutput -Text '[Select] All profiles deselected' -Color 'Gray'
        })
        
        # Event Handler: Force Remove Selected
        $controls['btnForceRemove'].Add_Click({
            & $script:WriteGuiOutput -Text '[Button] Force Remove Selected clicked' -Color 'Cyan'
            if ($script:guiState.Running) {
                & $script:WriteGuiOutput -Text '[Force Remove] Blocked - another operation is already running' -Color 'Yellow'
                [System.Windows.MessageBox]::Show("An operation is already running. Please wait.", "Busy", "OK", "Warning")
                return
            }
            
            $selected = @($script:profileList | Where-Object { $_.Selected -eq $true })
            & $script:WriteGuiOutput -Text "[Force Remove] Found $($selected.Count) selected profile(s)" -Color 'Gray'
            if ($selected.Count -eq 0) {
                & $script:WriteGuiOutput -Text '[Force Remove] No profiles selected - aborting' -Color 'Yellow'
                [System.Windows.MessageBox]::Show("No profiles selected. Use the checkboxes to select profiles to remove.", "No Selection", "OK", "Information")
                return
            }
            
            $userList = ($selected | ForEach-Object { "$($_.UserName) ($($_.SID))" }) -join "`n"
            $confirm = [System.Windows.MessageBox]::Show(
                "Are you sure you want to FORCE REMOVE the following $($selected.Count) profile(s)?`n`n$userList`n`nThis action CANNOT be undone!",
                "Confirm Force Removal",
                "YesNo",
                "Warning"
            )
            
            if ($confirm -ne 'Yes') {
                & $script:WriteGuiOutput -Text '[Force Remove] User cancelled confirmation dialog' -Color 'Yellow'
                return
            }
            & $script:WriteGuiOutput -Text '[Force Remove] User confirmed - proceeding with removal' -Color 'Yellow'
            
            $script:guiState.Running = $true
            $controls['btnForceRemove'].IsEnabled = $false
            $controls['progressBar'].Visibility = "Visible"
            $controls['progressBar'].IsIndeterminate = $true
            
            $targetComputer = $env:COMPUTERNAME
            if ($controls['rbRemoteComputers'].IsChecked -and $controls['txtComputerList'].Text) {
                $computers = $controls['txtComputerList'].Text -split "[\r\n,]+" | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                if ($computers.Count -gt 0) { $targetComputer = $computers[0] }
            }
            
            $isLocal = ($targetComputer -eq $env:COMPUTERNAME -or $targetComputer -eq 'localhost' -or $targetComputer -eq '.')
            $removedCount = 0
            $failedCount = 0
            $totalSelected = $selected.Count
            $currentIndex = 0
            
            & $script:WriteGuiOutput -Text "Starting force removal of $totalSelected profile(s) on $targetComputer (local=$isLocal)" -Color "Cyan"
            & $script:WriteGuiOutput -Text "---" -Color "Gray"
            
            foreach ($prof in $selected) {
                $currentIndex++
                & $script:WriteGuiOutput -Text "[$currentIndex/$totalSelected] Removing: $($prof.UserName) | SID: $($prof.SID)" -Color "Yellow"
                & $script:WriteGuiOutput -Text "  Path: $($prof.ProfilePath) | Size: $($prof.SizeMB) MB | Status: $($prof.Status)" -Color "Gray"
                [System.Windows.Forms.Application]::DoEvents()
                
                $success = $true
                
                # Step 1: Unload registry hive if loaded
                & $script:WriteGuiOutput -Text "  Step 1/4: Checking registry hive..." -Color "Gray"
                try {
                    if ($isLocal) {
                        $hiveLoaded = Test-Path "Registry::HKEY_USERS\$($prof.SID)"
                        if ($hiveLoaded) {
                            & $script:WriteGuiOutput -Text "  Hive loaded - unloading HKU\$($prof.SID)..." -Color "Gray"
                            $null = & reg.exe unload "HKU\$($prof.SID)" 2>&1
                            & $script:WriteGuiOutput -Text "  Hive unloaded." -Color "Gray"
                        }
                        else {
                            & $script:WriteGuiOutput -Text "  Hive not loaded, skipping." -Color "Gray"
                        }
                    }
                    else {
                        & $script:WriteGuiOutput -Text "  Remote target - hive unload skipped." -Color "Gray"
                    }
                } catch {
                    & $script:WriteGuiOutput -Text "  Warning: Could not unload hive: $_" -Color "Yellow"
                }
                
                # Step 2: Remove registry entry
                & $script:WriteGuiOutput -Text "  Step 2/4: Removing ProfileList registry key..." -Color "Gray"
                try {
                    if ($isLocal) {
                        $regKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$($prof.SID)"
                        if (Test-Path $regKey) {
                            Remove-Item -Path $regKey -Recurse -Force -ErrorAction Stop
                            & $script:WriteGuiOutput -Text "  Registry entry removed: $regKey" -Color "Gray"
                        }
                        else {
                            & $script:WriteGuiOutput -Text "  Registry key not found (already removed)." -Color "Gray"
                        }
                    }
                    else {
                        Invoke-Command -ComputerName $targetComputer -ScriptBlock {
                            param($sid)
                            $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid"
                            if (Test-Path $regPath) {
                                Remove-Item -Path $regPath -Recurse -Force -ErrorAction Stop
                            }
                        } -ArgumentList $prof.SID -ErrorAction Stop
                        & $script:WriteGuiOutput -Text "  Remote registry entry removed." -Color "Gray"
                    }
                }
                catch {
                    & $script:WriteGuiOutput -Text "  ERROR removing registry: $($_.Exception.Message)" -Color "Red"
                    $success = $false
                }
                
                # Step 3: Remove profile folder
                & $script:WriteGuiOutput -Text "  Step 3/4: Removing profile folder..." -Color "Gray"
                if ($prof.ProfilePath) {
                    try {
                        $folderPath = $prof.ProfilePath
                        if (-not $isLocal) {
                            $folderPath = "\\$targetComputer\" + ($prof.ProfilePath -replace ':', '$')
                        }
                        & $script:WriteGuiOutput -Text "  Target path: $folderPath" -Color "Gray"
                        if (Test-Path $folderPath) {
                            Remove-Item -Path $folderPath -Recurse -Force -ErrorAction Stop
                            & $script:WriteGuiOutput -Text "  Profile folder removed successfully." -Color "Gray"
                        }
                        else {
                            & $script:WriteGuiOutput -Text "  Profile folder not found (already removed or orphaned)." -Color "Gray"
                        }
                    }
                    catch {
                        & $script:WriteGuiOutput -Text "  ERROR removing folder: $($_.Exception.Message)" -Color "Red"
                        $success = $false
                    }
                }
                
                # Step 4: Remove local user account (so it disappears from Computer Management)
                & $script:WriteGuiOutput -Text "  Step 4/4: Checking for local user account..." -Color "Gray"
                try {
                    $shortName = if ($prof.UserName) { ($prof.UserName -split '\\')[-1] } else { $null }
                    if ($shortName -and $shortName -ne '(Unresolvable)') {
                        & $script:WriteGuiOutput -Text "  Looking up local account: $shortName" -Color "Gray"
                        if ($isLocal) {
                            $localUser = $null
                            try { $localUser = Get-LocalUser -Name $shortName -ErrorAction SilentlyContinue } catch {}
                            if ($localUser) {
                                Remove-LocalUser -Name $shortName -ErrorAction Stop
                                & $script:WriteGuiOutput -Text "  Local user account '$shortName' removed." -Color "Gray"
                            }
                            else {
                                & $script:WriteGuiOutput -Text "  No local account found for '$shortName' (domain account or already removed)." -Color "Gray"
                            }
                        }
                        else {
                            Invoke-Command -ComputerName $targetComputer -ScriptBlock {
                                param($name)
                                $u = $null
                                try { $u = Get-LocalUser -Name $name -ErrorAction SilentlyContinue } catch {}
                                if ($u) { Remove-LocalUser -Name $name -ErrorAction Stop }
                            } -ArgumentList $shortName -ErrorAction Stop
                            & $script:WriteGuiOutput -Text "  Remote user account '$shortName' removed." -Color "Gray"
                        }
                    }
                }
                catch {
                    & $script:WriteGuiOutput -Text "  Note: Could not remove user account: $($_.Exception.Message)" -Color "Yellow"
                }
                
                if ($success) {
                    $removedCount++
                    & $script:WriteGuiOutput -Text "  Successfully removed $($prof.UserName)." -Color "Green"
                }
                else {
                    $failedCount++
                    & $script:WriteGuiOutput -Text "  FAILED to fully remove $($prof.UserName)." -Color "Red"
                }
            }
            
            & $script:WriteGuiOutput -Text "---" -Color "Gray"
            & $script:WriteGuiOutput -Text "Force removal complete: $removedCount removed, $failedCount failed." -Color "Cyan"
            
            $script:guiState.Running = $false
            $controls['btnForceRemove'].IsEnabled = $true
            $controls['progressBar'].Visibility = "Collapsed"
            
            # Refresh the profile list
            & $script:RefreshProfileList -TargetComputer $targetComputer
            
            [System.Windows.MessageBox]::Show("Force removal complete.`n`nRemoved: $removedCount`nFailed: $failedCount", "Complete", "OK", "Information")
        })
        
        # Event Handler: Browse Buttons
        $controls['btnBrowseBackup'].Add_Click({
            & $script:WriteGuiOutput -Text '[Button] Browse Backup Path clicked' -Color 'Cyan'
            $folder = New-Object Windows.Forms.FolderBrowserDialog
            $folder.Description = "Select Backup Directory"
            if ($folder.ShowDialog() -eq "OK") {
                $controls['txtBackupPath'].Text = $folder.SelectedPath
                $controls['chkBackup'].IsChecked = $true
                & $script:WriteGuiOutput -Text "[Browse] Backup path set to: $($folder.SelectedPath)" -Color 'Gray'
            } else {
                & $script:WriteGuiOutput -Text '[Browse] Backup path selection cancelled' -Color 'Gray'
            }
        })
        
        $controls['btnBrowseLog'].Add_Click({
            & $script:WriteGuiOutput -Text '[Button] Browse Log Path clicked' -Color 'Cyan'
            $save = New-Object Windows.Forms.SaveFileDialog
            $save.Filter = "Log files (*.log)|*.log|Text files (*.txt)|*.txt|All files (*.*)|*.*"
            $save.FileName = "DelprofPS.log"
            if ($save.ShowDialog() -eq "OK") {
                $controls['txtLogPath'].Text = $save.FileName
                $controls['chkLogPath'].IsChecked = $true
                & $script:WriteGuiOutput -Text "[Browse] Log path set to: $($save.FileName)" -Color 'Gray'
            } else {
                & $script:WriteGuiOutput -Text '[Browse] Log path selection cancelled' -Color 'Gray'
            }
        })
        
        $controls['btnBrowseCSV'].Add_Click({
            & $script:WriteGuiOutput -Text '[Button] Browse CSV Path clicked' -Color 'Cyan'
            $save = New-Object Windows.Forms.SaveFileDialog
            $save.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
            $save.FileName = "DelprofPS_Results.csv"
            if ($save.ShowDialog() -eq "OK") {
                $controls['txtOutputPath'].Text = $save.FileName
                $controls['chkOutputCSV'].IsChecked = $true
                & $script:WriteGuiOutput -Text "[Browse] CSV output path set to: $($save.FileName)" -Color 'Gray'
            } else {
                & $script:WriteGuiOutput -Text '[Browse] CSV path selection cancelled' -Color 'Gray'
            }
        })
        
        $controls['btnBrowseHtml'].Add_Click({
            & $script:WriteGuiOutput -Text '[Button] Browse HTML Report Path clicked' -Color 'Cyan'
            $save = New-Object Windows.Forms.SaveFileDialog
            $save.Filter = "HTML files (*.html)|*.html|All files (*.*)|*.*"
            $save.FileName = "DelprofPS_Report.html"
            if ($save.ShowDialog() -eq "OK") {
                $controls['txtHtmlPath'].Text = $save.FileName
                $controls['chkHtmlReport'].IsChecked = $true
                & $script:WriteGuiOutput -Text "[Browse] HTML report path set to: $($save.FileName)" -Color 'Gray'
            } else {
                & $script:WriteGuiOutput -Text '[Browse] HTML report path selection cancelled' -Color 'Gray'
            }
        })
        
        # Event Handler: Load Config
        $controls['btnLoadConfig'].Add_Click({
            & $script:WriteGuiOutput -Text '[Button] Load Config clicked' -Color 'Cyan'
            $open = New-Object Windows.Forms.OpenFileDialog
            $open.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
            $open.Title = "Load DelprofPS Configuration"
            if ($open.ShowDialog() -eq "OK") {
                try {
                    $config = Get-Content $open.FileName -Raw | ConvertFrom-Json
                    & $script:ApplyConfig -config $config -source $open.FileName
                }
                catch {
                    & $script:WriteGuiOutput -Text "[Config] ERROR: Failed to load config: $($_.Exception.Message)" -Color 'Red'
                    [System.Windows.MessageBox]::Show("Failed to load configuration: $($_.Exception.Message)", "Error", "OK", "Error")
                }
            } else {
                & $script:WriteGuiOutput -Text '[Config] Load cancelled by user' -Color 'Gray'
            }
        })
        
        # Event Handler: Save Config
        $controls['btnSaveConfig'].Add_Click({
            & $script:WriteGuiOutput -Text '[Button] Save Config clicked' -Color 'Cyan'
            $save = New-Object Windows.Forms.SaveFileDialog
            $save.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
            $save.FileName = "DelprofPS.config.json"
            if ($save.ShowDialog() -eq "OK") {
                & $script:WriteGuiOutput -Text "[Config] Saving configuration to: $($save.FileName)" -Color 'Gray'
                try {
                    $config = [ordered]@{
                        # Connection
                        RemoteComputers = [bool]$controls['rbRemoteComputers'].IsChecked
                        ComputerList    = $controls['txtComputerList'].Text
                        
                        # Age / Type
                        DaysInactive    = [int]$controls['sldDaysInactive'].Value
                        AgeMethod       = $controls['cmbAgeMethod'].SelectedIndex
                        ProfileType     = $controls['cmbProfileType'].SelectedIndex
                        
                        # Pattern filters
                        Exclude         = @($controls['txtExclude'].Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                        Include         = @($controls['txtInclude'].Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                        
                        # Size filters
                        MinSize         = $controls['txtMinSize'].Text
                        MaxSize         = $controls['txtMaxSize'].Text
                        
                        # Operation mode
                        DeleteMode      = [bool]$controls['rbModeDelete'].IsChecked
                        
                        # Action checkboxes
                        IncludeCorrupted = [bool]$controls['chkIncludeCorrupted'].IsChecked
                        FixCorruption    = [bool]$controls['chkFixCorruption'].IsChecked
                        ShowSpace        = [bool]$controls['chkShowSpace'].IsChecked
                        Detailed         = [bool]$controls['chkDetailed'].IsChecked
                        UnloadHives      = [bool]$controls['chkUnloadHives'].IsChecked
                        IgnoreActive     = [bool]$controls['chkIgnoreActive'].IsChecked
                        IncludeSystem    = [bool]$controls['chkIncludeSystem'].IsChecked
                        IncludeSpecial   = [bool]$controls['chkIncludeSpecial'].IsChecked
                        Force            = [bool]$controls['chkForce'].IsChecked
                        Interactive      = [bool]$controls['chkInteractive'].IsChecked
                        Quiet            = [bool]$controls['chkQuiet'].IsChecked
                        TestMode         = [bool]$controls['chkTestMode'].IsChecked
                        
                        # Parallel
                        UseParallel      = [bool]$controls['chkUseParallel'].IsChecked
                        ThrottleLimit    = [int]$controls['sldThrottle'].Value
                        
                        # Output paths
                        LogPath          = if ($controls['chkLogPath'].IsChecked) { $controls['txtLogPath'].Text } else { '' }
                        OutputPath       = if ($controls['chkOutputCSV'].IsChecked) { $controls['txtOutputPath'].Text } else { '' }
                        HtmlReport       = if ($controls['chkHtmlReport'].IsChecked) { $controls['txtHtmlPath'].Text } else { '' }
                        BackupPath       = if ($controls['chkBackup'].IsChecked) { $controls['txtBackupPath'].Text } else { '' }
                        
                        # Email
                        SmtpServer       = $controls['txtSmtpServer'].Text
                        EmailTo          = $controls['txtEmailTo'].Text
                        EmailFrom        = $controls['txtEmailFrom'].Text
                    }
                    
                    $config | ConvertTo-Json -Depth 3 | Out-File $save.FileName -Encoding UTF8
                    & $script:WriteGuiOutput -Text "Configuration saved to $($save.FileName)" -Color "Green"
                }
                catch {
                    & $script:WriteGuiOutput -Text "[Config] ERROR: Failed to save config: $($_.Exception.Message)" -Color 'Red'
                    [System.Windows.MessageBox]::Show("Failed to save configuration: $($_.Exception.Message)", "Error", "OK", "Error")
                }
            } else {
                & $script:WriteGuiOutput -Text '[Config] Save cancelled by user' -Color 'Gray'
            }
        })
        
        # Event Handler: Clear Output
        $controls['btnClear'].Add_Click({
            $controls['txtOutput'].Clear()
            & $script:WriteGuiOutput -Text '[Button] Output cleared' -Color 'Gray'
        })
        
        # Event Handler: Stop
        $controls['btnStop'].Add_Click({
            $script:guiState.StopRequested = $true
            & $script:WriteGuiOutput -Text "Stop requested... waiting for current operation to complete..." -Color "Yellow"
            $controls['btnStop'].IsEnabled = $false
        })
        
        # Event Handler: Run
        $controls['btnRun'].Add_Click({
            & $script:WriteGuiOutput -Text '[Button] RUN clicked' -Color 'Cyan'
            if ($script:guiState.Running) {
                & $script:WriteGuiOutput -Text '[Run] Blocked - another operation is already running' -Color 'Yellow'
                return
            }
            
            $script:guiState.Running = $true
            $script:guiState.StopRequested = $false
            $controls['btnRun'].IsEnabled = $false
            $controls['btnStop'].IsEnabled = $true
            $controls['progressBar'].Visibility = "Visible"
            $controls['progressBar'].IsIndeterminate = $true
            
            # Build parameter splat
            $params = @{}
            
            # Computer selection
            if ($controls['rbRemoteComputers'].IsChecked -and $controls['txtComputerList'].Text) {
                $params['ComputerName'] = $controls['txtComputerList'].Text -split "[\r\n,]+" | ForEach-Object { $_.Trim() } | Where-Object { $_ }
            }
            else {
                $params['ComputerName'] = $env:COMPUTERNAME
            }
            
            # Basic parameters
            $params['DaysInactive'] = $controls['sldDaysInactive'].Value
            $params['AgeCalculation'] = @('NTUSER_DAT', 'ProfilePath', 'Registry', 'LastLogon', 'LastLogoff')[$controls['cmbAgeMethod'].SelectedIndex]
            $params['ProfileType'] = @('All', 'Local', 'Roaming', 'Temporary', 'Mandatory')[$controls['cmbProfileType'].SelectedIndex]
            
            # Filters
            if ($controls['txtInclude'].Text) { $params['Include'] = $controls['txtInclude'].Text -split ',' | ForEach-Object { $_.Trim() } }
            if ($controls['txtExclude'].Text) { $params['Exclude'] = $controls['txtExclude'].Text -split ',' | ForEach-Object { $_.Trim() } }
            if ($controls['txtMinSize'].Text -and [int]$controls['txtMinSize'].Text) { $params['MinProfileSizeMB'] = [long]$controls['txtMinSize'].Text }
            if ($controls['txtMaxSize'].Text -and [int]$controls['txtMaxSize'].Text) { $params['MaxProfileSizeMB'] = [long]$controls['txtMaxSize'].Text }
            
            # Switches
            if ($controls['chkIncludeCorrupted'].IsChecked) { $params['IncludeCorrupted'] = $true }
            if ($controls['chkShowSpace'].IsChecked) { $params['ShowSpace'] = $true }
            if ($controls['chkDetailed'].IsChecked) { $params['Detailed'] = $true }
            if ($controls['chkQuiet'].IsChecked) { $params['Quiet'] = $true }
            if ($controls['chkUnloadHives'].IsChecked) { $params['UnloadHives'] = $true }
            if ($controls['chkForce'].IsChecked) { $params['Force'] = $true }
            if ($controls['chkIgnoreActive'].IsChecked) { $params['IgnoreActiveSessions'] = $true }
            if ($controls['chkIncludeSystem'].IsChecked) { $params['IncludeSystemProfiles'] = $true }
            if ($controls['chkIncludeSpecial'].IsChecked) { $params['IncludeSpecialProfiles'] = $true }
            if ($controls['chkFixCorruption'].IsChecked) { $params['FixCorruption'] = $true }
            if ($controls['chkTestMode'].IsChecked) { $params['Test'] = $true }
            if ($controls['chkUseParallel'].IsChecked) { 
                $params['UseParallel'] = $true 
                $params['ThrottleLimit'] = $controls['sldThrottle'].Value
            }
            
            # Operation mode
            if ($controls['rbModePreview'].IsChecked) { $params['Preview'] = $true }
            if ($controls['rbModeDelete'].IsChecked) { $params['Delete'] = $true }
            if ($controls['chkInteractive'].IsChecked) { $params['Interactive'] = $true }
            
            # Paths
            if ($controls['chkBackup'].IsChecked -and $controls['txtBackupPath'].Text) { $params['BackupPath'] = $controls['txtBackupPath'].Text }
            if ($controls['chkLogPath'].IsChecked -and $controls['txtLogPath'].Text) { $params['LogPath'] = $controls['txtLogPath'].Text }
            if ($controls['chkOutputCSV'].IsChecked -and $controls['txtOutputPath'].Text) { $params['OutputPath'] = $controls['txtOutputPath'].Text }
            if ($controls['chkHtmlReport'].IsChecked -and $controls['txtHtmlPath'].Text) { $params['HtmlReport'] = $controls['txtHtmlPath'].Text }
            
            # Email
            if ($controls['txtSmtpServer'].Text) { $params['SmtpServer'] = $controls['txtSmtpServer'].Text }
            if ($controls['txtEmailTo'].Text) { $params['EmailTo'] = $controls['txtEmailTo'].Text }
            if ($controls['txtEmailFrom'].Text) { $params['EmailFrom'] = $controls['txtEmailFrom'].Text }
            
            & $script:WriteGuiOutput -Text '[Run] Building parameter set...' -Color 'Gray'
            & $script:WriteGuiOutput -Text "[Run] Target: $($params['ComputerName'] -join ', ')" -Color 'Gray'
            & $script:WriteGuiOutput -Text "[Run] DaysInactive=$($params['DaysInactive']), AgeMethod=$($params['AgeCalculation']), ProfileType=$($params['ProfileType'])" -Color 'Gray'
            if ($params['Include']) { & $script:WriteGuiOutput -Text "[Run] Include filter: $($params['Include'] -join ', ')" -Color 'Gray' }
            if ($params['Exclude']) { & $script:WriteGuiOutput -Text "[Run] Exclude filter: $($params['Exclude'] -join ', ')" -Color 'Gray' }
            if ($params['BackupPath']) { & $script:WriteGuiOutput -Text "[Run] Backup path: $($params['BackupPath'])" -Color 'Gray' }
            if ($params['LogPath']) { & $script:WriteGuiOutput -Text "[Run] Log path: $($params['LogPath'])" -Color 'Gray' }
            if ($params['OutputPath']) { & $script:WriteGuiOutput -Text "[Run] CSV output: $($params['OutputPath'])" -Color 'Gray' }
            if ($params['HtmlReport']) { & $script:WriteGuiOutput -Text "[Run] HTML report: $($params['HtmlReport'])" -Color 'Gray' }
            $modeStr = if ($params['Delete']) { 'DELETE' } elseif ($params['Preview']) { 'PREVIEW' } else { 'LIST' }
            & $script:WriteGuiOutput -Text "[Run] Mode: $modeStr" -Color $(if ($params['Delete']) { 'Red' } else { 'Gray' })
            
            # Log command to output
            $cmdParts = @('.\DelprofPS.ps1')
            foreach ($key in $params.Keys) {
                $value = $params[$key]
                if ($value -is [switch] -or $value -is [bool]) {
                    if ($value) { $cmdParts += "-$key" }
                }
                elseif ($value -is [array]) {
                    $cmdParts += "-$key `"$($value -join ', ')`""
                }
                else {
                    $cmdParts += "-$key `"$value`""
                }
            }
            & $script:WriteGuiOutput -Text '[Run] Executing command:' -Color 'Cyan'
            & $script:WriteGuiOutput -Text ($cmdParts -join ' ') -Color 'White'
            & $script:WriteGuiOutput -Text '[Run] Launching background runspace...' -Color 'Gray'
            & $script:WriteGuiOutput -Text '---' -Color 'Gray'
            
            # Run in background runspace to keep UI responsive
            $runspace = [runspacefactory]::CreateRunspace()
            $runspace.Open()
            $runspace.SessionStateProxy.SetVariable('params', $params)
            $runspace.SessionStateProxy.SetVariable('scriptDir', $script:scriptRoot)
            
            $powershell = [powershell]::Create().AddScript({
                try {
                    $mainScript = Join-Path $scriptDir 'DelprofPS.ps1'
                    & $mainScript @params
                }
                catch {
                    Write-Error "ERROR: $($_.Exception.Message)"
                }
            })
            
            $powershell.Runspace = $runspace
            
            # Use PSDataCollection for output so we can read it incrementally
            $outputCollection = New-Object System.Management.Automation.PSDataCollection[PSObject]
            $asyncResult = $powershell.BeginInvoke($outputCollection, $outputCollection)
            
            # Mutable state hashtable — reference type persists across ticks inside .GetNewClosure()
            $runPollState = @{ InfoIdx = 0; WarnIdx = 0; ErrIdx = 0; OutIdx = 0 }
            
            $controlsRef = $controls
            $paramsRef = $params
            $guiStateRef = $script:guiState
            
            # DispatcherTimer polls streams every 200ms without blocking the UI
            $runTimer = New-Object System.Windows.Threading.DispatcherTimer
            $script:activeRunTimer = $runTimer
            $runTimer.Interval = [TimeSpan]::FromMilliseconds(200)
            $runTimer.Add_Tick({
                try {
                    # Guard against stale timer from a previous GUI session
                    if ($null -eq $powershell -or $null -eq $asyncResult) { $runTimer.Stop(); return }
                    
                    # Drain new output objects
                    while ($runPollState.OutIdx -lt $outputCollection.Count) {
                        $item = $outputCollection[$runPollState.OutIdx]
                        if ($item) {
                            $controlsRef['txtOutput'].AppendText("[$(Get-Date -Format 'HH:mm:ss')] $item`r`n")
                            $controlsRef['txtOutput'].ScrollToEnd()
                        }
                        $runPollState.OutIdx++
                    }
                    
                    # Drain Information stream (Write-Host output goes here in PS5.1 runspaces)
                    while ($runPollState.InfoIdx -lt $powershell.Streams.Information.Count) {
                        $info = $powershell.Streams.Information[$runPollState.InfoIdx]
                        $controlsRef['txtOutput'].AppendText("[$(Get-Date -Format 'HH:mm:ss')] $($info.MessageData)`r`n")
                        $controlsRef['txtOutput'].ScrollToEnd()
                        $runPollState.InfoIdx++
                    }
                    
                    # Drain Warning stream
                    while ($runPollState.WarnIdx -lt $powershell.Streams.Warning.Count) {
                        $warn = $powershell.Streams.Warning[$runPollState.WarnIdx]
                        $controlsRef['txtOutput'].AppendText("[$(Get-Date -Format 'HH:mm:ss')] [WARNING] $($warn.Message)`r`n")
                        $controlsRef['txtOutput'].ScrollToEnd()
                        $runPollState.WarnIdx++
                    }
                    
                    # Drain Error stream
                    while ($runPollState.ErrIdx -lt $powershell.Streams.Error.Count) {
                        $err = $powershell.Streams.Error[$runPollState.ErrIdx]
                        $controlsRef['txtOutput'].AppendText("[$(Get-Date -Format 'HH:mm:ss')] [ERROR] $($err.Exception.Message)`r`n")
                        $controlsRef['txtOutput'].ScrollToEnd()
                        $runPollState.ErrIdx++
                    }
                    
                    # Check for stop request
                    if ($guiStateRef.StopRequested) {
                        $powershell.Stop()
                    }
                    
                    # When completed, clean up
                    if ($asyncResult.IsCompleted) {
                        $runTimer.Stop()
                        $script:activeRunTimer = $null
                        
                        try {
                            $powershell.EndInvoke($asyncResult)
                        }
                        catch {
                            $controlsRef['txtOutput'].AppendText("[$(Get-Date -Format 'HH:mm:ss')] Operation stopped or encountered an error.`r`n")
                            $controlsRef['txtOutput'].ScrollToEnd()
                        }
                        finally {
                            $powershell.Dispose()
                            $runspace.Close()
                            $runspace.Dispose()
                        }
                        
                        $guiStateRef.Running = $false
                        $controlsRef['btnRun'].IsEnabled = $true
                        $controlsRef['btnStop'].IsEnabled = $false
                        $controlsRef['progressBar'].Visibility = "Collapsed"
                        
                        $controlsRef['txtOutput'].AppendText("[$(Get-Date -Format 'HH:mm:ss')] ---`r`n")
                        $controlsRef['txtOutput'].AppendText("[$(Get-Date -Format 'HH:mm:ss')] Operation completed.`r`n")
                        $controlsRef['txtOutput'].ScrollToEnd()
                        
                        if (-not $paramsRef['Quiet']) {
                            [System.Windows.MessageBox]::Show("Profile management operation completed!", "Complete", "OK", "Information")
                        }
                    }
                }
                catch {
                    # Log the error visibly, then clean up
                    try {
                        $controlsRef['txtOutput'].AppendText("[$(Get-Date -Format 'HH:mm:ss')] [Timer Error] $($_.Exception.Message)`r`n")
                        $controlsRef['txtOutput'].ScrollToEnd()
                    } catch {}
                    try { $runTimer.Stop() } catch {}
                    $script:activeRunTimer = $null
                    $guiStateRef.Running = $false
                    $controlsRef['btnRun'].IsEnabled = $true
                    $controlsRef['btnStop'].IsEnabled = $false
                    $controlsRef['progressBar'].Visibility = "Collapsed"
                }
            }.GetNewClosure())
            $runTimer.Start()
        })
        
        # Clean up timers when window closes (prevents stale timers on re-run)
        $window.Add_Closed({
            & $script:WriteGuiOutput -Text '[Window] GUI window closing - cleaning up timers...' -Color 'Gray'
            if ($script:activeRefreshTimer) {
                try { $script:activeRefreshTimer.Stop() } catch {}
                $script:activeRefreshTimer = $null
            }
            if ($script:activeRunTimer) {
                try { $script:activeRunTimer.Stop() } catch {}
                $script:activeRunTimer = $null
            }
        })
        
        # Show Window
        $window.ShowDialog() | Out-Null
    }
    #endregion

    #region UI Mode Check
    if ($UI) {
        $script:UIMode = $true
        Show-DelprofPSGUI
        return
    }
    #endregion

    #region Initialization
    $script:UIMode = $false
    $script:StartTime = Get-Date
    $script:Version = '2.0.0'
    $script:RunId = [guid]::NewGuid().ToString('N')
    Write-Host "[INIT] Delprof2-PS v$($script:Version) starting at $($script:StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Gray
    Write-Host "[INIT] RunId: $($script:RunId)" -ForegroundColor Gray
    Write-Host "[INIT] Running as: $([Security.Principal.WindowsIdentity]::GetCurrent().Name)" -ForegroundColor Gray
    Write-Host "[INIT] PowerShell: $($PSVersionTable.PSVersion) | OS: $([Environment]::OSVersion.VersionString)" -ForegroundColor Gray
    $script:TotalProfilesProcessed = 0
    $script:TotalProfilesDeleted = 0
    $script:TotalSpaceFreed = 0
    $script:SuppressDeletion = $false
    $script:Results = [System.Collections.Generic.List[object]]::new()
    $script:ComputerQueue = [System.Collections.Generic.List[string]]::new()

    # Well-known SIDs to protect
    $script:ProtectedSIDs = @(
        'S-1-5-18',      # SYSTEM
        'S-1-5-19',      # LOCAL SERVICE
        'S-1-5-20',      # NETWORK SERVICE
        'S-1-5-21-%-500' # Built-in Administrator (domain)
    )

    $script:SystemProfileNames = @(
        'Default', 'Public', 'Default User', 'All Users',
        'systemprofile', 'LocalService', 'NetworkService',
        'Administrator', 'Guest'
    )

    # Store Force flag for per-cmdlet use (avoid global $ErrorActionPreference override)
    $script:ForceErrorAction = if ($Force) { 'SilentlyContinue' } else { 'Stop' }

    # Build credential splat for remote cmdlets (empty when using current identity)
    $script:CredentialSplat = @{}
    if ($Credential -ne [System.Management.Automation.PSCredential]::Empty) {
        $script:CredentialSplat = @{ Credential = $Credential }
        Write-Host "[INIT] Using explicit credential for $($Credential.UserName)" -ForegroundColor Gray
    } else {
        Write-Host "[INIT] Using current identity for authentication" -ForegroundColor Gray
    }

    #region Security - Script Integrity Verification
    if ($VerifyIntegrity) {
        Write-Host "[INTEGRITY] Verifying script integrity..." -ForegroundColor Gray
        $hashFile = Join-Path $PSScriptRoot 'DelprofPS.sha256'
        if (Test-Path $hashFile) {
            $expectedHash = (Get-Content $hashFile -Raw).Trim().Split()[0]
            $actualHash = (Get-FileHash -Path $PSCommandPath -Algorithm SHA256).Hash
            if ($expectedHash -eq $actualHash) {
                Write-Host "[INTEGRITY] Script hash verified successfully." -ForegroundColor Green
            }
            else {
                Write-Host "[INTEGRITY] WARNING: Script hash mismatch!" -ForegroundColor Red
                Write-Host "  Expected: $expectedHash" -ForegroundColor Yellow
                Write-Host "  Actual:   $actualHash" -ForegroundColor Yellow
                Write-Host "  The script may have been tampered with." -ForegroundColor Red
                if (-not $Force) {
                    Write-Host "  Use -Force to continue anyway, or verify the script source." -ForegroundColor Yellow
                    exit 1
                }
            }
        }
        else {
            Write-Host "[INTEGRITY] No hash file found at $hashFile" -ForegroundColor Yellow
            Write-Host "  Generate one with: (Get-FileHash .\delprofPS.ps1 -Algorithm SHA256).Hash | Out-File .\DelprofPS.sha256" -ForegroundColor Gray
        }
    }
    #endregion

    #region Security - ComputerName Input Sanitisation
    Write-Host "[SECURITY] Validating computer name inputs..." -ForegroundColor Gray
    $validHostnamePattern = '^[a-zA-Z0-9][a-zA-Z0-9\-\.]{0,253}[a-zA-Z0-9]$|^[a-zA-Z0-9]$|^localhost$|^\.$'
    $sanitisedComputers = [System.Collections.Generic.List[string]]::new()
    foreach ($computer in $ComputerName) {
        $trimmed = $computer.Trim()
        if ($trimmed -match $validHostnamePattern) {
            $sanitisedComputers.Add($trimmed)
        }
        else {
            Write-Host "[SECURITY] Rejected invalid computer name: '$trimmed' - does not match RFC hostname pattern" -ForegroundColor Red
        }
    }
    if ($sanitisedComputers.Count -eq 0) {
        Write-Host "[SECURITY] No valid computer names provided. Exiting." -ForegroundColor Red
        exit 1
    }
    $ComputerName = $sanitisedComputers.ToArray()
    Write-Host "[SECURITY] $($sanitisedComputers.Count) valid computer name(s): $($ComputerName -join ', ')" -ForegroundColor Gray
    #endregion

    #region Security - Config Schema Validation & Loading
    if ($ConfigFile -and (Test-Path $ConfigFile)) {
        Write-Host "[CONFIG] Loading configuration from: $ConfigFile" -ForegroundColor Gray
        try {
            $config = Get-Content $ConfigFile -Raw | ConvertFrom-Json

            # Schema validation: check types and ranges
            $configValid = $true
            $configWarnings = @()

            if ($null -ne $config.DaysInactive) {
                if ($config.DaysInactive -is [int] -or $config.DaysInactive -is [long] -or $config.DaysInactive -is [double]) {
                    if ($config.DaysInactive -lt 0 -or $config.DaysInactive -gt 3650) {
                        $configWarnings += "DaysInactive ($($config.DaysInactive)) out of valid range 0-3650, ignoring"
                        $configValid = $false
                    }
                }
                else {
                    $configWarnings += "DaysInactive is not a number, ignoring"
                    $configValid = $false
                }
            }

            if ($null -ne $config.Exclude) {
                if ($config.Exclude -isnot [System.Array] -and $config.Exclude -isnot [string]) {
                    $configWarnings += "Exclude must be an array of strings, ignoring"
                }
            }

            if ($null -ne $config.Include) {
                if ($config.Include -isnot [System.Array] -and $config.Include -isnot [string]) {
                    $configWarnings += "Include must be an array of strings, ignoring"
                }
            }

            if ($null -ne $config.MaxRetries) {
                if ($config.MaxRetries -is [int] -or $config.MaxRetries -is [long] -or $config.MaxRetries -is [double]) {
                    if ($config.MaxRetries -lt 0 -or $config.MaxRetries -gt 50) {
                        $configWarnings += "MaxRetries ($($config.MaxRetries)) out of valid range 0-50, ignoring"
                    }
                }
                else {
                    $configWarnings += "MaxRetries is not a number, ignoring"
                }
            }

            # Reject unexpected/dangerous keys
            $allowedKeys = @('DaysInactive', 'Exclude', 'Include', 'MaxRetries', 'LogPath', 'OutputPath',
                             'HtmlReport', 'BackupPath', 'SmtpServer', 'EmailTo', 'EmailFrom',
                             'AgeCalculation', 'ProfileType', 'RetryDelaySeconds')
            $configKeys = $config.PSObject.Properties.Name
            foreach ($key in $configKeys) {
                if ($key -notin $allowedKeys) {
                    $configWarnings += "Unknown config key '$key' ignored (not in allowed schema)"
                }
            }

            # Report warnings
            foreach ($warn in $configWarnings) {
                Write-Host "[CONFIG] WARNING: $warn" -ForegroundColor Yellow
            }

            # Apply validated config values only if not already specified via parameters
            if (-not $PSBoundParameters.ContainsKey('DaysInactive') -and $null -ne $config.DaysInactive -and $configValid) {
                $DaysInactive = [int]$config.DaysInactive
            }
            if (-not $PSBoundParameters.ContainsKey('Exclude') -and $config.Exclude) { $Exclude = $config.Exclude }
            if (-not $PSBoundParameters.ContainsKey('Include') -and $config.Include) { $Include = $config.Include }
            if (-not $PSBoundParameters.ContainsKey('MaxRetries') -and $null -ne $config.MaxRetries -and
                $config.MaxRetries -ge 0 -and $config.MaxRetries -le 50) {
                $MaxRetries = [int]$config.MaxRetries
            }
            Write-Host "[CONFIG] Configuration loaded from $ConfigFile" -ForegroundColor Green
        }
        catch {
            Write-Host "[CONFIG] Failed to load config file: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    #endregion

    #region Path Validation
    # Validate output paths early to fail fast before processing
    $pathsToValidate = @{
        LogPath = $LogPath
        OutputPath = $OutputPath
        HtmlReport = $HtmlReport
        BackupPath = $BackupPath
    }
    foreach ($pathEntry in $pathsToValidate.GetEnumerator()) {
        if ($pathEntry.Value) {
            $parentDir = Split-Path $pathEntry.Value -Parent
            if ($parentDir -and -not (Test-Path $parentDir)) {
                try {
                    New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
                    Write-Host "[PATH] Created directory for $($pathEntry.Key): $parentDir" -ForegroundColor Gray
                }
                catch {
                    Write-Host "[PATH] ERROR: Cannot create directory for $($pathEntry.Key): $parentDir" -ForegroundColor Red
                    if (-not $Force) {
                        Write-Host "[PATH] Use -Force to continue anyway, or fix the path." -ForegroundColor Yellow
                        exit 1
                    }
                }
            }
        }
    }
    #endregion

    #region Logging Functions
    function Write-DPLog {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Message,
            
            [Parameter()]
            [ValidateSet('INFO', 'WARNING', 'ERROR', 'SUCCESS', 'DEBUG', 'VERBOSE')]
            [string]$Level = 'INFO'
        )
        
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = "[$timestamp] [$Level] $Message"
        
        # Write to log file if specified
        if ($LogPath) {
            try {
                Add-Content -Path $LogPath -Value $logEntry -ErrorAction SilentlyContinue
            }
            catch {
                # Silent fail for logging errors
            }
        }
        
        # Console output
        if (-not $Quiet) {
            switch ($Level) {
                'ERROR' { Write-Host $logEntry -ForegroundColor Red }
                'WARNING' { Write-Host $logEntry -ForegroundColor Yellow }
                'SUCCESS' { Write-Host $logEntry -ForegroundColor Green }
                'DEBUG' { Write-Host $logEntry -ForegroundColor DarkGray }
                'VERBOSE' { Write-Host $logEntry -ForegroundColor DarkGray }
                default { Write-Host $logEntry }
            }
        }
    }

    function Write-DPHeader {
        if (-not $Quiet) {
            Write-Host -Object ("`n" + ('=' * 80)) -ForegroundColor Cyan
            Write-Host -Object " Delprof2-PS v$script:Version - User Profile Management Tool" -ForegroundColor Cyan
            $modeText = if ($Preview) { 'PREVIEW/SIMULATION' } elseif ($Delete) { 'DELETE' } else { 'LIST/ANALYZE' }
            $modeColor = if ($Preview) { 'Magenta' } elseif ($Delete) { 'Red' } else { 'Green' }
            Write-Host " Mode: $modeText" -ForegroundColor $modeColor
            Write-Host " Criteria: Profiles older than $DaysInactive days" -ForegroundColor Cyan
            Write-Host -Object (('=' * 80) + "`n") -ForegroundColor Cyan
        }
        Write-DPLog -Message "Script started. Version: $script:Version, Delete mode: $Delete, Days inactive: $DaysInactive" -Level 'INFO'
    }
    #endregion

    #region Utility Functions
    function Test-AdminRight {
        $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }

    function ConvertTo-UserName {
        param([string]$SID)
        try {
            $securityIdentifier = New-Object System.Security.Principal.SecurityIdentifier($SID)
            $ntAccount = $securityIdentifier.Translate([System.Security.Principal.NTAccount])
            Write-DPLog -Message "    [SID] $SID -> $($ntAccount.Value)" -Level 'DEBUG'
            return $ntAccount.Value
        }
        catch {
            Write-DPLog -Message "    [SID] Failed to resolve $SID : $($_.Exception.Message)" -Level 'DEBUG'
            return $null
        }
    }

    function Test-IsProtectedProfile {
        param(
            [string]$UserName,
            [string]$SID
        )
        
        # Check system profile names
        foreach ($sysName in $script:SystemProfileNames) {
            if ($UserName -like "*\$sysName" -or $UserName -eq $sysName) {
                return $true
            }
        }
        
        # Check protected SIDs
        foreach ($protectedSid in $script:ProtectedSIDs) {
            if ($protectedSid -match '%') {
                # Pattern match for domain admin
                $pattern = $protectedSid -replace '%', '\d+'
                if ($SID -match $pattern) { return $true }
            } else {
                if ($SID -eq $protectedSid) { return $true }
            }
        }
        
        # Special accounts
        if (-not $IncludeSpecialProfiles) {
            if ($SID -in @('S-1-5-18', 'S-1-5-19', 'S-1-5-20')) {
                return $true
            }
        }
        
        return $false
    }

    function Format-Byte {
        param([long]$Bytes)
        
        if ($Bytes -lt 0) { return 'Error' }
        if ($Bytes -eq 0) { return '0 B' }
        
        $sizes = @('B', 'KB', 'MB', 'GB', 'TB')
        $order = [math]::Floor([math]::Log($Bytes, 1024))
        $order = [math]::Min($order, $sizes.Count - 1)
        
        $formatted = [math]::Round($Bytes / [math]::Pow(1024, $order), 2)
        return "$formatted $($sizes[$order])"
    }

    function Get-ProfileFolderSize {
        param([string]$Path)
        
        if (-not (Test-Path $Path)) {
            Write-DPLog -Message "    [Size] Path not found: $Path" -Level 'DEBUG'
            return 0
        }
        
        Write-DPLog -Message "    [Size] Calculating folder size for $Path via robocopy /L /XJ..." -Level 'DEBUG'
        try {
            $totalSize = [long]0
            $roboOut = & robocopy $Path 'C:\RobocopyNull' /L /E /BYTES /NJH /NC /NDL /NFL /XJ /R:0 /W:0 2>&1
            $roboText = ($roboOut | Out-String)
            if ($roboText -match 'Bytes\s*:\s*(\d+)') {
                $totalSize = [long]$Matches[1]
            }
            Write-DPLog -Message "    [Size] Result: $([math]::Round($totalSize / 1MB, 2)) MB ($totalSize bytes)" -Level 'DEBUG'
            return $totalSize
        }
        catch {
            Write-DPLog -Message "    [Size] Error calculating size for $Path : $($_.Exception.Message)" -Level 'WARNING'
            return -1
        }
    }

    function Get-ProfileSizeBreakdown {
        param([string]$ProfilePath)
        
        Write-DPLog -Message "    [Breakdown] Calculating folder breakdown for $ProfilePath" -Level 'DEBUG'
        $breakdown = @{}
        $folders = @('Documents', 'Downloads', 'Desktop', 'AppData', 'Pictures', 'Videos', 'Music')
        
        foreach ($folder in $folders) {
            $folderPath = Join-Path $ProfilePath $folder
            if (Test-Path $folderPath) {
                $size = Get-ProfileFolderSize -Path $folderPath
                $breakdown[$folder] = Format-Byte -Bytes $size
                Write-DPLog -Message "    [Breakdown] $folder = $($breakdown[$folder])" -Level 'DEBUG'
            }
            else {
                $breakdown[$folder] = 'N/A'
            }
        }
        
        return $breakdown
    }

    function Test-ProfileLockedFile {
        param([string]$ProfilePath)
        
        Write-DPLog -Message "    [LockedFiles] Checking for locked files in $ProfilePath" -Level 'DEBUG'
        $lockedFiles = @()
        try {
            $files = Get-ChildItem -Path $ProfilePath -Recurse -File -ErrorAction SilentlyContinue | Select-Object -First 10
            Write-DPLog -Message "    [LockedFiles] Sampling first $(@($files).Count) files" -Level 'DEBUG'
            foreach ($file in $files) {
                try {
                    $stream = [System.IO.File]::Open($file.FullName, 'Open', 'Read', 'None')
                    $stream.Close()
                }
                catch {
                    $lockedFiles += $file.FullName
                    Write-DPLog -Message "    [LockedFiles] LOCKED: $($file.FullName)" -Level 'DEBUG'
                }
            }
        }
        catch {
            Write-DPLog -Message "    [LockedFiles] Error checking files: $($_.Exception.Message)" -Level 'DEBUG'
        }
        Write-DPLog -Message "    [LockedFiles] Found $($lockedFiles.Count) locked file(s)" -Level 'DEBUG'
        return $lockedFiles
    }

    function Get-AgeColor {
        param([int]$AgeInDays)
        if ($AgeInDays -lt 30) { return 'Green' }
        elseif ($AgeInDays -lt 90) { return 'Yellow' }
        elseif ($AgeInDays -lt 180) { return 'Magenta' }
        else { return 'Red' }
    }

    function Backup-Profile {
        param(
            [string]$SourcePath,
            [string]$UserName
        )
        
        if (-not $BackupPath) {
            Write-DPLog -Message "  [Backup] No backup path configured, skipping backup for $UserName" -Level 'DEBUG'
            return $true
        }
        
        Write-DPLog -Message "  [Backup] Starting backup for $UserName from $SourcePath" -Level 'INFO'
        try {
            $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $backupFile = Join-Path $BackupPath "$($UserName)_$timestamp.zip"
            Write-DPLog -Message "  [Backup] Target: $backupFile" -Level 'DEBUG'
            
            if (-not (Test-Path $BackupPath)) {
                Write-DPLog -Message "  [Backup] Creating backup directory: $BackupPath" -Level 'DEBUG'
                New-Item -ItemType Directory -Path $BackupPath -Force | Out-Null
            }
            
            Write-DPLog -Message "  [Backup] Compressing $SourcePath..." -Level 'DEBUG'
            Compress-Archive -Path $SourcePath -DestinationPath $backupFile -CompressionLevel Optimal -Force
            $backupSize = (Get-Item $backupFile).Length
            Write-DPLog -Message "Profile backed up to $backupFile ($(Format-Byte -Bytes $backupSize))" -Level 'SUCCESS'
            return $true
        }
        catch {
            Write-DPLog -Message "Failed to backup profile $UserName : $($_.Exception.Message)" -Level 'ERROR'
            return $false
        }
    }

    function Send-NotificationEmail {
        param([hashtable]$Summary)
        
        if (-not $SmtpServer -or -not $EmailTo) {
            Write-DPLog -Message '[Email] No SMTP server or recipient configured, skipping notification' -Level 'DEBUG'
            return
        }
        
        Write-DPLog -Message "[Email] Sending notification to $EmailTo via $SmtpServer" -Level 'INFO'
        try {
            $subject = "Delprof2-PS Report - $($Summary.ProfilesDeleted) profiles deleted"
            $body = @"
Delprof2-PS has completed processing.

Summary:
- Computers processed: $($Summary.Computers)
- Profiles processed: $($Summary.ProfilesProcessed)
- Profiles deleted: $($Summary.ProfilesDeleted)
- Space freed: $($Summary.SpaceFreed)
- Duration: $($Summary.Duration)

Report generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
"@
            
            Send-MailMessage -SmtpServer $SmtpServer -To $EmailTo -From $EmailFrom -Subject $subject -Body $body
            Write-DPLog -Message "Notification email sent to $EmailTo" -Level 'SUCCESS'
        }
        catch {
            Write-DPLog -Message "Failed to send email: $($_.Exception.Message)" -Level 'ERROR'
        }
    }
    #endregion

    #region Event Logging
    function Write-EventLogEntry {
        param(
            [string]$Message,
            [ValidateSet('Information', 'Warning', 'Error')]
            [string]$EntryType = 'Information',
            [int]$EventId = 1000
        )
        
        try {
            $source = 'Delprof2PS'
            $logName = 'Application'
            
            # Create event source if it doesn't exist
            if (-not [System.Diagnostics.EventLog]::SourceExists($source)) {
                try {
                    New-EventLog -LogName $logName -Source $source -ErrorAction Stop
                }
                catch {
                    # May not have permission to create source
                    return
                }
            }
            
            Write-EventLog -LogName $logName -Source $source -EventId $EventId -EntryType $EntryType -Message $Message -ErrorAction SilentlyContinue
        }
        catch {
            # Silent fail - event logging is optional
        }
    }
    #endregion

    #region HTML Reporting
    function Export-HtmlReport {
        param(
            [string]$Path,
            [array]$Results,
            [hashtable]$Summary
        )
        
        $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Delprof2-PS Report</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 1400px; margin: 0 auto; background: white; padding: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #0078d4; border-bottom: 2px solid #0078d4; padding-bottom: 10px; }
        h2 { color: #333; margin-top: 30px; }
        .summary { background: #f0f8ff; padding: 15px; border-radius: 5px; margin: 20px 0; }
        .summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; }
        .summary-item { background: white; padding: 10px; border-radius: 3px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
        .summary-label { font-weight: bold; color: #666; font-size: 0.9em; }
        .summary-value { font-size: 1.2em; color: #0078d4; margin-top: 5px; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th { background: #0078d4; color: white; padding: 12px; text-align: left; }
        td { padding: 10px; border-bottom: 1px solid #ddd; }
        tr:hover { background: #f5f5f5; }
        .deleted { background-color: #d4edda; }
        .active { background-color: #fff3cd; }
        .error { background-color: #f8d7da; }
        .badge { padding: 3px 8px; border-radius: 3px; font-size: 0.85em; font-weight: bold; }
        .badge-local { background: #e3f2fd; color: #1565c0; }
        .badge-roaming { background: #f3e5f5; color: #7b1fa2; }
        .badge-temp { background: #fff3e0; color: #e65100; }
        .badge-mandatory { background: #fce4ec; color: #c2185b; }
        .badge-corrupted { background: #ffebee; color: #c62828; }
        .footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; color: #666; font-size: 0.9em; }
        .chart-container { margin: 20px 0; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Delprof2-PS Report</h1>
        <p>Generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        
        <div class="summary">
            <h2>Summary</h2>
            <div class="summary-grid">
                <div class="summary-item">
                    <div class="summary-label">Computers Processed</div>
                    <div class="summary-value">$($Summary.Computers)</div>
                </div>
                <div class="summary-item">
                    <div class="summary-label">Profiles Processed</div>
                    <div class="summary-value">$($Summary.ProfilesProcessed)</div>
                </div>
                <div class="summary-item">
                    <div class="summary-label">Profiles Deleted</div>
                    <div class="summary-value">$($Summary.ProfilesDeleted)</div>
                </div>
                <div class="summary-item">
                    <div class="summary-label">Space Freed</div>
                    <div class="summary-value">$($Summary.SpaceFreed)</div>
                </div>
                <div class="summary-item">
                    <div class="summary-label">Duration</div>
                    <div class="summary-value">$($Summary.Duration)</div>
                </div>
            </div>
        </div>
        
        <h2>Profile Details</h2>
        <table>
            <thead>
                <tr>
                    <th>Computer</th>
                    <th>User</th>
                    <th>Profile Type</th>
                    <th>Last Used</th>
                    <th>Age (Days)</th>
                    <th>Size</th>
                    <th>Status</th>
                </tr>
            </thead>
            <tbody>
"@
        
        foreach ($result in $Results) {
            $rowClass = ''
            if ($result.Deleted) { $rowClass = 'deleted' }
            elseif ($result.IsActiveSession) { $rowClass = 'active' }
            elseif ($result.Error) { $rowClass = 'error' }
            
            $typeClass = switch ($result.ProfileType) {
                'Local' { 'badge-local' }
                'Roaming' { 'badge-roaming' }
                'Temporary' { 'badge-temp' }
                'Mandatory' { 'badge-mandatory' }
                default { 'badge-corrupted' }
            }
            
            $status = if ($result.Deleted) { 'Deleted' } elseif ($result.IsActiveSession) { 'Active' } elseif ($result.Error) { 'Error' } else { 'Kept' }
            
            $html += "                <tr class='$rowClass'>" +
                "<td>$($result.ComputerName)</td>" +
                "<td>$($result.UserName)</td>" +
                "<td><span class='badge $typeClass'>$($result.ProfileType)</span></td>" +
                "<td>$($result.LastUsed)</td>" +
                "<td>$($result.AgeInDays)</td>" +
                "<td>$($result.SizeFormatted)</td>" +
                "<td>$status</td></tr>`n"
        }
        
        $html += @"
            </tbody>
        </table>
        
        <div class="footer">
            <p>Report generated by Delprof2-PS v$script:Version</p>
        </div>
    </div>
</body>
</html>
"@
        
        try {
            $html | Out-File -FilePath $Path -Encoding UTF8 -Force
            Write-DPLog -Message "HTML report saved to $Path" -Level 'SUCCESS'
            Write-EventLogEntry -Message "HTML report generated: $Path" -EntryType Information -EventId 1001
        }
        catch {
            Write-DPLog -Message "Failed to save HTML report: $($_.Exception.Message)" -Level 'ERROR'
        }
    }
    #endregion

    #region Interactive Mode
    function Select-ProfilesInteractive {
        param([array]$Profiles)
        
        if (-not $Profiles) { return @() }
        
        Write-Host "`n=== INTERACTIVE PROFILE SELECTION ===" -ForegroundColor Cyan
        Write-Host "Use arrow keys to navigate, Space to toggle selection, Enter to confirm`n" -ForegroundColor Gray
        
        $selected = New-Object System.Collections.Generic.List[int]
        $currentIndex = 0
        
        function Show-Menu {
            Clear-Host
            Write-Host "=== SELECT PROFILES TO DELETE ===" -ForegroundColor Cyan
            Write-Host "[Space] Select/Deselect  [Enter] Confirm  [A] Select All  [N] Select None  [Q] Quit`n" -ForegroundColor Gray
            
            for ($i = 0; $i -lt $Profiles.Count; $i++) {
                $prof = $Profiles[$i]
                $prefix = if ($i -eq $currentIndex) { '>' } else { ' ' }
                $marker = if ($selected -contains $i) { '[X]' } else { '[ ]' }
                $color = if ($i -eq $currentIndex) { 'Yellow' } elseif ($prof.IsActiveSession) { 'Red' } else { 'White' }
                $active = if ($prof.IsActiveSession) { ' [ACTIVE]' } else { '' }
                
                Write-Host " $prefix $marker $($prof.UserName) - $($prof.AgeInDays) days - $($prof.SizeFormatted)$active" -ForegroundColor $color
            }
            
            Write-Host "`nSelected: $($selected.Count) profiles ($(($Profiles | Where-Object { $selected -contains $Profiles.IndexOf($_) } | Measure-Object -Property SizeBytes -Sum).Sum / 1MB -as [int]) MB total)" -ForegroundColor Green
        }
        
        do {
            Show-Menu
            $key = [Console]::ReadKey($true)
            
            switch ($key.Key) {
                'UpArrow' { if ($currentIndex -gt 0) { $currentIndex-- } }
                'DownArrow' { if ($currentIndex -lt $Profiles.Count - 1) { $currentIndex++ } }
                'Spacebar' {
                    if ($selected -contains $currentIndex) {
                        $selected.Remove($currentIndex)
                    } else {
                        $selected.Add($currentIndex)
                    }
                }
                'A' { $selected = [System.Collections.Generic.List[int]](0..($Profiles.Count - 1)) }
                'N' { $selected.Clear() }
                'Q' { return @() }
            }
        } while ($key.Key -ne 'Enter')
        
        return $selected | ForEach-Object { $Profiles[$_] }
    }
    #endregion
    function Test-ComputerConnection {
        param([string]$Computer)
        
        Write-DPLog -Message "  [ConnTest] Pinging $Computer..." -Level 'DEBUG'
        try {
            # Test connection
            $ping = Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction SilentlyContinue
            if (-not $ping) {
                Write-DPLog -Message "  [ConnTest] $Computer is unreachable" -Level 'DEBUG'
                return @{ Success = $false; Error = 'Host unreachable' }
            }
            Write-DPLog -Message "  [ConnTest] Ping OK. Testing admin share \\$Computer\ADMIN`$..." -Level 'DEBUG'
            
            # Test admin access
            $testPath = "\\$Computer\ADMIN`$"
            if (-not (Test-Path $testPath -ErrorAction SilentlyContinue)) {
                Write-DPLog -Message "  [ConnTest] Admin share not accessible on $Computer" -Level 'DEBUG'
                return @{ Success = $false; Error = 'Admin share not accessible' }
            }
            
            Write-DPLog -Message "  [ConnTest] $Computer connectivity OK" -Level 'DEBUG'
            return @{ Success = $true; Error = $null }
        }
        catch {
            Write-DPLog -Message "  [ConnTest] Exception testing $Computer`: $($_.Exception.Message)" -Level 'DEBUG'
            return @{ Success = $false; Error = $_.Exception.Message }
        }
    }

    function Invoke-RemoteCommand {
        param(
            [string]$ComputerName,
            [scriptblock]$ScriptBlock,
            [hashtable]$ArgumentList = @{}
        )
        
        $isLocal = ($ComputerName -eq $env:COMPUTERNAME -or $ComputerName -eq 'localhost' -or $ComputerName -eq '.')
        Write-DPLog -Message "  [RemoteCmd] Invoking command on $ComputerName (local=$isLocal)" -Level 'DEBUG'
        try {
            if ($isLocal) {
                Write-DPLog -Message "  [RemoteCmd] Executing locally..." -Level 'DEBUG'
                $result = & $ScriptBlock @ArgumentList
                Write-DPLog -Message "  [RemoteCmd] Local execution completed successfully" -Level 'DEBUG'
                return @{ Success = $true; Data = $result; Error = $null }
            }
            else {
                Write-DPLog -Message "  [RemoteCmd] Creating PS session to $ComputerName..." -Level 'DEBUG'
                $session = New-PSSession -ComputerName $ComputerName -ErrorAction Stop
                Write-DPLog -Message "  [RemoteCmd] Session established, executing command..." -Level 'DEBUG'
                $result = Invoke-Command -Session $session -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList
                Remove-PSSession -Session $session
                Write-DPLog -Message "  [RemoteCmd] Remote execution completed successfully" -Level 'DEBUG'
                return @{ Success = $true; Data = $result; Error = $null }
            }
        }
        catch {
            Write-DPLog -Message "  [RemoteCmd] FAILED on $ComputerName : $($_.Exception.Message)" -Level 'WARNING'
            return @{ Success = $false; Data = $null; Error = $_.Exception.Message }
        }
    }
    #endregion

    #region Profile Analysis Functions
    function Get-ActiveSession {
        param([string]$ComputerName)
        
        Write-DPLog -Message "  [Sessions] Detecting active sessions on $ComputerName..." -Level 'DEBUG'
        try {
            $sessions = @()
            
            # Method 1: Query user.exe (quser)
            Write-DPLog -Message "  [Sessions] Method 1: quser /server:$ComputerName" -Level 'DEBUG'
            try {
                $quserOutput = quser /server:$ComputerName 2>$null
                if ($quserOutput) {
                    $quserSessions = $quserOutput | Select-String -Pattern '(\S+)\s+\d+\s+(\S+)' | ForEach-Object {
                        $matches[1]
                    }
                    $sessions += $quserSessions
                    Write-DPLog -Message "  [Sessions] quser found $($quserSessions.Count) session(s)" -Level 'DEBUG'
                }
                else {
                    Write-DPLog -Message "  [Sessions] quser returned no sessions" -Level 'DEBUG'
                }
            }
            catch {
                Write-DPLog -Message "  [Sessions] quser not available or failed" -Level 'DEBUG'
            }
            
            # Method 2: Get interactive logon sessions via WMI Win32_LogonSession
            #   (Win32_LoggedOnUser returns loaded hives, NOT actual sessions - too many false positives)
            Write-DPLog -Message "  [Sessions] Method 2: WMI Win32_LogonSession (Interactive/RemoteInteractive)" -Level 'DEBUG'
            try {
                $logonSessions = Get-WmiObject -Class Win32_LogonSession -ComputerName $ComputerName -ErrorAction SilentlyContinue @script:CredentialSplat |
                    Where-Object { $_.LogonType -eq 2 -or $_.LogonType -eq 10 -or $_.LogonType -eq 11 }  # 2=Interactive, 10=RemoteInteractive, 11=CachedInteractive
                if ($logonSessions) {
                    $wmiSessions = @()
                    foreach ($ls in $logonSessions) {
                        $related = Get-WmiObject -Class Win32_LoggedOnUser -ComputerName $ComputerName -ErrorAction SilentlyContinue @script:CredentialSplat |
                            Where-Object { $_.Dependent -like "*LogonId=`"$($ls.LogonId)`"*" } |
                            ForEach-Object {
                                $_.Antecedent -match 'Domain="([^"]+)",Name="([^"]+)"' | Out-Null
                                "$($matches[1])\$($matches[2])"
                            }
                        $wmiSessions += $related
                    }
                    $wmiSessions = $wmiSessions | Select-Object -Unique
                    $sessions += $wmiSessions
                    Write-DPLog -Message "  [Sessions] WMI interactive sessions found $(@($wmiSessions).Count) user(s)" -Level 'DEBUG'
                }
                else {
                    Write-DPLog -Message "  [Sessions] WMI found no interactive logon sessions" -Level 'DEBUG'
                }
            }
            catch {
                Write-DPLog -Message "  [Sessions] WMI LogonSession query failed: $($_.Exception.Message)" -Level 'DEBUG'
            }
            
            # Method 3: Get explorer.exe processes
            Write-DPLog -Message "  [Sessions] Method 3: explorer.exe process owners" -Level 'DEBUG'
            try {
                $explorerUsers = Get-WmiObject -Class Win32_Process -ComputerName $ComputerName -Filter "Name='explorer.exe'" -ErrorAction SilentlyContinue |
                    ForEach-Object { 
                        $owner = $_.GetOwner()
                        if ($owner) { "$($owner.Domain)\$($owner.User)" }
                    } | Select-Object -Unique
                $sessions += $explorerUsers
                Write-DPLog -Message "  [Sessions] Explorer.exe found $(@($explorerUsers).Count) user(s)" -Level 'DEBUG'
            }
            catch {
                Write-DPLog -Message "  [Sessions] Explorer.exe query failed" -Level 'DEBUG'
            }
            
            $uniqueSessions = $sessions | Select-Object -Unique
            Write-DPLog -Message "  [Sessions] Total unique active sessions: $(@($uniqueSessions).Count) - $($uniqueSessions -join ', ')" -Level 'DEBUG'
            return $uniqueSessions
        }
        catch {
            Write-DPLog -Message "  [Sessions] Exception: $($_.Exception.Message)" -Level 'DEBUG'
            return @()
        }
    }

    function Get-ProfileAge {
        param(
            [string]$ProfilePath,
            [string]$SID,
            [string]$Method,
            [string]$ComputerName
        )
        
        Write-DPLog -Message "    [Age] Calculating age for SID $SID using method '$Method'" -Level 'DEBUG'
        $lastUsed = $null
        $source = 'Unknown'
        
        switch ($Method) {
            'NTUSER_DAT' {
                $ntUserDat = Join-Path $ProfilePath 'NTUSER.DAT'
                if (Test-Path $ntUserDat -ErrorAction SilentlyContinue) {
                    try {
                        $lastUsed = (Get-Item $ntUserDat -Force).LastWriteTime
                        $source = 'NTUSER.DAT'
                    }
                    catch {
                        $lastUsed = $null
                    }
                }
                
                if (-not $lastUsed) {
                    # Fallback to profile path
                    try {
                        $lastUsed = (Get-Item $ProfilePath).LastWriteTime
                        $source = 'ProfilePath'
                    }
                    catch {
                        $lastUsed = [DateTime]::MinValue
                        $source = 'Error'
                    }
                }
            }
            
            'ProfilePath' {
                if (Test-Path $ProfilePath -ErrorAction SilentlyContinue) {
                    try {
                        $lastUsed = (Get-Item $ProfilePath).LastWriteTime
                        $source = 'ProfilePath'
                    }
                    catch {
                        $lastUsed = [DateTime]::MinValue
                        $source = 'Error'
                    }
                }
            }
            
            'Registry' {
                try {
                    $profileKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$SID"
                    if (Test-Path $profileKey) {
                        $profileInfo = Get-ItemProperty $profileKey
                        # Try to find timestamp in various registry values
                        if ($profileInfo.LocalProfileLoadTimeHigh) {
                            # Convert FILETIME if available
                            $ftLow = [uint32]$profileInfo.LocalProfileLoadTimeLow
                            $ftHigh = [uint32]$profileInfo.LocalProfileLoadTimeHigh
                            $lastUsed = [DateTime]::FromFileTime([long]$ftLow + ([long]$ftHigh -shl 32))
                            $source = 'RegistryLoadTime'
                        }
                        else {
                            $lastUsed = (Get-Item $profileKey).LastWriteTime
                            $source = 'RegistryKey'
                        }
                    }
                }
                catch {
                    $lastUsed = [DateTime]::MinValue
                    $source = 'Error'
                }
            }
            
            'LastLogon' {
                try {
                    $userName = ConvertTo-UserName -SID $SID
                    if ($userName) {
                        $adUser = [ADSI]"WinNT://$($userName -replace '\\','/'),user"
                        if ($adUser.LastLogin) {
                            $lastUsed = $adUser.LastLogin
                            $source = 'LastLogon'
                        }
                    }
                }
                catch {
                    $lastUsed = $null
                }
                
                if (-not $lastUsed) {
                    # Fallback to NTUSER.DAT
                    $ntUserDat = Join-Path $ProfilePath 'NTUSER.DAT'
                    if (Test-Path $ntUserDat -ErrorAction SilentlyContinue) {
                        $lastUsed = (Get-Item $ntUserDat -Force).LastWriteTime
                        $source = 'NTUSER.DAT (fallback)'
                    }
                }
            }
        }
        
        if (-not $lastUsed) {
            $lastUsed = [DateTime]::MinValue
            $source = 'Unknown'
        }
        
        Write-DPLog -Message "    [Age] Result: LastUsed=$($lastUsed.ToString('yyyy-MM-dd HH:mm')), Source=$source" -Level 'DEBUG'
        return @{ LastUsed = $lastUsed; Source = $source }
    }

    function Get-ProfileType {
        param([hashtable]$ProfileInfo)
        
        $profilePath = $ProfileInfo.ProfilePath
        Write-DPLog -Message "    [Type] Determining profile type for $profilePath" -Level 'DEBUG'
        
        # Check for roaming
        if ($ProfileInfo.RoamingConfigured -eq 1) {
            Write-DPLog -Message "    [Type] Roaming flag set in registry" -Level 'DEBUG'
            return 'Roaming'
        }
        
        # Check for mandatory
        if ($profilePath -like '*.man' -or (Test-Path "$profilePath.man" -ErrorAction SilentlyContinue)) {
            Write-DPLog -Message "    [Type] Mandatory profile detected" -Level 'DEBUG'
            return 'Mandatory'
        }
        
        # Check for temporary
        if ($ProfileInfo.TemporaryProfile -eq 1 -or $profilePath -like '*TEMP*') {
            Write-DPLog -Message "    [Type] Temporary profile detected" -Level 'DEBUG'
            return 'Temporary'
        }
        
        # Check if corrupted
        if (-not (Test-Path $profilePath -ErrorAction SilentlyContinue)) {
            Write-DPLog -Message "    [Type] Profile path missing - corrupted" -Level 'DEBUG'
            return 'Corrupted (Path Missing)'
        }
        
        if (-not (Test-Path (Join-Path $profilePath 'NTUSER.DAT') -ErrorAction SilentlyContinue)) {
            Write-DPLog -Message "    [Type] NTUSER.DAT missing - corrupted" -Level 'DEBUG'
            return 'Corrupted (No NTUSER.DAT)'
        }
        
        Write-DPLog -Message "    [Type] Local profile" -Level 'DEBUG'
        return 'Local'
    }

    function Repair-CorruptedProfile {
        [CmdletBinding(SupportsShouldProcess = $true)]
        param(
            [string]$ComputerName,
            [string]$SID,
            [string]$UserName,
            [string]$ProfilePath,
            [string]$CorruptionType
        )
        
        $fixed = $false
        $actionTaken = 'No action taken'
        
        Write-DPLog -Message "Corruption detected for $UserName ($SID)`: $CorruptionType" -Level 'WARNING'
        
        if (-not $Quiet) {
            Write-Host "`n  CORRUPTION DETECTED" -ForegroundColor Red -BackgroundColor Black
            Write-Host "  User: $UserName" -ForegroundColor Yellow
            Write-Host "  SID: $SID" -ForegroundColor Yellow
            Write-Host "  Path: $ProfilePath" -ForegroundColor Yellow
            Write-Host "  Issue: $CorruptionType" -ForegroundColor Red
        }
        
        # Determine available repair options based on corruption type
        $repairOptions = @()
        
        switch ($CorruptionType) {
            'Corrupted (Path Missing)' {
                $repairOptions = @(
                    @{ Key = 'R'; Label = 'Remove orphaned registry key'; Description = 'Deletes the stale registry entry pointing to non-existent folder'; Risk = 'Low' }
                    @{ Key = 'S'; Label = 'Skip'; Description = 'Leave as-is and continue'; Risk = 'None' }
                )
            }
            'Corrupted (No NTUSER.DAT)' {
                $repairOptions = @(
                    @{ Key = 'D'; Label = 'Delete entire profile'; Description = 'Remove registry key and delete profile folder'; Risk = 'Medium' }
                    @{ Key = 'R'; Label = 'Recreate NTUSER.DAT'; Description = 'Copy default NTUSER.DAT to fix the profile'; Risk = 'Low' }
                    @{ Key = 'F'; Label = 'Force-remove folder only'; Description = 'Delete folder contents but keep registry key'; Risk = 'Medium' }
                    @{ Key = 'S'; Label = 'Skip'; Description = 'Leave as-is and continue'; Risk = 'None' }
                )
            }
        }
        
        # Display interactive menu
        if (-not $Quiet) {
            Write-Host "`n  Available repair options:" -ForegroundColor Cyan
            foreach ($opt in $repairOptions) {
                $riskColor = switch ($opt.Risk) {
                    'Low' { 'Green' }
                    'Medium' { 'Yellow' }
                    'High' { 'Red' }
                    default { 'White' }
                }
                Write-Host "    [$($opt.Key)] $($opt.Label)" -ForegroundColor White -NoNewline
                Write-Host " [Risk: $($opt.Risk)]" -ForegroundColor $riskColor
                Write-Host "        $($opt.Description)" -ForegroundColor Gray
            }
            Write-Host "`n  NOTE: This operation requires administrative privileges." -ForegroundColor Magenta
            Write-Host "  The user will NOT be able to log in until repaired." -ForegroundColor Magenta
        }
        
        # Get user choice (respecting Force parameter for non-interactive)
        $choice = $null
        if ($Force -and -not $Interactive) {
            # In Force mode without Interactive, default to Skip to prevent accidental damage
            $choice = 'S'
            if (-not $Quiet) {
                Write-Host "`n  Force mode active - defaulting to Skip. Use -FixCorruption with -Interactive for manual control." -ForegroundColor Yellow
            }
        }
        else {
            # Interactive prompt
            $validChoices = $repairOptions.Key
            while ($choice -notin $validChoices) {
                if (-not $Quiet) {
                    Write-Host "`n  Enter your choice [$(($validChoices -join '/'))]: " -ForegroundColor Cyan -NoNewline
                }
                $choice = Read-Host
                $choice = $choice.ToUpper().Trim()
            }
        }
        
        # Execute chosen action
        switch ($choice) {
            'R' { # Remove registry key or Recreate NTUSER.DAT
                if ($CorruptionType -eq 'Corrupted (Path Missing)') {
                    # Remove orphaned registry key
                    $targetDesc = "Remove orphaned registry key for $UserName ($SID)"
                    if ($PSCmdlet.ShouldProcess($targetDesc, 'Remove Registry Key')) {
                        try {
                            $regKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$SID"
                            if ($ComputerName -eq $env:COMPUTERNAME -or $ComputerName -eq 'localhost' -or $ComputerName -eq '.') {
                                if (Test-Path $regKey) {
                                    Remove-Item -Path $regKey -Recurse -Force
                                }
                            }
                            else {
                                $scriptBlock = {
                                    param($sid)
                                    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
                                    if (Test-Path "$regPath\$sid") {
                                        Remove-Item -Path "$regPath\$sid" -Recurse -Force
                                    }
                                }
                                Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $SID @script:CredentialSplat
                            }
                            $fixed = $true
                            $actionTaken = 'Removed orphaned registry key'
                            Write-DPLog -Message "Removed orphaned registry key for $UserName" -Level 'SUCCESS'
                        }
                        catch {
                            $actionTaken = "Failed to remove registry key: $($_.Exception.Message)"
                            Write-DPLog -Message "Failed to remove registry key for $UserName`: $($_.Exception.Message)" -Level 'ERROR'
                        }
                    }
                }
                else {
                    # Recreate NTUSER.DAT
                    $targetDesc = "Recreate NTUSER.DAT for $UserName at $ProfilePath"
                    if ($PSCmdlet.ShouldProcess($targetDesc, 'Recreate NTUSER.DAT')) {
                        try {
                            $defaultNtUser = "C:\Users\Default\NTUSER.DAT"
                            if (Test-Path $defaultNtUser) {
                                $targetNtUser = Join-Path $ProfilePath 'NTUSER.DAT'
                                Copy-Item -Path $defaultNtUser -Destination $targetNtUser -Force
                                
                                # Set proper permissions (simplified - full ACL would require more code)
                                $acl = Get-Acl $targetNtUser
                                $sidObject = New-Object System.Security.Principal.SecurityIdentifier($SID)
                                $rule = New-Object System.Security.AccessControl.FileSystemAccessRule($sidObject, 'FullControl', 'Allow')
                                $acl.SetAccessRule($rule)
                                Set-Acl $targetNtUser $acl
                                
                                $fixed = $true
                                $actionTaken = 'Recreated NTUSER.DAT from default'
                                Write-DPLog -Message "Recreated NTUSER.DAT for $UserName" -Level 'SUCCESS'
                            }
                            else {
                                $actionTaken = 'Default NTUSER.DAT not found'
                                Write-DPLog -Message "Default NTUSER.DAT not found for copying" -Level 'ERROR'
                            }
                        }
                        catch {
                            $actionTaken = "Failed to recreate NTUSER.DAT: $($_.Exception.Message)"
                            Write-DPLog -Message "Failed to recreate NTUSER.DAT for $UserName`: $($_.Exception.Message)" -Level 'ERROR'
                        }
                    }
                }
            }
            
            'D' { # Delete entire profile
                $targetDesc = "Delete entire corrupted profile for $UserName ($SID) at $ProfilePath"
                if ($PSCmdlet.ShouldProcess($targetDesc, 'Delete Corrupted Profile')) {
                    try {
                        # Remove registry key
                        $regKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$SID"
                        if ($ComputerName -eq $env:COMPUTERNAME -or $ComputerName -eq 'localhost' -or $ComputerName -eq '.') {
                            if (Test-Path $regKey) {
                                Remove-Item -Path $regKey -Recurse -Force
                            }
                            if (Test-Path $ProfilePath) {
                                Remove-Item -Path $ProfilePath -Recurse -Force
                            }
                        }
                        else {
                            $scriptBlock = {
                                param($sid, $profilePath)
                                $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
                                if (Test-Path "$regPath\$sid") {
                                    Remove-Item -Path "$regPath\$sid" -Recurse -Force
                                }
                                if (Test-Path $profilePath) {
                                    Remove-Item -Path $profilePath -Recurse -Force
                                }
                            }
                            Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $SID, $ProfilePath @script:CredentialSplat
                        }
                        $fixed = $true
                        $actionTaken = 'Deleted entire corrupted profile'
                        Write-DPLog -Message "Deleted corrupted profile for $UserName" -Level 'SUCCESS'
                    }
                    catch {
                        $actionTaken = "Failed to delete profile: $($_.Exception.Message)"
                        Write-DPLog -Message "Failed to delete corrupted profile for $UserName`: $($_.Exception.Message)" -Level 'ERROR'
                    }
                }
            }
            
            'F' { # Force-remove folder only
                $targetDesc = "Remove profile folder for $UserName at $ProfilePath"
                if ($PSCmdlet.ShouldProcess($targetDesc, 'Remove Folder')) {
                    try {
                        if ($ComputerName -eq $env:COMPUTERNAME -or $ComputerName -eq 'localhost' -or $ComputerName -eq '.') {
                            if (Test-Path $ProfilePath) {
                                # Remove read-only attributes
                                Get-ChildItem -Path $ProfilePath -Recurse -Force -ErrorAction SilentlyContinue | 
                                    ForEach-Object { $_.Attributes = $_.Attributes -band -bnot [System.IO.FileAttributes]::ReadOnly }
                                Remove-Item -Path $ProfilePath -Recurse -Force
                            }
                        }
                        else {
                            $scriptBlock = {
                                param($profilePath)
                                if (Test-Path $profilePath) {
                                    Get-ChildItem -Path $profilePath -Recurse -Force -ErrorAction SilentlyContinue | 
                                        ForEach-Object { $_.Attributes = $_.Attributes -band -bnot [System.IO.FileAttributes]::ReadOnly }
                                    Remove-Item -Path $profilePath -Recurse -Force
                                }
                            }
                            Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $ProfilePath @script:CredentialSplat
                        }
                        $fixed = $true
                        $actionTaken = 'Removed profile folder only'
                        Write-DPLog -Message "Removed profile folder for $UserName (registry key preserved)" -Level 'SUCCESS'
                    }
                    catch {
                        $actionTaken = "Failed to remove folder: $($_.Exception.Message)"
                        Write-DPLog -Message "Failed to remove folder for $UserName`: $($_.Exception.Message)" -Level 'ERROR'
                    }
                }
            }
            
            'S' { # Skip
                $actionTaken = 'Skipped by administrator'
                if (-not $Quiet) {
                    Write-Host "  Skipped corruption repair for $UserName" -ForegroundColor Yellow
                }
            }
        }
        
        # Audit log entry for corruption repair actions
        Write-EventLogEntry -Message "Corruption repair on $ComputerName for '$UserName' ($SID): Choice=$choice, Action=$actionTaken, Fixed=$fixed" -EntryType Information -EventId 1012
        
        return [PSCustomObject]@{
            Fixed = $fixed
            ActionTaken = $actionTaken
            Choice = $choice
        }
    }
    #endregion

    #region Core Profile Functions
    function Get-UserProfile {
        param([string]$ComputerName)
        
        Write-DPLog -Message "Scanning profiles on $ComputerName..." -Level 'INFO'
        Write-DPLog -Message "  [Registry] Opening ProfileList registry key..." -Level 'DEBUG'
        
        $profiles = @()
        
        try {
            $profileListPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
            
            if ($ComputerName -ne $env:COMPUTERNAME -and $ComputerName -ne 'localhost' -and $ComputerName -ne '.') {
                Write-DPLog -Message "  [Registry] Using WMI for remote registry access on $ComputerName" -Level 'DEBUG'
                # Use WMI for remote registry
                $profileKeys = Get-WmiObject -ComputerName $ComputerName -Class StdRegProv -Namespace 'root\default' -ErrorAction Stop @script:CredentialSplat |
                    ForEach-Object { $_.EnumKey(2147483650, 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList') } |
                    Select-Object -ExpandProperty sNames |
                    Where-Object { $_ -match '^S-1-5-21' }
                
                # Cache WMI connection outside loop to avoid reconnecting per-SID
                $regProv = Get-WmiObject -ComputerName $ComputerName -Class StdRegProv -Namespace 'root\default' -ErrorAction Stop @script:CredentialSplat
                
                Write-DPLog -Message "  [Registry] Found $(@($profileKeys).Count) SIDs on remote $ComputerName" -Level 'DEBUG'
                foreach ($sid in $profileKeys) {
                    try {
                        $profilePathValue = $regProv.GetStringValue(2147483650, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid", 'ProfileImagePath')
                        $roamingValue = $regProv.GetDWORDValue(2147483650, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid", 'RoamingConfigured')
                        $tempValue = $regProv.GetDWORDValue(2147483650, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid", 'TemporaryProfile')
                        
                        $profilePath = $profilePathValue.sValue
                        if ($profilePath) {
                            $profilePath = $profilePath -replace '%SystemDrive%', 'C:'
                        }
                        
                        Write-DPLog -Message "    [Registry] $sid -> $profilePath" -Level 'DEBUG'
                        $profiles += @{
                            SID = $sid
                            ProfilePath = $profilePath
                            RoamingConfigured = $roamingValue.uValue
                            TemporaryProfile = $tempValue.uValue
                        }
                    }
                    catch {
                        Write-DPLog -Message "Error reading profile $sid on $ComputerName`: $($_.Exception.Message)" -Level 'WARNING'
                    }
                }
            }
            else {
                Write-DPLog -Message "  [Registry] Using local registry access" -Level 'DEBUG'
                # Local registry access
                $profileKeys = Get-ChildItem $profileListPath -ErrorAction Stop | 
                    Where-Object { $_.PSChildName -match '^S-1-5-21' }
                
                Write-DPLog -Message "  [Registry] Found $(@($profileKeys).Count) profile SIDs locally" -Level 'DEBUG'
                foreach ($key in $profileKeys) {
                    try {
                        $props = Get-ItemProperty $key.PSPath
                        Write-DPLog -Message "    [Registry] $($key.PSChildName) -> $($props.ProfileImagePath)" -Level 'DEBUG'
                        $profiles += @{
                            SID = $key.PSChildName
                            ProfilePath = $props.ProfileImagePath
                            RoamingConfigured = $props.RoamingConfigured
                            TemporaryProfile = $props.TemporaryProfile
                        }
                    }
                    catch {
                        Write-DPLog -Message "Error reading profile $($key.PSChildName)`:`r`n$($_.Exception.Message)" -Level 'WARNING'
                    }
                }
            }
        }
        catch {
            Write-DPLog -Message "Failed to enumerate profiles on $ComputerName`: $($_.Exception.Message)" -Level 'ERROR'
            return $null
        }
        
        Write-DPLog -Message "  [Registry] Total profiles read: $($profiles.Count)" -Level 'DEBUG'
        return $profiles
    }

    function Test-ProfileFilter {
        param(
            [string]$UserName,
            [string]$SID,
            [string]$ProfilePath,
            [long]$ProfileSize,
            [string]$ActualProfileType
        )
        
        # Include filter
        if ($Include) {
            $includeMatch = $false
            foreach ($pattern in $Include) {
                if ($UserName -like $pattern) { $includeMatch = $true; break }
            }
            if (-not $includeMatch) {
                Write-DPLog -Message "    [Filter] $UserName excluded - does not match Include patterns: $($Include -join ', ')" -Level 'DEBUG'
                return $false
            }
            Write-DPLog -Message "    [Filter] $UserName matches Include pattern" -Level 'DEBUG'
        }
        
        # Exclude filter
        if ($Exclude) {
            foreach ($pattern in $Exclude) {
                if ($UserName -like $pattern) {
                    Write-DPLog -Message "    [Filter] $UserName excluded - matches Exclude pattern '$pattern'" -Level 'DEBUG'
                    return $false
                }
            }
        }
        
        # Size filters
        if ($MinProfileSizeMB -and $ProfileSize -lt ($MinProfileSizeMB * 1MB)) {
            Write-DPLog -Message "    [Filter] $UserName excluded - size $([math]::Round($ProfileSize/1MB,1))MB < min ${MinProfileSizeMB}MB" -Level 'DEBUG'
            return $false
        }
        if ($MaxProfileSizeMB -and $ProfileSize -gt ($MaxProfileSizeMB * 1MB)) {
            Write-DPLog -Message "    [Filter] $UserName excluded - size $([math]::Round($ProfileSize/1MB,1))MB > max ${MaxProfileSizeMB}MB" -Level 'DEBUG'
            return $false
        }
        
        # Profile type filter (compare script-level filter against actual profile type)
        if ($ProfileType -ne 'All') {
            if ($ActualProfileType -notlike "$ProfileType*") {
                Write-DPLog -Message "    [Filter] $UserName excluded - type '$ActualProfileType' does not match filter '$ProfileType'" -Level 'DEBUG'
                return $false
            }
        }
        
        Write-DPLog -Message "    [Filter] $UserName passed all filters" -Level 'DEBUG'
        return $true
    }
    #endregion

    #region Profile Deletion Functions
    function Dismount-RegistryHive {
        param([string]$ProfilePath)
        
        Write-DPLog -Message "    [Hive] Checking for loaded registry hives matching $ProfilePath" -Level 'DEBUG'
        try {
            # Find loaded hives
            $loadedHives = Get-ChildItem 'HKU:' -ErrorAction SilentlyContinue | 
                Where-Object { $_.Name -match '^HKEY_USERS\\S-1-5-21' }
            Write-DPLog -Message "    [Hive] Found $(@($loadedHives).Count) user hives currently loaded" -Level 'DEBUG'
            
            foreach ($hive in $loadedHives) {
                $sid = Split-Path $hive.Name -Leaf
                $hiveProfilePath = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid" -ErrorAction SilentlyContinue).ProfileImagePath
                
                if ($hiveProfilePath -eq $ProfilePath) {
                    Write-DPLog -Message "    [Hive] Match found - unloading hive for SID $sid" -Level 'INFO'
                    $result = reg.exe unload "HKU\$sid" 2>&1
                    if ($LASTEXITCODE -ne 0) {
                        Write-DPLog -Message "    [Hive] Failed to unload registry hive: $result" -Level 'WARNING'
                        return $false
                    }
                    Write-DPLog -Message "    [Hive] Successfully unloaded hive for SID $sid" -Level 'DEBUG'
                }
            }
            return $true
        }
        catch {
            Write-DPLog -Message "    [Hive] Error during hive unload: $($_.Exception.Message)" -Level 'WARNING'
            return $false
        }
    }

    function Remove-ProfileWithRetry {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Internal helper called from ShouldProcess-guarded callers')]
        param(
            [string]$ProfilePath,
            [string]$SID,
            [string]$ComputerName
        )
        
        $attempt = 0
        $success = $false
        Write-DPLog -Message "    [Delete] Removing profile folder $ProfilePath (max $MaxRetries attempts)" -Level 'DEBUG'
        
        while ($attempt -lt $MaxRetries -and -not $success) {
            $attempt++
            Write-DPLog -Message "    [Delete] Attempt $attempt of $MaxRetries..." -Level 'DEBUG'
            
            try {
                # Unload registry hive if requested
                if ($UnloadHives) {
                    Write-DPLog -Message "    [Delete] Unloading registry hive for $ProfilePath" -Level 'DEBUG'
                    Dismount-RegistryHive -ProfilePath $ProfilePath
                }
                
                # Remove read-only attributes
                Write-DPLog -Message "    [Delete] Clearing read-only attributes..." -Level 'DEBUG'
                Get-ChildItem -Path $ProfilePath -Recurse -Force -ErrorAction SilentlyContinue | 
                    ForEach-Object { $_.Attributes = $_.Attributes -band -bnot [System.IO.FileAttributes]::ReadOnly }
                
                # Remove the directory
                if ($ComputerName -eq $env:COMPUTERNAME -or $ComputerName -eq 'localhost' -or $ComputerName -eq '.') {
                    Write-DPLog -Message "    [Delete] Removing local directory..." -Level 'DEBUG'
                    Remove-Item -Path $ProfilePath -Recurse -Force -ErrorAction Stop
                }
                else {
                    Write-DPLog -Message "    [Delete] Removing remote directory via Invoke-Command..." -Level 'DEBUG'
                    # Remote deletion using Invoke-Command
                    $scriptBlock = {
                        param($Path)
                        Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue | 
                            ForEach-Object { $_.Attributes = $_.Attributes -band -bnot [System.IO.FileAttributes]::ReadOnly }
                        Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
                    }
                    Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $ProfilePath -ErrorAction Stop @script:CredentialSplat
                }
                
                Write-DPLog -Message "    [Delete] Folder removed successfully on attempt $attempt" -Level 'DEBUG'
                $success = $true
            }
            catch {
                if ($attempt -lt $MaxRetries) {
                    Write-DPLog -Message "Attempt $attempt failed for $ProfilePath, retrying in $RetryDelaySeconds seconds..." -Level 'WARNING'
                    Start-Sleep -Seconds $RetryDelaySeconds
                }
                else {
                    Write-DPLog -Message "Failed to delete $ProfilePath after $MaxRetries attempts`: $($_.Exception.Message)" -Level 'ERROR'
                }
            }
        }
        
        return $success
    }

    function Remove-UserProfile {
        [CmdletBinding(SupportsShouldProcess = $true)]
        param(
            [string]$ComputerName,
            [string]$SID,
            [string]$ProfilePath,
            [string]$UserName
        )
        
        Write-DPLog -Message "Deleting profile for '$UserName' on $ComputerName..." -Level 'INFO'
        Write-DPLog -Message "  [Remove] SID=$SID, Path=$ProfilePath" -Level 'DEBUG'
        
        $success = $true
        
        # Step 1: Remove from registry
        Write-DPLog -Message "  [Remove] Step 1: Removing ProfileList registry key for $SID" -Level 'DEBUG'
        if ($PSCmdlet.ShouldProcess("Remove registry key for $SID", "Delete Profile Registry Entry")) {
            try {
                $regKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$SID"
                
                if ($ComputerName -eq $env:COMPUTERNAME -or $ComputerName -eq 'localhost' -or $ComputerName -eq '.') {
                    if (Test-Path $regKey) {
                        Remove-Item -Path $regKey -Recurse -Force
                        Write-DPLog -Message "  [Remove] Local registry key deleted" -Level 'DEBUG'
                    }
                    else {
                        Write-DPLog -Message "  [Remove] Registry key not found (already removed)" -Level 'DEBUG'
                    }
                }
                else {
                    Write-DPLog -Message "  [Remove] Removing remote registry key via Invoke-Command" -Level 'DEBUG'
                    $scriptBlock = {
                        param($sid)
                        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
                        if (Test-Path "$regPath\$sid") {
                            Remove-Item -Path "$regPath\$sid" -Recurse -Force
                        }
                    }
                    Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $SID @script:CredentialSplat
                }
                
                Write-DPLog -Message "Registry key removed for $UserName" -Level 'SUCCESS'
            }
            catch {
                Write-DPLog -Message "Failed to remove registry key for $UserName`: $($_.Exception.Message)" -Level 'ERROR'
                $success = $false
            }
        }
        
        # Step 2: Remove profile directory
        Write-DPLog -Message "  [Remove] Step 2: Removing profile directory $ProfilePath" -Level 'DEBUG'
        if ($success -and $PSCmdlet.ShouldProcess("Remove directory $ProfilePath", "Delete Profile Folder")) {
            if (Remove-ProfileWithRetry -ProfilePath $ProfilePath -SID $SID -ComputerName $ComputerName) {
                Write-DPLog -Message "Profile directory removed for $UserName" -Level 'SUCCESS'
            }
            else {
                Write-DPLog -Message "Failed to remove profile directory for $UserName" -Level 'ERROR'
                $success = $false
            }
        }
        
        Write-DPLog -Message "  [Remove] Result for $UserName`: $( if ($success) { 'SUCCESS' } else { 'FAILED' } )" -Level 'DEBUG'
        return $success
    }
    #endregion

    #region Main Processing Functions
    function Invoke-ComputerProcessing {
        [CmdletBinding(SupportsShouldProcess = $true)]
        param([string]$ComputerName)
        
        Write-DPLog -Message "Processing computer: $ComputerName" -Level 'INFO'
        
        # Test connection
        Write-DPLog -Message "  [Process] Testing connectivity to $ComputerName..." -Level 'DEBUG'
        $connection = Test-ComputerConnection -Computer $ComputerName
        if (-not $connection.Success) {
            Write-DPLog -Message "Cannot connect to $ComputerName`: $($connection.Error)" -Level 'ERROR'
            return
        }
        Write-DPLog -Message "  [Process] Connection to $ComputerName verified" -Level 'DEBUG'
        
        # Get active sessions
        $activeSessions = @()
        if (-not $IgnoreActiveSessions) {
            Write-DPLog -Message "  [Process] Checking active sessions (IgnoreActiveSessions=$IgnoreActiveSessions)..." -Level 'DEBUG'
            $activeSessions = Get-ActiveSession -ComputerName $ComputerName
            Write-DPLog -Message "Active sessions on $ComputerName`: $($activeSessions -join ', ')" -Level 'INFO'
        }
        else {
            Write-DPLog -Message "  [Process] Skipping active session check (IgnoreActiveSessions=True)" -Level 'DEBUG'
        }
        
        # Get profiles
        $profiles = Get-UserProfile -ComputerName $ComputerName
        if ($null -eq $profiles) {
            Write-DPLog -Message "  [Process] No profiles returned for $ComputerName - aborting" -Level 'DEBUG'
            return
        }
        
        Write-DPLog -Message "Found $($profiles.Count) profiles on $ComputerName" -Level 'INFO'
        
        # Process each profile
        foreach ($profileInfo in $profiles) {
            $sid = $profileInfo.SID
            $profilePath = $profileInfo.ProfilePath
            
            Write-DPLog -Message "  [Profile] ---- Processing SID: $sid ----" -Level 'DEBUG'
            
            if (-not $profilePath) {
                Write-DPLog -Message "  [Profile] SID $sid has no path, skipping" -Level 'WARNING'
                continue
            }
            Write-DPLog -Message "  [Profile] Path: $profilePath" -Level 'DEBUG'
            
            # Resolve username
            Write-DPLog -Message "  [Profile] Resolving SID to username..." -Level 'DEBUG'
            $userName = ConvertTo-UserName -SID $sid
            if (-not $userName) {
                $userName = "Unknown ($sid)"
                Write-DPLog -Message "  [Profile] SID could not be resolved, using: $userName" -Level 'DEBUG'
            }
            else {
                Write-DPLog -Message "  [Profile] Resolved: $userName" -Level 'DEBUG'
            }
            
            # Check if protected
            if (-not $IncludeSystemProfiles) {
                $protectCheck = Test-IsProtectedProfile -UserName $userName -SID $sid
                if ($protectCheck) {
                    Write-DPLog -Message "  [Profile] Skipping protected profile: $userName" -Level 'DEBUG'
                    continue
                }
            }
            
            # Get profile type
            $profType = Get-ProfileType -ProfileInfo $profileInfo
            Write-DPLog -Message "  [Profile] Type: $profType" -Level 'DEBUG'
            
            # Skip corrupted unless requested or FixCorruption is enabled
            if ($profType -like 'Corrupted*' -and -not $IncludeCorrupted -and -not $FixCorruption) {
                Write-DPLog -Message "  [Profile] Skipping corrupted profile: $userName ($profType)" -Level 'DEBUG'
                continue
            }
            
            # Handle corruption repair if requested
            if ($profType -like 'Corrupted*' -and $FixCorruption) {
                $repairResult = Repair-CorruptedProfile -ComputerName $ComputerName -SID $sid -UserName $userName -ProfilePath $profilePath -CorruptionType $profType
                
                # Add repair result to the profile info for reporting
                if (-not $script:RepairResults) {
                    $script:RepairResults = [System.Collections.Generic.List[object]]::new()
                }
                $script:RepairResults.Add([PSCustomObject]@{
                    ComputerName = $ComputerName
                    UserName = $userName
                    SID = $sid
                    CorruptionType = $profType
                    Fixed = $repairResult.Fixed
                    ActionTaken = $repairResult.ActionTaken
                })
                
                # If corruption was fixed by recreating NTUSER.DAT, continue processing normally
                if ($repairResult.Fixed -and $repairResult.Choice -eq 'R') {
                    # Re-check profile type - it should now be valid
                    $profType = Get-ProfileType -ProfileInfo $profileInfo
                    Write-DPLog -Message "Profile $userName repaired successfully - continuing with normal processing" -Level 'SUCCESS'
                }
                # If profile was deleted as part of repair, skip further processing
                elseif ($repairResult.Choice -in @('D', 'F')) {
                    Write-DPLog -Message "Profile $userName handled via corruption repair - skipping deletion phase" -Level 'INFO'
                    continue
                }
                # If skipped or failed, move to next profile
                elseif ($repairResult.Choice -eq 'S' -or -not $repairResult.Fixed) {
                    Write-DPLog -Message "Corruption repair skipped or failed for $userName - continuing to next profile" -Level 'WARNING'
                    continue
                }
            }
            
            # Get profile age
            $ageInfo = Get-ProfileAge -ProfilePath $profilePath -SID $sid -Method $AgeCalculation -ComputerName $ComputerName
            $lastUsed = $ageInfo.LastUsed
            $ageSource = $ageInfo.Source
            
            # Calculate age in days
            $ageInDays = if ($lastUsed -eq [DateTime]::MinValue) { -1 } else { [math]::Floor(((Get-Date) - $lastUsed).TotalDays) }
            Write-DPLog -Message "  [Profile] $userName - Age: $ageInDays days (source: $ageSource, threshold: $DaysInactive days)" -Level 'DEBUG'
            
            # Check age criteria
            if ($ageInDays -lt $DaysInactive -and $ageInDays -ge 0) {
                Write-DPLog -Message "  [Profile] $userName is too recent ($ageInDays < $DaysInactive days), skipping" -Level 'DEBUG'
                continue
            }
            
            # Get profile size (only calculate when needed for display or filtering)
            $needsSize = $ShowSpace -or $MinProfileSizeMB -or $MaxProfileSizeMB -or $Detailed
            Write-DPLog -Message "  [Profile] Calculating size: $needsSize" -Level 'DEBUG'
            $sizeBytes = if ($needsSize) {
                Get-ProfileFolderSize -Path $profilePath
            } else { 0 }
            $sizeFormatted = Format-Byte -Bytes $sizeBytes
            if ($needsSize) {
                Write-DPLog -Message "  [Profile] Size: $sizeFormatted ($sizeBytes bytes)" -Level 'DEBUG'
            }
            
            # Apply filters
            Write-DPLog -Message "  [Profile] Applying include/exclude/size/type filters..." -Level 'DEBUG'
            $passesFilter = Test-ProfileFilter -UserName $userName -SID $sid -ProfilePath $profilePath -ProfileSize $sizeBytes -ActualProfileType $profType
            if (-not $passesFilter) {
                Write-DPLog -Message "  [Profile] $userName filtered out by include/exclude/size/type criteria" -Level 'DEBUG'
                continue
            }
            Write-DPLog -Message "  [Profile] $userName passed all filters" -Level 'DEBUG'
            
            # Check for active session (compare short usernames for reliable matching)
            $hasActiveSession = $false
            $shortName = $userName.Split('\')[-1].ToLower()
            foreach ($session in $activeSessions) {
                $sessionShort = $session.Split('\')[-1].ToLower()
                if ($shortName -eq $sessionShort) {
                    $hasActiveSession = $true
                    Write-DPLog -Message "  [Profile] $userName matched active session: $session" -Level 'DEBUG'
                    break
                }
            }
            
            if ($hasActiveSession -and -not $IgnoreActiveSessions) {
                Write-DPLog -Message "  [Profile] Skipping $userName - active session detected" -Level 'WARNING'
                continue
            }
            if ($hasActiveSession) {
                Write-DPLog -Message "  [Profile] $userName has active session but IgnoreActiveSessions is enabled" -Level 'DEBUG'
            }
            
            # Build result object
            $result = [PSCustomObject]@{
                ComputerName = $ComputerName
                UserName = $userName.Split('\')[-1]
                Domain = if ($userName -contains '\') { $userName.Split('\')[0] } else { $env:USERDOMAIN }
                SID = $sid
                ProfilePath = $profilePath
                ProfileType = $profType
                LastUsed = if ($lastUsed -eq [DateTime]::MinValue) { 'Unknown' } else { $lastUsed.ToString('yyyy-MM-dd HH:mm:ss') }
                AgeInDays = $ageInDays
                AgeSource = $ageSource
                SizeBytes = $sizeBytes
                SizeFormatted = $sizeFormatted
                IsActiveSession = $hasActiveSession
                EligibleForDeletion = $true
                Deleted = $false
                Error = $null
            }
            
            # Display info
            if (-not $Quiet) {
                $color = if ($hasActiveSession) { 'Yellow' } else { (Get-AgeColor -AgeInDays $ageInDays) }
                $sizeStr = if ($ShowSpace) { " [Size: $sizeFormatted]" } else { '' }
                $activeStr = if ($hasActiveSession) { ' [ACTIVE]' } else { '' }
                Write-Host "  $userName - $ageInDays days ($ageSource)$sizeStr$activeStr" -ForegroundColor $color
                
                # Show detailed folder breakdown if requested
                if ($Detailed) {
                    $breakdown = Get-ProfileSizeBreakdown -ProfilePath $profilePath
                    $breakdownStr = $breakdown.GetEnumerator() | Where-Object { $_.Value -ne 'N/A' } | 
                        ForEach-Object { "$($_.Key): $($_.Value)" } | Join-String -Separator ', '
                    if ($breakdownStr) {
                        Write-Host "    Folders: $breakdownStr" -ForegroundColor Gray
                    }
                }
            }
            
            # Perform deletion if requested (skipped during mass deletion enumeration pass)
            if ($Delete -and -not $hasActiveSession -and -not $script:SuppressDeletion) {
                Write-DPLog -Message "  [Profile] DELETING $userName ($ageInDays days, $sizeFormatted)..." -Level 'DEBUG'
                $targetDesc = "Delete profile for '$userName' on '$ComputerName' ($ageInDays days old, $sizeFormatted)"
                if ($PSCmdlet.ShouldProcess($targetDesc, 'Delete User Profile')) {
                    # Backup profile before deletion if requested
                    $backupSuccess = $true
                    if ($BackupPath) {
                        $backupSuccess = Backup-Profile -SourcePath $profilePath -UserName $userName
                    }
                    
                    if ($backupSuccess) {
                        if (Remove-UserProfile -ComputerName $ComputerName -SID $sid -ProfilePath $profilePath -UserName $userName) {
                            $result.Deleted = $true
                            $script:TotalProfilesDeleted++
                            $script:TotalSpaceFreed += $sizeBytes
                            Write-EventLogEntry -Message "Deleted profile: $userName on $ComputerName ($sizeFormatted)" -EntryType Information -EventId 1010
                        }
                        else {
                            $result.Error = 'Deletion failed'
                            Write-EventLogEntry -Message "Failed to delete profile: $userName on $ComputerName" -EntryType Error -EventId 1011
                        }
                    }
                    else {
                        $result.Error = 'Backup failed - deletion cancelled'
                        Write-DPLog -Message "Deletion cancelled for $userName - backup failed" -Level 'ERROR'
                    }
                }
            }
            
            $script:Results.Add($result)
            $script:TotalProfilesProcessed++
            Write-DPLog -Message "  [Profile] $userName added to results (Deleted=$($result.Deleted), Error=$($result.Error))" -Level 'DEBUG'
        }
        
        Write-DPLog -Message "Finished processing $ComputerName - $script:TotalProfilesProcessed profile(s) processed so far" -Level 'INFO'
    }

    function Show-Summary {
        if (-not $Quiet) {
            $duration = (Get-Date) - $script:StartTime
            
            Write-Host -Object ("`n" + ('=' * 80)) -ForegroundColor Cyan
            Write-Host -Object " SUMMARY" -ForegroundColor Cyan
            Write-Host -Object ('=' * 80) -ForegroundColor Cyan
            Write-Host " Computers processed: $($ComputerName.Count)"
            Write-Host " Profiles processed:  $script:TotalProfilesProcessed"
            if ($Delete) {
                Write-Host " Profiles deleted:    $script:TotalProfilesDeleted" -ForegroundColor $(if ($script:TotalProfilesDeleted -gt 0) { 'Green' } else { 'White' })
                Write-Host " Space freed:         $(Format-Byte -Bytes $script:TotalSpaceFreed)" -ForegroundColor $(if ($script:TotalSpaceFreed -gt 0) { 'Green' } else { 'White' })
            }
            else {
                # Dry run preview - show what WOULD be deleted
                $wouldDelete = $script:Results | Where-Object { $_.EligibleForDeletion -and -not $_.IsActiveSession }
                $wouldDeleteCount = $wouldDelete.Count
                $wouldDeleteSize = ($wouldDelete | Measure-Object -Property SizeBytes -Sum).Sum
                Write-Host " Would delete:        $wouldDeleteCount profiles ($(Format-Byte -Bytes $wouldDeleteSize))" -ForegroundColor Yellow
            }
            Write-Host " Duration:            $($duration.ToString('hh\:mm\:ss'))"
            
            # Top 5 largest profiles
            if ($script:Results.Count -gt 0) {
                Write-Host "`n TOP 5 LARGEST PROFILES:" -ForegroundColor Cyan
                $top5 = $script:Results | Sort-Object SizeBytes -Descending | Select-Object -First 5
                foreach ($prof in $top5) {
                    Write-Host "  $($prof.UserName) on $($prof.ComputerName): $($prof.SizeFormatted) ($($prof.AgeInDays) days)" -ForegroundColor Gray
                }
            }
            
            # Age breakdown analysis
            if ($script:Results.Count -gt 0) {
                Write-Host "`n AGE BREAKDOWN:" -ForegroundColor Cyan
                $ageGroups = $script:Results | Group-Object -Property { 
                    if ($_.AgeInDays -lt 30) { '0-30 days' }
                    elseif ($_.AgeInDays -lt 60) { '31-60 days' }
                    elseif ($_.AgeInDays -lt 90) { '61-90 days' }
                    elseif ($_.AgeInDays -lt 180) { '91-180 days' }
                    elseif ($_.AgeInDays -ge 180) { '180+ days' }
                    else { 'Unknown' }
                } | Sort-Object Name
                
                foreach ($group in $ageGroups) {
                    $groupSize = ($group.Group | Measure-Object -Property SizeBytes -Sum).Sum
                    Write-Host "  $($group.Name): $($group.Count) profiles ($(Format-Byte -Bytes $groupSize))"
                }
            }
            
            Write-Host ('=' * 80) -ForegroundColor Cyan
        }
        
        Write-DPLog -Message "Script completed. Processed: $script:TotalProfilesProcessed, Deleted: $script:TotalProfilesDeleted" -Level 'INFO'
    }
    #endregion

    #region Script Entry Point
    Write-DPHeader
    Write-DPLog -Message "Parameters: DaysInactive=$DaysInactive, AgeCalculation=$AgeCalculation, Delete=$Delete, Preview=$Preview, Force=$Force" -Level 'DEBUG'
    Write-DPLog -Message "Parameters: ShowSpace=$ShowSpace, Detailed=$Detailed, UnloadHives=$UnloadHives, UseParallel=$UseParallel" -Level 'DEBUG'
    Write-DPLog -Message "Parameters: Include=[$($Include -join ', ')], Exclude=[$($Exclude -join ', ')], ProfileType=$ProfileType" -Level 'DEBUG'
    Write-DPLog -Message "Parameters: ComputerName=[$($ComputerName -join ', ')], Interactive=$Interactive, IgnoreActiveSessions=$IgnoreActiveSessions" -Level 'DEBUG'
    if ($ConfigFile) { Write-DPLog -Message "Parameters: ConfigFile=$ConfigFile" -Level 'DEBUG' }
    if ($LogPath) { Write-DPLog -Message "Parameters: LogPath=$LogPath" -Level 'DEBUG' }
    if ($OutputPath) { Write-DPLog -Message "Parameters: OutputPath=$OutputPath" -Level 'DEBUG' }
    if ($HtmlReport) { Write-DPLog -Message "Parameters: HtmlReport=$HtmlReport" -Level 'DEBUG' }
    if ($BackupPath) { Write-DPLog -Message "Parameters: BackupPath=$BackupPath" -Level 'DEBUG' }
    
    # Log to event log
    Write-EventLogEntry -Message "Delprof2-PS started. Version: $script:Version, Delete mode: $Delete, Days inactive: $DaysInactive" -EntryType Information -EventId 1000
    
    # Test mode - validate prerequisites without processing
    if ($Test) {
        Write-DPLog -Message '[Test] Entering TEST MODE - validating prerequisites only' -Level 'INFO'
        Write-Host "`n=== TEST MODE - Validating Prerequisites ===" -ForegroundColor Cyan
        foreach ($computer in $ComputerName) {
            Write-DPLog -Message "[Test] Testing connectivity to $computer..." -Level 'INFO'
            Write-Host "Testing $computer..." -NoNewline
            $result = Test-ComputerConnection -Computer $computer.Trim()
            if ($result.Success) {
                Write-Host " OK" -ForegroundColor Green
                Write-DPLog -Message "[Test] ${computer}: Connection OK" -Level 'INFO'
                try {
                    $profiles = Get-UserProfile -ComputerName $computer.Trim()
                    Write-Host "  Found $($profiles.Count) profiles" -ForegroundColor Gray
                    Write-DPLog -Message "[Test] ${computer}: Found $($profiles.Count) profiles" -Level 'INFO'
                }
                catch {
                    Write-Host "  ERROR: Could not enumerate profiles" -ForegroundColor Red
                    Write-DPLog -Message "[Test] ${computer}: Failed to enumerate profiles - $($_.Exception.Message)" -Level 'ERROR'
                }
            }
            else {
                Write-Host " FAILED - $($result.Error)" -ForegroundColor Red
                Write-DPLog -Message "[Test] ${computer}: FAILED - $($result.Error)" -Level 'ERROR'
            }
        }
        Write-DPLog -Message '[Test] Test mode complete. No changes were made.' -Level 'INFO'
        Write-Host "`nTest mode complete. No changes were made." -ForegroundColor Cyan
        exit 0
    }
    
    # Preview mode banner
    if ($Preview) {
        Write-Host "`n==================================================================================================================================================—" -ForegroundColor Magenta
        Write-Host "=                        PREVIEW/SIMULATION MODE                         =" -ForegroundColor Magenta
        Write-Host "=          No profiles will be deleted - showing what WOULD happen       =" -ForegroundColor Magenta
        Write-Host "====================================================================================================================================================" -ForegroundColor Magenta
    }
    
    # Validate admin rights for local execution
    Write-DPLog -Message '[Init] Checking administrator privileges...' -Level 'DEBUG'
    if ($ComputerName -contains $env:COMPUTERNAME -or $ComputerName -contains 'localhost' -or $ComputerName -contains '.') {
        if (-not (Test-AdminRight)) {
            Write-DPLog -Message "Administrator privileges required for local execution. Restart as admin." -Level 'ERROR'
            Write-EventLogEntry -Message "Failed to start - admin rights required" -EntryType Error -EventId 1005
            exit 1
        }
    }
    
    # Validate delete warnings
    if ($Delete -and -not $Force -and -not $Quiet -and -not $Interactive) {
        Write-Host "`nWARNING: Delete mode is enabled. Profiles will be permanently removed!" -ForegroundColor Red -BackgroundColor Black
        if ($IgnoreActiveSessions) {
            Write-Host "WARNING: Active sessions will be deleted - this can cause data loss!" -ForegroundColor Red -BackgroundColor Black
        }
        Write-Host "Press Ctrl+C to cancel, or " -NoNewline -ForegroundColor Yellow
        Write-Host "Enter to continue..." -NoNewline -ForegroundColor Yellow
        $null = Read-Host
    }
    
    # Build computer queue
    Write-DPLog -Message "[Init] Building computer queue with $($ComputerName.Count) computer(s)" -Level 'DEBUG'
    foreach ($computer in $ComputerName) {
        $script:ComputerQueue.Add($computer.Trim())
        Write-DPLog -Message "  [Queue] Added: $($computer.Trim())" -Level 'DEBUG'
    }
    
    # Pre-validate all computers if using interactive mode
    $validComputers = [System.Collections.Generic.List[string]]::new()
    if ($Interactive) {
        Write-Host "`nValidating computers for interactive mode..." -ForegroundColor Cyan
        foreach ($computer in $ComputerName) {
            $result = Test-ComputerConnection -Computer $computer.Trim()
            if ($result.Success) {
                $validComputers.Add($computer.Trim())
            }
            else {
                Write-DPLog -Message "Skipping $computer in interactive mode - $($result.Error)" -Level 'WARNING'
            }
        }
    }
}

process {
    if ($script:UIMode) { return }
    Write-DPLog -Message '--- Process block started ---' -Level 'INFO'
    # Mass deletion safeguard - enumerate first, then delete from collected results
    if ($Delete -and -not $Force -and -not $Quiet -and -not $Interactive) {
        Write-DPLog -Message '[Process] Mode: Mass deletion with safeguard (enumerate then delete)' -Level 'INFO'
        # Temporarily suppress deletion so the enumeration pass only collects data
        $script:SuppressDeletion = $true
        
        # Single enumeration pass - profiles are collected into $script:Results
        $computerCount = $ComputerName.Count
        Write-DPLog -Message "[Process] Enumeration pass: scanning $computerCount computer(s)..." -Level 'INFO'
        for ($i = 0; $i -lt $computerCount; $i++) {
            $percentComplete = [math]::Floor(($i / $computerCount) * 100)
            Write-Progress -Activity "Enumerating Profiles" -Status "Scanning $($ComputerName[$i]) - $($i + 1) of $computerCount" -PercentComplete $percentComplete
            Invoke-ComputerProcessing -ComputerName $ComputerName[$i].Trim()
        }
        Write-Progress -Activity "Enumerating Profiles" -Completed
        
        $script:SuppressDeletion = $false
        Write-DPLog -Message "[Process] Enumeration complete. Total results: $($script:Results.Count)" -Level 'INFO'
        
        # Count eligible profiles from the collected results
        $eligibleProfiles = $script:Results | Where-Object { $_.EligibleForDeletion -and -not $_.IsActiveSession }
        $estimatedDeletions = $eligibleProfiles.Count
        Write-DPLog -Message "[Process] Eligible for deletion: $estimatedDeletions profile(s)" -Level 'INFO'
        
        # Mass deletion warning threshold
        if ($estimatedDeletions -gt 50) {
            Write-Host "`n  MASS DELETION WARNING " -ForegroundColor Red -BackgroundColor Black
            Write-Host "This operation will delete $estimatedDeletions profiles!" -ForegroundColor Red
            Write-Host "This is an unusually large number of deletions." -ForegroundColor Yellow
            Write-Host "Are you sure you want to proceed?" -ForegroundColor Yellow
            Write-Host "Type 'YES' to confirm: " -NoNewline -ForegroundColor Yellow
            $confirm = Read-Host
            if ($confirm -ne 'YES') {
                Write-Host "Operation cancelled by user." -ForegroundColor Red
                exit 1
            }
        }
        
        # Perform deletions on already-collected results (no second scan needed)
        Write-DPLog -Message "[Process] Beginning deletion phase for $($eligibleProfiles.Count) profile(s)..." -Level 'INFO'
        $deleteIndex = 0
        foreach ($prof in $eligibleProfiles) {
            $deleteIndex++
            Write-DPLog -Message "[Process] Deleting [$deleteIndex/$($eligibleProfiles.Count)]: $($prof.UserName) on $($prof.ComputerName)" -Level 'INFO'
            $targetDesc = "Delete profile for '$($prof.UserName)' on '$($prof.ComputerName)' ($($prof.AgeInDays) days old, $($prof.SizeFormatted))"
            if ($PSCmdlet.ShouldProcess($targetDesc, 'Delete User Profile')) {
                # Backup profile before deletion if requested
                $backupSuccess = $true
                if ($BackupPath) {
                    $backupSuccess = Backup-Profile -SourcePath $prof.ProfilePath -UserName $prof.UserName
                }
                
                if ($backupSuccess) {
                    if (Remove-UserProfile -ComputerName $prof.ComputerName -SID $prof.SID -ProfilePath $prof.ProfilePath -UserName $prof.UserName) {
                        $prof.Deleted = $true
                        $script:TotalProfilesDeleted++
                        $script:TotalSpaceFreed += $prof.SizeBytes
                        Write-EventLogEntry -Message "Deleted profile: $($prof.UserName) on $($prof.ComputerName) ($($prof.SizeFormatted))" -EntryType Information -EventId 1010
                    }
                    else {
                        $prof.Error = 'Deletion failed'
                        Write-EventLogEntry -Message "Failed to delete profile: $($prof.UserName) on $($prof.ComputerName)" -EntryType Error -EventId 1011
                    }
                }
                else {
                    $prof.Error = 'Backup failed - deletion cancelled'
                    Write-DPLog -Message "Deletion cancelled for $($prof.UserName) - backup failed" -Level 'ERROR'
                }
            }
        }
    }
    
    # Interactive mode processing
    elseif ($Interactive) {
        Write-DPLog -Message "[Process] Mode: Interactive - scanning $($validComputers.Count) validated computer(s)" -Level 'INFO'
        foreach ($computer in $validComputers) {
            Write-DPLog -Message "[Process] Interactive scan: $computer" -Level 'INFO'
            Invoke-ComputerProcessing -ComputerName $computer
        }
        
        # After collecting all eligible profiles, let user select interactively
        $eligibleProfiles = $script:Results | Where-Object { $_.EligibleForDeletion -and -not $_.IsActiveSession }
        
        if ($eligibleProfiles.Count -eq 0) {
            Write-DPLog -Message '[Interactive] No eligible profiles found for selection' -Level 'WARNING'
            Write-Host "`nNo eligible profiles found for interactive selection." -ForegroundColor Yellow
        }
        else {
            Write-DPLog -Message "[Interactive] Presenting $($eligibleProfiles.Count) profile(s) for interactive selection" -Level 'INFO'
            $selectedProfiles = Select-ProfilesInteractive -Profiles $eligibleProfiles
            
            if ($selectedProfiles.Count -gt 0) {
                Write-DPLog -Message "[Interactive] User selected $($selectedProfiles.Count) profile(s) for deletion" -Level 'INFO'
                Write-Host "`nDeleting $($selectedProfiles.Count) selected profiles..." -ForegroundColor Yellow
                foreach ($prof in $selectedProfiles) {
                    Write-DPLog -Message "[Interactive] Deleting: $($prof.UserName) on $($prof.ComputerName)" -Level 'INFO'
                    if (Remove-UserProfile -ComputerName $prof.ComputerName -SID $prof.SID -ProfilePath $prof.ProfilePath -UserName $prof.UserName) {
                        # Update the result
                        $resultItem = $script:Results | Where-Object { $_.SID -eq $prof.SID -and $_.ComputerName -eq $prof.ComputerName }
                        if ($resultItem) {
                            $resultItem.Deleted = $true
                            $script:TotalProfilesDeleted++
                            $script:TotalSpaceFreed += $resultItem.SizeBytes
                        }
                        Write-EventLogEntry -Message "Deleted profile: $($prof.UserName) on $($prof.ComputerName)" -EntryType Information -EventId 1010
                    }
                }
            }
            else {
                Write-DPLog -Message '[Interactive] User selected no profiles for deletion' -Level 'INFO'
                Write-Host "`nNo profiles selected for deletion." -ForegroundColor Green
            }
        }
    }
    elseif ($UseParallel -and $ComputerName.Count -gt 1) {
        Write-DPLog -Message "[Process] Mode: Parallel processing ($($ComputerName.Count) computers, throttle=$ThrottleLimit)" -Level 'INFO'
        # Parallel processing using runspace pool (PS 5.1 compatible)
        # ForEach-Object -Parallel requires PS 7+ and doesn't share script scope,
        # so we use a RunspacePool with jobs that each invoke the script per-computer.
        $runspacePool = [runspacefactory]::CreateRunspacePool(1, $ThrottleLimit)
        $runspacePool.Open()
        $jobs = [System.Collections.Generic.List[object]]::new()

        $parallelScript = {
            param($ScriptPath, $Computer, $Params)
            # Re-invoke the script for a single computer in list mode (no -Delete)
            # to collect profile data; deletion happens in the main thread below.
            $splatParams = @{
                ComputerName = $Computer
                DaysInactive = $Params.DaysInactive
                AgeCalculation = $Params.AgeCalculation
                ProfileType = $Params.ProfileType
                Quiet = $true
            }
            if ($Params.Include) { $splatParams['Include'] = $Params.Include }
            if ($Params.Exclude) { $splatParams['Exclude'] = $Params.Exclude }
            if ($Params.ShowSpace) { $splatParams['ShowSpace'] = $true }
            if ($Params.LogPath) { $splatParams['LogPath'] = $Params.LogPath }
            if ($Params.UnloadHives) { $splatParams['UnloadHives'] = $true }
            if ($Params.IncludeCorrupted) { $splatParams['IncludeCorrupted'] = $true }
            if ($Params.IncludeSystemProfiles) { $splatParams['IncludeSystemProfiles'] = $true }
            if ($Params.IncludeSpecialProfiles) { $splatParams['IncludeSpecialProfiles'] = $true }
            if ($Params.MinProfileSizeMB) { $splatParams['MinProfileSizeMB'] = $Params.MinProfileSizeMB }
            if ($Params.MaxProfileSizeMB) { $splatParams['MaxProfileSizeMB'] = $Params.MaxProfileSizeMB }
            if ($Params.Detailed) { $splatParams['Detailed'] = $true }
            & $ScriptPath @splatParams
        }

        $parallelParams = @{
            DaysInactive = $DaysInactive
            AgeCalculation = $AgeCalculation
            ProfileType = $ProfileType
            Include = $Include
            Exclude = $Exclude
            ShowSpace = [bool]$ShowSpace
            LogPath = $LogPath
            UnloadHives = [bool]$UnloadHives
            IncludeCorrupted = [bool]$IncludeCorrupted
            IncludeSystemProfiles = [bool]$IncludeSystemProfiles
            IncludeSpecialProfiles = [bool]$IncludeSpecialProfiles
            MinProfileSizeMB = $MinProfileSizeMB
            MaxProfileSizeMB = $MaxProfileSizeMB
            Detailed = [bool]$Detailed
        }

        foreach ($computer in $ComputerName) {
            Write-DPLog -Message "[Parallel] Submitting job for $($computer.Trim())" -Level 'DEBUG'
            $ps = [powershell]::Create().AddScript($parallelScript).AddArgument($PSCommandPath).AddArgument($computer.Trim()).AddArgument($parallelParams)
            $ps.RunspacePool = $runspacePool
            $jobs.Add([PSCustomObject]@{ PowerShell = $ps; Handle = $ps.BeginInvoke() })
        }
        Write-DPLog -Message "[Parallel] All $($jobs.Count) jobs submitted, collecting results..." -Level 'INFO'

        # Collect results from all runspaces
        $jobIndex = 0
        foreach ($job in $jobs) {
            $jobIndex++
            Write-Progress -Activity "Parallel Processing" -Status "Collecting results $jobIndex of $($jobs.Count)" -PercentComplete ([math]::Floor(($jobIndex / $jobs.Count) * 100))
            try {
                $jobResults = $job.PowerShell.EndInvoke($job.Handle)
                if ($jobResults) {
                    Write-DPLog -Message "[Parallel] Job $jobIndex returned $(@($jobResults).Count) result(s)" -Level 'DEBUG'
                    foreach ($r in $jobResults) {
                        $script:Results.Add($r)
                        $script:TotalProfilesProcessed++
                        if ($r.Deleted) {
                            $script:TotalProfilesDeleted++
                            $script:TotalSpaceFreed += $r.SizeBytes
                        }
                    }
                } else {
                    Write-DPLog -Message "[Parallel] Job $jobIndex returned no results" -Level 'DEBUG'
                }
            }
            catch {
                Write-DPLog -Message "Parallel job $jobIndex failed: $($_.Exception.Message)" -Level 'ERROR'
            }
            finally {
                $job.PowerShell.Dispose()
            }
        }
        Write-Progress -Activity "Parallel Processing" -Completed

        $runspacePool.Close()
        $runspacePool.Dispose()
    }
    else {
        # Standard sequential processing with progress bar
        Write-DPLog -Message "[Process] Mode: Sequential processing for $($ComputerName.Count) computer(s)" -Level 'INFO'
        $computerCount = $ComputerName.Count
        for ($i = 0; $i -lt $computerCount; $i++) {
            Write-DPLog -Message "[Process] Sequential [$($i+1)/$computerCount]: $($ComputerName[$i])" -Level 'INFO'
            $percentComplete = [math]::Floor(($i / $computerCount) * 100)
            Write-Progress -Activity "Processing Computers" -Status "Processing $($ComputerName[$i]) - $($i + 1) of $computerCount" -PercentComplete $percentComplete
            Invoke-ComputerProcessing -ComputerName $ComputerName[$i].Trim()
        }
        Write-Progress -Activity "Processing Computers" -Completed
    }
}

end {
    if ($script:UIMode) { return }
    Write-DPLog -Message '--- End block started ---' -Level 'INFO'
    Show-Summary
    
    # Preview mode completion message
    if ($Preview) {
        Write-Host "`n==================================================================================================================================================" -ForegroundColor Magenta
        Write-Host "=                    PREVIEW MODE COMPLETE                               =" -ForegroundColor Magenta
        Write-Host "=        No profiles were deleted. Use -Delete to perform deletion.      =" -ForegroundColor Magenta
        Write-Host "====================================================================================================================================================" -ForegroundColor Magenta
    }
    
    # Generate HTML report if requested
    Write-DPLog -Message '[End] Checking post-processing tasks...' -Level 'DEBUG'
    if ($HtmlReport -and $script:Results.Count -gt 0) {
        Write-DPLog -Message "[End] Generating HTML report: $HtmlReport" -Level 'INFO'
        $summary = @{
            Computers = $ComputerName.Count
            ProfilesProcessed = $script:TotalProfilesProcessed
            ProfilesDeleted = $script:TotalProfilesDeleted
            SpaceFreed = Format-Byte -Bytes $script:TotalSpaceFreed
            Duration = ((Get-Date) - $script:StartTime).ToString('hh\:mm\:ss')
        }
        Export-HtmlReport -Path $HtmlReport -Results $script:Results -Summary $summary
    }
    
    # Send email notification if configured
    if ($SmtpServer -and $EmailTo) {
        Write-DPLog -Message "[End] Sending email notification to $EmailTo via $SmtpServer" -Level 'INFO'
        $summary = @{
            Computers = $ComputerName.Count
            ProfilesProcessed = $script:TotalProfilesProcessed
            ProfilesDeleted = $script:TotalProfilesDeleted
            SpaceFreed = Format-Byte -Bytes $script:TotalSpaceFreed
            Duration = ((Get-Date) - $script:StartTime).ToString('hh\:mm\:ss')
        }
        Send-NotificationEmail -Summary $summary
    }
    
    # Log completion to event log
    Write-EventLogEntry -Message "Delprof2-PS completed. Processed: $script:TotalProfilesProcessed, Deleted: $script:TotalProfilesDeleted, Space freed: $(Format-Byte -Bytes $script:TotalSpaceFreed)" -EntryType Information -EventId 1002
    
    # Export to CSV if requested
    if ($OutputPath -and $script:Results.Count -gt 0) {
        Write-DPLog -Message "[End] Exporting $($script:Results.Count) results to CSV: $OutputPath" -Level 'INFO'
        try {
            $script:Results | Export-Csv -Path $OutputPath -NoTypeInformation -Force
            Write-DPLog -Message "Results exported to $OutputPath" -Level 'SUCCESS'
        }
        catch {
            Write-DPLog -Message "Failed to export to CSV`: $($_.Exception.Message)" -Level 'ERROR'
        }
    }
    
    # Security - Log Integrity Hash
    Write-DPLog -Message '[End] Finalising...' -Level 'DEBUG'
    if ($LogPath -and (Test-Path $LogPath)) {
        Write-DPLog -Message "[End] Appending integrity hash to log file: $LogPath" -Level 'DEBUG'
        try {
            $logHash = (Get-FileHash -Path $LogPath -Algorithm SHA256).Hash
            $hashEntry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [INTEGRITY] Log SHA256: $logHash"
            Add-Content -Path $LogPath -Value $hashEntry
            Write-DPLog -Message "Log integrity hash appended to $LogPath" -Level 'VERBOSE'
        }
        catch {
            # Silent fail - integrity hashing is supplementary
        }
    }

    # Return results for pipeline
    $duration = (Get-Date) - $script:StartTime
    Write-DPLog -Message "[End] Script completed in $($duration.ToString('hh\:mm\:ss')). Returning $($script:Results.Count) result(s) to pipeline." -Level 'INFO'
    if ($script:Results.Count -gt 0) {
        return $script:Results
    }
}
#endregion
