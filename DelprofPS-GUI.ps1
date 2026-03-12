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
        $xaml.SelectNodes("//*[@x:Name]") | ForEach-Object {
            $name = $_.Name
            $controls[$name] = $window.FindName($name)
        }
        
        # Script-level variables for GUI state
        $script:guiRunning = $false
        $script:guiStopRequested = $false
        
        # Output function for GUI
        $script:WriteGuiOutput = {
            param([string]$Text, [string]$Color = "White")
            $timestamp = Get-Date -Format "HH:mm:ss"
            $coloredText = "[$timestamp] $Text"
            $controls['txtOutput'].AppendText("$coloredText`r`n")
            $controls['txtOutput'].ScrollToEnd()
        }
        
        # Event Handler: Throttle Slider
        $controls['sldThrottle'].Add_ValueChanged({
            $controls['txtThrottleValue'].Text = $controls['sldThrottle'].Value
        })
        
        # Event Handler: Days Slider
        $controls['sldDaysInactive'].Add_ValueChanged({
            $controls['txtDaysValue'].Text = "$($controls['sldDaysInactive'].Value) days"
        })
        
        # Event Handler: Browse Buttons
        $controls['btnBrowseBackup'].Add_Click({
            $folder = New-Object Windows.Forms.FolderBrowserDialog
            $folder.Description = "Select Backup Directory"
            if ($folder.ShowDialog() -eq "OK") {
                $controls['txtBackupPath'].Text = $folder.SelectedPath
                $controls['chkBackup'].IsChecked = $true
            }
        })
        
        $controls['btnBrowseLog'].Add_Click({
            $save = New-Object Windows.Forms.SaveFileDialog
            $save.Filter = "Log files (*.log)|*.log|Text files (*.txt)|*.txt|All files (*.*)|*.*"
            $save.FileName = "DelprofPS.log"
            if ($save.ShowDialog() -eq "OK") {
                $controls['txtLogPath'].Text = $save.FileName
                $controls['chkLogPath'].IsChecked = $true
            }
        })
        
        $controls['btnBrowseCSV'].Add_Click({
            $save = New-Object Windows.Forms.SaveFileDialog
            $save.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
            $save.FileName = "DelprofPS_Results.csv"
            if ($save.ShowDialog() -eq "OK") {
                $controls['txtOutputPath'].Text = $save.FileName
                $controls['chkOutputCSV'].IsChecked = $true
            }
        })
        
        $controls['btnBrowseHtml'].Add_Click({
            $save = New-Object Windows.Forms.SaveFileDialog
            $save.Filter = "HTML files (*.html)|*.html|All files (*.*)|*.*"
            $save.FileName = "DelprofPS_Report.html"
            if ($save.ShowDialog() -eq "OK") {
                $controls['txtHtmlPath'].Text = $save.FileName
                $controls['chkHtmlReport'].IsChecked = $true
            }
        })
        
        # Event Handler: Load Config
        $controls['btnLoadConfig'].Add_Click({
            $open = New-Object Windows.Forms.OpenFileDialog
            $open.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
            $open.Title = "Load DelprofPS Configuration"
            if ($open.ShowDialog() -eq "OK") {
                try {
                    $config = Get-Content $open.FileName | ConvertFrom-Json
                    if ($config.DaysInactive) { $controls['sldDaysInactive'].Value = $config.DaysInactive }
                    if ($config.Exclude) { $controls['txtExclude'].Text = ($config.Exclude -join ', ') }
                    if ($config.Include) { $controls['txtInclude'].Text = ($config.Include -join ', ') }
                    if ($config.LogPath) { 
                        $controls['txtLogPath'].Text = $config.LogPath
                        $controls['chkLogPath'].IsChecked = $true
                    }
                    if ($config.OutputPath) { 
                        $controls['txtOutputPath'].Text = $config.OutputPath
                        $controls['chkOutputCSV'].IsChecked = $true
                    }
                    if ($config.HtmlReport) { 
                        $controls['txtHtmlPath'].Text = $config.HtmlReport
                        $controls['chkHtmlReport'].IsChecked = $true
                    }
                    if ($config.BackupPath) { 
                        $controls['txtBackupPath'].Text = $config.BackupPath
                        $controls['chkBackup'].IsChecked = $true
                    }
                    & $script:WriteGuiOutput -Text "Configuration loaded from $($open.FileName)" -Color "Green"
                }
                catch {
                    [System.Windows.MessageBox]::Show("Failed to load configuration: $($_.Exception.Message)", "Error", "OK", "Error")
                }
            }
        })
        
        # Event Handler: Save Config
        $controls['btnSaveConfig'].Add_Click({
            $save = New-Object Windows.Forms.SaveFileDialog
            $save.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
            $save.FileName = "DelprofPS.config.json"
            if ($save.ShowDialog() -eq "OK") {
                try {
                    $config = @{
                        DaysInactive = $controls['sldDaysInactive'].Value
                        Exclude = @($controls['txtExclude'].Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                        Include = @($controls['txtInclude'].Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                    }
                    if ($controls['chkLogPath'].IsChecked -and $controls['txtLogPath'].Text) { $config.LogPath = $controls['txtLogPath'].Text }
                    if ($controls['chkOutputCSV'].IsChecked -and $controls['txtOutputPath'].Text) { $config.OutputPath = $controls['txtOutputPath'].Text }
                    if ($controls['chkHtmlReport'].IsChecked -and $controls['txtHtmlPath'].Text) { $config.HtmlReport = $controls['txtHtmlPath'].Text }
                    if ($controls['chkBackup'].IsChecked -and $controls['txtBackupPath'].Text) { $config.BackupPath = $controls['txtBackupPath'].Text }
                    if ($controls['txtSmtpServer'].Text) { $config.SmtpServer = $controls['txtSmtpServer'].Text }
                    if ($controls['txtEmailTo'].Text) { $config.EmailTo = $controls['txtEmailTo'].Text }
                    
                    $config | ConvertTo-Json -Depth 3 | Out-File $save.FileName
                    & $script:WriteGuiOutput -Text "Configuration saved to $($save.FileName)" -Color "Green"
                }
                catch {
                    [System.Windows.MessageBox]::Show("Failed to save configuration: $($_.Exception.Message)", "Error", "OK", "Error")
                }
            }
        })
        
        # Event Handler: Clear Output
        $controls['btnClear'].Add_Click({
            $controls['txtOutput'].Clear()
        })
        
        # Event Handler: Stop
        $controls['btnStop'].Add_Click({
            $script:guiStopRequested = $true
            & $script:WriteGuiOutput -Text "Stop requested... waiting for current operation to complete..." -Color "Yellow"
            $controls['btnStop'].IsEnabled = $false
        })
        
        # Event Handler: Run
        $controls['btnRun'].Add_Click({
            if ($script:guiRunning) { return }
            
            $script:guiRunning = $true
            $script:guiStopRequested = $false
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
            & $script:WriteGuiOutput -Text "Executing command:" -Color "Cyan"
            & $script:WriteGuiOutput -Text ($cmdParts -join ' ') -Color "White"
            & $script:WriteGuiOutput -Text "---" -Color "Gray"
            
            # Run in background runspace to keep UI responsive
            $runspace = [runspacefactory]::CreateRunspace()
            $runspace.Open()
            $runspace.SessionStateProxy.SetVariable('params', $params)
            $runspace.SessionStateProxy.SetVariable('PSScriptRoot', $PSScriptRoot)
            
            $powershell = [powershell]::Create().AddScript({
                try {
                    $mainScript = Join-Path $PSScriptRoot 'DelprofPS.ps1'
                    & $mainScript @params
                }
                catch {
                    Write-Error "ERROR: $($_.Exception.Message)"
                }
            })
            
            $powershell.Runspace = $runspace
            $asyncResult = $powershell.BeginInvoke()
            
            # Monitor completion
            while (-not $asyncResult.IsCompleted) {
                Start-Sleep -Milliseconds 100
                [System.Windows.Forms.Application]::DoEvents()
                if ($script:guiStopRequested) {
                    $powershell.Stop()
                    break
                }
            }
            
            try {
                $powershell.EndInvoke($asyncResult)
            }
            catch {
                & $script:WriteGuiOutput -Text "Operation stopped or encountered an error." -Color "Yellow"
            }
            finally {
                $powershell.Dispose()
                $runspace.Close()
                $runspace.Dispose()
            }
            
            $script:guiRunning = $false
            $controls['btnRun'].IsEnabled = $true
            $controls['btnStop'].IsEnabled = $false
            $controls['progressBar'].Visibility = "Collapsed"
            
            & $script:WriteGuiOutput -Text "---" -Color "Gray"
            & $script:WriteGuiOutput -Text "Operation completed." -Color "Green"
            
            if (-not $params['Quiet']) {
                [System.Windows.MessageBox]::Show("Profile management operation completed!", "Complete", "OK", "Information")
            }
        })
        
        # Show Window
        $window.ShowDialog() | Out-Null
    }
    #endregion
