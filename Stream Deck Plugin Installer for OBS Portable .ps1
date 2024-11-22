Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
function StreamDeckOBSPortableInstaller {
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Stream Deck OBS Plugin Installer"
    Height="520" Width="700"
    WindowStartupLocation="CenterScreen"
    Background="Transparent"
    WindowStyle="None"
    AllowsTransparency="True"
    ResizeMode="NoResize">
 
  <Window.Resources>
    <Style x:Key="ModernButtonStyle" TargetType="Button">
      <Setter Property="Background" Value="#2563EB"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Height" Value="40"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="border"
                Background="{TemplateBinding Background}"
                CornerRadius="6"
                Padding="20,8">
              <ContentPresenter x:Name="content"
                      HorizontalAlignment="Center"
                      VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="border" Property="Background" Value="#1D4ED8"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="border" Property="Background" Value="#27272A"/>
                <Setter Property="Foreground" Value="#52525B"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="NavigationButtonStyle" TargetType="Button" BasedOn="{StaticResource ModernButtonStyle}">
      <Setter Property="Width" Value="100"/>
      <Setter Property="Margin" Value="8,0"/>
    </Style>
    <Style x:Key="StepNumberStyle" TargetType="TextBlock">
      <Setter Property="FontSize" Value="48"/>
      <Setter Property="FontWeight" Value="Light"/>
      <Setter Property="Foreground" Value="#3F3F46"/>
      <Setter Property="VerticalAlignment" Value="Top"/>
      <Setter Property="Margin" Value="0,0,24,0"/>
    </Style>
  </Window.Resources>
  <Border Background="#171717"
      BorderBrush="#27272A"
      BorderThickness="1"
      CornerRadius="12">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>
      <!-- Title Bar -->
      <Grid Grid.Row="0"
         Margin="24,24,24,0"
         Background="Transparent"
         x:Name="TitleBar">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel>
          <TextBlock Text="Stream Deck OBS Plugin Installer"
               FontSize="24"
               FontWeight="Bold"
               Foreground="White"/>
          <TextBlock Text="Install the Stream Deck plugin for OBS Studio in portable mode"
               Foreground="#71717A"
               Margin="0,4,0,0"/>
        </StackPanel>
        <Button x:Name="CloseButton"
            Grid.Column="1"
            Content="✕"
            Width="32"
            Height="32"
            Background="Transparent"
            Foreground="#71717A"
            BorderThickness="0"
            Cursor="Hand"/>
      </Grid>
      <!-- Content Area -->
      <Grid Grid.Row="1" Margin="24">
        <Grid x:Name="WelcomePage" Visibility="Visible">
          <StackPanel>
            <TextBlock Text="Before we begin"
                 FontSize="20"
                 FontWeight="SemiBold"
                 Foreground="White"
                 Margin="0,0,0,16"/>
           
            <TextBlock Text="Please ensure you have:"
                 Foreground="#E4E4E7"
                 Margin="0,0,0,12"/>
           
            <ItemsControl Margin="8,0,0,0">
              <ItemsControl.ItemTemplate>
                <DataTemplate>
                  <StackPanel Orientation="Horizontal" Margin="0,8">
                    <TextBlock Text="•" Foreground="#2563EB" Margin="0,0,8,0"/>
                    <TextBlock Text="{Binding}" Foreground="#E4E4E7"/>
                  </StackPanel>
                </DataTemplate>
              </ItemsControl.ItemTemplate>
              <ItemsControl.Items>
                <x:String>Stream Deck software installed on your computer</x:String>
                <x:String>OBS Studio Portable version installed</x:String>
                <x:String>Administrator privileges (recommended)</x:String>
              </ItemsControl.Items>
            </ItemsControl>
          </StackPanel>
        </Grid>
        <Grid x:Name="LocationPage" Visibility="Collapsed">
          <StackPanel>
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
             
              <TextBlock Text="1" Style="{StaticResource StepNumberStyle}"/>
             
              <StackPanel Grid.Column="1">
                <TextBlock Text="Select OBS Studio Location"
                     FontSize="20"
                     FontWeight="SemiBold"
                     Foreground="White"
                     Margin="0,0,0,16"/>
               
                <TextBlock Text="Choose the location of your OBS Studio Portable installation"
                     Foreground="#E4E4E7"
                     Margin="0,0,0,16"/>
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <Border Background="#27272A"
                      CornerRadius="6"
                      BorderThickness="1"
                      BorderBrush="#3F3F46"
                      Padding="16,12">
                    <TextBlock x:Name="ObsLocationText"
                         Text="Searching..."
                         Foreground="#E4E4E7"/>
                  </Border>
                  <Button x:Name="BrowseButton"
                      Grid.Column="1"
                      Content="Browse"
                      Width="100"
                      Style="{StaticResource ModernButtonStyle}"
                      Margin="12,0,0,0"/>
                </Grid>
              </StackPanel>
            </Grid>
          </StackPanel>
        </Grid>
        <Grid x:Name="InstallPage" Visibility="Collapsed">
          <StackPanel>
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
             
              <TextBlock Text="2" Style="{StaticResource StepNumberStyle}"/>
             
              <StackPanel Grid.Column="1">
                <TextBlock Text="Installing Plugin"
                     FontSize="20"
                     FontWeight="SemiBold"
                     Foreground="White"
                     Margin="0,0,0,16"/>
               
                <TextBlock x:Name="InstallStatus"
                     Text="Preparing to install..."
                     Foreground="#E4E4E7"
                     Margin="0,0,0,16"/>
                <ProgressBar x:Name="InstallProgressBar"
                     Height="6"
                     Background="#27272A"
                     Foreground="#2563EB"
                     BorderThickness="0"/>
              </StackPanel>
            </Grid>
          </StackPanel>
        </Grid>
        <Grid x:Name="CompletePage" Visibility="Collapsed">
          <StackPanel>
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
             
              <TextBlock Text="✓"
                   Style="{StaticResource StepNumberStyle}"
                   Foreground="#22C55E"/>
             
              <StackPanel Grid.Column="1">
                <TextBlock Text="Installation Complete"
                     FontSize="20"
                     FontWeight="SemiBold"
                     Foreground="White"
                     Margin="0,0,0,16"/>
               
                <TextBlock x:Name="CompleteStatus"
                     Text="The Stream Deck plugin has been successfully installed."
                     Foreground="#E4E4E7"
                     Margin="0,0,0,16"/>
              </StackPanel>
            </Grid>
          </StackPanel>
        </Grid>
      </Grid>
      <!-- Navigation Bar -->
      <Grid Grid.Row="2"
         Background="#18181B"
         Margin="0">
        <Border Margin="24">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <Button x:Name="BackButton"
                Content="Back"
                Style="{StaticResource NavigationButtonStyle}"
                Background="Transparent"
                HorizontalAlignment="Left"
                Visibility="Collapsed"/>
            <StackPanel Grid.Column="1"
                 Orientation="Horizontal"
                 HorizontalAlignment="Right">
              <Button x:Name="CancelButton"
                  Content="Cancel"
                  Style="{StaticResource NavigationButtonStyle}"
                  Background="Transparent"/>
             
              <Button x:Name="NextButton"
                  Content="Next"
                  Style="{StaticResource NavigationButtonStyle}"/>
            </StackPanel>
          </Grid>
        </Border>
      </Grid>
    </Grid>
  </Border>
</Window>
"@
    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
    # Get UI elements
    $welcomePage = $window.FindName("WelcomePage")
    $locationPage = $window.FindName("LocationPage")
    $installPage = $window.FindName("InstallPage")
    $completePage = $window.FindName("CompletePage")
    $backButton = $window.FindName("BackButton")
    $nextButton = $window.FindName("NextButton")
    $cancelButton = $window.FindName("CancelButton")
    $closeButton = $window.FindName("CloseButton")
    $browseButton = $window.FindName("BrowseButton")
    $obsLocationText = $window.FindName("ObsLocationText")
    $installStatus = $window.FindName("InstallStatus")
    $completeStatus = $window.FindName("CompleteStatus")
    $progressBar = $window.FindName("InstallProgressBar")
    # Window dragging
    $window.Add_MouseLeftButtonDown({ $window.DragMove() })
    # Create synchronized hashtable
    $syncHash = [hashtable]::Synchronized(@{
            Window         = $window
            CurrentPage    = "Welcome"
            ObsLocation    = $null
            InstallStatus  = $installStatus
            ProgressBar    = $progressBar
            CompleteStatus = $completeStatus
            Completed      = $false
            Cancelled      = $false
            RunspacePool   = [runspacefactory]::CreateRunspacePool(1, 3)
        })
    $syncHash.RunspacePool.Open()
    # Navigation logic
    function Set-CurrentPage {
        param([string]$page)
   
        $welcomePage.Visibility = "Collapsed"
        $locationPage.Visibility = "Collapsed"
        $installPage.Visibility = "Collapsed"
        $completePage.Visibility = "Collapsed"
   
        switch ($page) {
            "Welcome" {
                $welcomePage.Visibility = "Visible"
                $backButton.Visibility = "Collapsed"
                $nextButton.Content = "Next"
            }
            "Location" {
                $locationPage.Visibility = "Visible"
                $backButton.Visibility = "Visible"
                $nextButton.Content = "Install"
                $nextButton.IsEnabled = ($null -ne $syncHash.ObsLocation)
            }
            "Install" {
                $installPage.Visibility = "Visible"
                $backButton.Visibility = "Collapsed"
                $nextButton.Visibility = "Collapsed"
                $cancelButton.Visibility = "Collapsed"
            }
            "Complete" {
                $completePage.Visibility = "Visible"
                $backButton.Visibility = "Collapsed"
                $nextButton.Content = "Finish"
                $cancelButton.Visibility = "Collapsed"
            }
        }
        $syncHash.CurrentPage = $page
    }
    # Button event handlers
    $nextButton.Add_Click({
            switch ($syncHash.CurrentPage) {
                "Welcome" { Set-CurrentPage "Location" }
                "Location" {
                    Set-CurrentPage "Install"
                    Start-Installation
                }
                "Complete" { $window.Close() }
            }
        })
    $backButton.Add_Click({
            switch ($syncHash.CurrentPage) {
                "Location" { Set-CurrentPage "Welcome" }
            }
        })
    $cancelButton.Add_Click({ $window.Close() })
    $closeButton.Add_Click({ $window.Close() })
    # Browse button handler
    $browseButton.Add_Click({
            $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
            $folderBrowser.Description = "Select OBS Studio Installation Folder"
            if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $selectedPath = $folderBrowser.SelectedPath
                if (Test-Path "$selectedPath\bin\64bit\obs64.exe") {
                    $syncHash.ObsLocation = $selectedPath
                    $obsLocationText.Text = $selectedPath
                    $nextButton.IsEnabled = $true
                }
                else {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Selected folder does not contain a valid OBS Studio installation.nPlease select the root folder of OBS Studio.",
                        "Invalid Selection",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Warning
                    )
                    $nextButton.IsEnabled = $false
                }
            }
        })
    # Auto-detection of OBS installation
    $window.Add_SourceInitialized({
            $searchJob = [powershell]::Create().AddScript({
                    param($syncHash)
                    function Find-ObsInstallation {
                        $commonPaths = @(
                            "${env:ProgramFiles}\obs-studio",
                            "${env:ProgramFiles(x86)}\obs-studio",
                            "$env:APPDATA\obs-studio",
                            "$env:LocalAppData\obs-studio"
                        )
                        function Test-ObsRoot {
                            param ([string]$Path)
                            $requiredItems = @(
                                "bin\64bit\obs64.exe",
                                "data",
                                "obs-plugins"
                            )
                            foreach ($item in $requiredItems) {
                                if (-not (Test-Path (Join-Path $Path $item))) {
                                    return $false
                                }
                            }
                            return $true
                        }
                        # Check common paths first
                        foreach ($path in $commonPaths) {
                            if (Test-Path $path -and (Test-ObsRoot $path)) {
                                return $path
                            }
                        }
                        # Search in common drive locations
                        $searchLocations = @(
                            [Environment]::GetFolderPath('ProgramFiles'),
                            [Environment]::GetFolderPath('ProgramFilesX86'),
                            [Environment]::GetFolderPath('LocalApplicationData'),
                            [Environment]::GetFolderPath('ApplicationData')
                        )
                        $searchLocations += Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType = 3" |
                        Select-Object -ExpandProperty DeviceID |
                        ForEach-Object { "$_\" }
                        foreach ($location in $searchLocations) {
                            try {
                                $obsExes = Get-ChildItem -Path $location -File -Filter "obs64.exe" -Depth 5 -ErrorAction SilentlyContinue |
                                Where-Object { $_.FullName -like "*\bin\64bit\obs64.exe" }
                                foreach ($obsExe in $obsExes) {
                                    $potentialRoot = Split-Path (Split-Path (Split-Path $obsExe.FullName -Parent) -Parent) -Parent
                                    if (Test-ObsRoot $potentialRoot) {
                                        return $potentialRoot
                                    }
                                }
                            }
                            catch {
                                continue
                            }
                        }
                        return $null
                    }
                    $obsLocation = Find-ObsInstallation
                    $syncHash.Window.Dispatcher.Invoke([Action] {
                            if ($obsLocation) {
                                $syncHash.ObsLocation = $obsLocation
                                $obsLocationText = $syncHash.Window.FindName("ObsLocationText")
                                $nextButton = $syncHash.Window.FindName("NextButton")
         
                                $obsLocationText.Text = $obsLocation
                                $nextButton.IsEnabled = $true
                            }
                            else {
                                $obsLocationText = $syncHash.Window.FindName("ObsLocationText")
                                $obsLocationText.Text = "No OBS Studio installation found. Please browse manually."
                            }
                        })
                }).AddArgument($syncHash)
            $searchJob.RunspacePool = $syncHash.RunspacePool
            $searchJob.BeginInvoke()
        })
    # Installation logic
    function Start-Installation {
        $installJob = [powershell]::Create().AddScript({
                param($syncHash)
                try {
                    $obsPortableFolder = $syncHash.ObsLocation
                    # Define source and target paths
                    $obsPluginsTarget = "${obsPortableFolder}\obs-plugins\64bit"
                    $obsDataPluginsTarget = "${obsPortableFolder}\data\obs-plugins\StreamDeckPlugin"
                    $programDataBase = "C:\ProgramData\obs-studio\plugins\StreamDeckPlugin"
                    $filesToCopy = @(
                        @{
                            Source = "$programDataBase\bin\64bit\StreamDeckPlugin.dll"
                            Target = $obsPluginsTarget
                        },
                        @{
                            Source = "$programDataBase\bin\64bit\StreamDeckPlugin.pdb"
                            Target = $obsPluginsTarget
                        },
                        @{
                            Source = "$programDataBase\Data\StreamDeckPluginQt6.dll"
                            Target = $obsDataPluginsTarget
                        },
                        @{
                            Source = "$programDataBase\Data\StreamDeckPluginQt6.pdb"
                            Target = $obsDataPluginsTarget
                        },
                        @{
                            Source = "$programDataBase\Data\Locale"
                            Target = $obsDataPluginsTarget
                        }
                    )
                    $totalFiles = $filesToCopy.Count
                    $successfulCopies = 0
                    $failedCopies = 0
                    foreach ($file in $filesToCopy) {
                        $syncHash.Window.Dispatcher.Invoke([Action] {
                                $syncHash.InstallStatus.Text = "Copying: $($file.Source)"
                                $syncHash.ProgressBar.Value = ($successfulCopies / $totalFiles) * 100
                            })
                        try {
                            if (Test-Path $file.Source) {
                                if (!(Test-Path $file.Target)) {
                                    New-Item -ItemType Directory -Force -Path $file.Target | Out-Null
                                }
                                Copy-Item -Path $file.Source -Destination $file.Target -Force -Recurse
                                $successfulCopies++
                            }
                            else {
                                $failedCopies++
                            }
                        }
                        catch {
                            $failedCopies++
                        }
                    }
                    $syncHash.Window.Dispatcher.Invoke([Action] {
                            if ($failedCopies -eq 0) {
                                $syncHash.CompleteStatus.Text = "The Stream Deck plugin has been successfully installed to your OBS Studio installation."
                            }
                            else {
                                $syncHash.CompleteStatus.Text = "Installation completed with some issues. Successfully copied $successfulCopies out of $totalFiles files."
                            }
                            Set-CurrentPage "Complete"
                        })
                }
                catch {
                    $syncHash.Window.Dispatcher.Invoke([Action] {
                            $syncHash.CompleteStatus.Text = "Installation failed: $($_.Exception.Message)"
                            Set-CurrentPage "Complete"
                        })
                }
            }).AddArgument($syncHash)
        $installJob.RunspacePool = $syncHash.RunspacePool
        $installJob.BeginInvoke()
    }
    # Set initial page
    Set-CurrentPage "Welcome"
    # Show the window
    $window.ShowDialog() | Out-Null
    # Cleanup
    $syncHash.RunspacePool.Close()
    $syncHash.RunspacePool.Dispose()
}
StreamDeckOBSPortableInstaller