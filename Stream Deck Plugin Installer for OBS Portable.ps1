param([switch]$NoGui)
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

function StreamDeckOBSPortableInstaller {
  param([switch]$NoGui)
  if ($NoGui) {
    Write-Host "Running in no-GUI debug mode..." -ForegroundColor Cyan
      # Automatic OBS location detection
      function Find-ObsInstallation {
        $commonPaths = @(
          "$env:ProgramFiles\obs-studio",
          "$env:ProgramFiles(x86)\obs-studio",
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
        foreach ($path in $commonPaths) {
          if ((Test-Path $path) -and (Test-ObsRoot $path)) {
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
        } catch {
          continue
        }
      }
        return $null
      }
      function Start-Installation-NoGui {
        $programDataBase = "C:\ProgramData\obs-studio\plugins\StreamDeckPlugin"
        $obsPortableFolder = Find-ObsInstallation
        if ($obsPortableFolder) {
          Write-Host "Detected OBS Studio Portable location: $obsPortableFolder" -ForegroundColor Cyan
        } else {
          $obsPortableFolder = Read-Host "Enter the path to your OBS Studio Portable installation folder"
        }
        $obsPluginsTarget = "$obsPortableFolder\obs-plugins\64bit"
        $obsDataPluginsTarget = "$obsPortableFolder\data\obs-plugins\StreamDeckPlugin"
        $filesToCopy = @(
          @{ Source = "$programDataBase\bin\64bit\StreamDeckPlugin.dll"; Target = $obsPluginsTarget },
          @{ Source = "$programDataBase\bin\64bit\StreamDeckPlugin.pdb"; Target = $obsPluginsTarget },
          @{ Source = "$programDataBase\Data\StreamDeckPluginQt6.dll"; Target = $obsDataPluginsTarget },
          @{ Source = "$programDataBase\Data\StreamDeckPluginQt6.pdb"; Target = $obsDataPluginsTarget },
          @{ Source = "$programDataBase\Data\Locale"; Target = $obsDataPluginsTarget }
        )
        $obs32Dll = "$programDataBase\Data\StreamDeckPluginOBS32.dll"
        $obs32Pdb = "$programDataBase\Data\StreamDeckPluginOBS32.pdb"
        if (Test-Path $obs32Dll) { $filesToCopy += @(@{ Source = $obs32Dll; Target = $obsDataPluginsTarget }) }
        if (Test-Path $obs32Pdb) { $filesToCopy += @(@{ Source = $obs32Pdb; Target = $obsDataPluginsTarget }) }
        $totalFiles = $filesToCopy.Count
        $successfulCopies = 0
        $failedCopies = 0
        Write-Host "--- Installation started: $(Get-Date) ---" -ForegroundColor Yellow
        Write-Host "Total files to copy: $totalFiles" -ForegroundColor Yellow
        foreach ($file in $filesToCopy) {
          $logMsg = "Attempting to copy: $($file.Source) to $($file.Target)"
          Write-Host $logMsg -ForegroundColor Gray
          try {
            if (Test-Path $file.Source) {
              $targetDir = $file.Target
              if ($file.Source -notlike "*Locale") { $targetDir = Split-Path $file.Target -Parent }
              if (!(Test-Path $targetDir)) { New-Item -ItemType Directory -Force -Path $targetDir | Out-Null }
              Copy-Item -Path $file.Source -Destination $file.Target -Force -Recurse
              $successfulCopies++
              $logMsg = "Successfully copied: $($file.Source)"
              Write-Host $logMsg -ForegroundColor Green
            } else {
              $failedCopies++
              $logMsg = "Error: Source file not found: $($file.Source)"
              Write-Host $logMsg -ForegroundColor Red
            }
          } catch {
            $failedCopies++
            $logMsg = "Error copying $($file.Source): $($_.Exception.Message)"
            Write-Host $logMsg -ForegroundColor Red
          }
        }
        $summaryMsg = "Successfully copied $successfulCopies out of $totalFiles files."
        if ($failedCopies -eq 0) {
          Write-Host "The Stream Deck plugin has been successfully installed to your OBS Studio installation." -ForegroundColor Green
        } else {
          Write-Host "Installation completed with some issues. $summaryMsg" -ForegroundColor Yellow
        }
      }
      Start-Installation-NoGui
      return
    }
    # GUI code only runs if -not $NoGui
    if (-not $NoGui) {
      # Inline XAML (self-contained). Keep the terminator '@ on column 1 with no leading whitespace.
      $xaml = @'
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
            <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="6" Padding="20,8">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="border" Property="Background" Value="#1D4ED8"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="border" Property="Background" Value="#27272A"/><Setter Property="Foreground" Value="#52525B"/></Trigger>
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

  <Border Background="#171717" BorderBrush="#27272A" BorderThickness="1" CornerRadius="12" Padding="0">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>
      <!-- Title Bar -->
      <Grid Grid.Row="0" Margin="24,24,24,0" Background="Transparent" x:Name="TitleBar">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel>
          <TextBlock Text="Stream Deck OBS Plugin Installer" FontSize="24" FontWeight="Bold" Foreground="White"/>
          <TextBlock Text="Install the Stream Deck plugin for OBS Studio in portable mode" Foreground="#71717A" Margin="0,4,0,0"/>
        </StackPanel>
        <Button x:Name="CloseButton" Grid.Column="1" Content="✕" Width="32" Height="32" Background="Transparent" Foreground="#71717A" BorderThickness="0" Cursor="Hand"/>
      </Grid>

      <Grid Grid.Row="1" Margin="24">
        <!-- Welcome -->
        <Grid x:Name="WelcomePage" Visibility="Visible">
          <StackPanel>
            <TextBlock Text="Before we begin" FontSize="20" FontWeight="SemiBold" Foreground="White" Margin="0,0,0,16"/>
            <TextBlock Text="Please ensure you have:" Foreground="#E4E4E7" Margin="0,0,0,12"/>
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

        <!-- Location -->
        <Grid x:Name="LocationPage" Visibility="Collapsed">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="1" Style="{StaticResource StepNumberStyle}"/>
            <StackPanel Grid.Column="1">
              <TextBlock Text="Select OBS Studio Location" FontSize="20" FontWeight="SemiBold" Foreground="White" Margin="0,0,0,16"/>
              <TextBlock Text="Choose the location of your OBS Studio Portable installation" Foreground="#E4E4E7" Margin="0,0,0,16"/>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Border Background="#27272A" CornerRadius="6" BorderThickness="1" BorderBrush="#3F3F46" Padding="16,12">
                  <TextBlock x:Name="ObsLocationText" Text="Searching..." Foreground="#E4E4E7"/>
                </Border>
                <Button x:Name="BrowseButton" Grid.Column="1" Content="Browse" Width="100" Style="{StaticResource ModernButtonStyle}" Margin="12,0,0,0"/>
              </Grid>
            </StackPanel>
          </Grid>
        </Grid>

        <!-- Install -->
        <Grid x:Name="InstallPage" Visibility="Collapsed">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="2" Style="{StaticResource StepNumberStyle}"/>
            <Grid Grid.Column="1">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <!-- Title -->
              <StackPanel Orientation="Vertical" Grid.Row="0" Margin="0,0,0,4">
                <TextBlock Text="Installing Plugin" FontSize="20" FontWeight="SemiBold" Foreground="White"/>
              </StackPanel>
              <!-- Status line separated from title to avoid overlap -->
              <TextBlock x:Name="InstallStatus" Text="Preparing to install..." Foreground="#E4E4E7" Margin="0,0,0,12" Grid.Row="1"/>
              <!-- File list fills the body; transparent scroll when needed so UI stays integrated -->
              <ScrollViewer Grid.Row="2" VerticalScrollBarVisibility="Auto" Background="Transparent" BorderThickness="0" Padding="0">
                <StackPanel x:Name="FileListPanel" Margin="0"/>
              </ScrollViewer>
              <!-- Progress bar stays pinned to bottom of area -->
              <ProgressBar x:Name="InstallProgressBar" Height="6" Background="#27272A" Foreground="#2563EB" BorderThickness="0" Grid.Row="3"/>
            </Grid>
          </Grid>
        </Grid>

        <!-- Complete -->
        <Grid x:Name="CompletePage" Visibility="Collapsed">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="✓" Style="{StaticResource StepNumberStyle}" Foreground="#22C55E"/>
            <StackPanel Grid.Column="1">
              <TextBlock Text="Installation Complete" FontSize="20" FontWeight="SemiBold" Foreground="White" Margin="0,0,0,16"/>
              <ScrollViewer Height="160" VerticalScrollBarVisibility="Auto" Background="Transparent" BorderThickness="0" Margin="0,0,8,12">
                <StackPanel x:Name="ResultFileListPanel"/>
              </ScrollViewer>
              <TextBlock x:Name="CompleteStatus" Text="The Stream Deck plugin has been successfully installed." Foreground="#E4E4E7" Margin="0,0,0,16"/>
            </StackPanel>
          </Grid>
        </Grid>
      </Grid>

      <!-- Navigation Bar -->
      <Grid Grid.Row="2" Background="#18181B" Margin="0">
        <Border Margin="24">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <Button x:Name="BackButton" Content="Back" Style="{StaticResource NavigationButtonStyle}" Background="Transparent" HorizontalAlignment="Left" Visibility="Collapsed"/>
            <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
              <Button x:Name="CancelButton" Content="Cancel" Style="{StaticResource NavigationButtonStyle}" Background="Transparent"/>
              <Button x:Name="NextButton" Content="Next" Style="{StaticResource NavigationButtonStyle}"/>
            </StackPanel>
          </Grid>
        </Border>
      </Grid>
    </Grid>
  </Border>
</Window>
'@
      $xmlDoc = New-Object System.Xml.XmlDocument
      $xmlDoc.LoadXml($xaml)
      $reader = New-Object System.Xml.XmlNodeReader $xmlDoc
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
              $nextButton.Visibility = "Visible"
            }
            "Location" {
                $locationPage.Visibility = "Visible"
                $backButton.Visibility = "Visible"
              $nextButton.Content = "Install"
              $nextButton.IsEnabled = ($null -ne $syncHash.ObsLocation)
              $nextButton.Visibility = "Visible"
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
              $nextButton.Visibility = "Visible"
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
                            if ((Test-Path $path) -and (Test-ObsRoot $path)) {
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
                    # Add support for OBS Studio 32+ files if present
                    $obs32Dll = "$programDataBase\Data\StreamDeckPluginOBS32.dll"
                    $obs32Pdb = "$programDataBase\Data\StreamDeckPluginOBS32.pdb"
                    if (Test-Path $obs32Dll) {
                      $filesToCopy += @(@{ Source = $obs32Dll; Target = $obsDataPluginsTarget })
                    }
                    if (Test-Path $obs32Pdb) {
                      $filesToCopy += @(@{ Source = $obs32Pdb; Target = $obsDataPluginsTarget })
                    }
                    $totalFiles = $filesToCopy.Count
                    $successfulCopies = 0
                    $failedCopies = 0
                    $logFile = "$env:TEMP\StreamDeckOBSInstaller.log"
                    Add-Content -Path $logFile -Value "--- Installation started: $(Get-Date) ---"
                    # Prepare file list UI: add one child per expected file so we can mark them individually
                    $syncHash.Window.Dispatcher.Invoke([Action] {
                      $w = $syncHash.Window
                      if ($w.FindName('FileListPanel')) {
                        $panel = $w.FindName('FileListPanel')
                        $panel.Children.Clear()
                        for ($j = 0; $j -lt $filesToCopy.Count; $j++) {
                          $srcName = [System.IO.Path]::GetFileName($filesToCopy[$j].Source)
                          $tb = New-Object System.Windows.Controls.TextBlock
                          $tb.Text = "○ $srcName"
                          $tb.Foreground = (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(228,228,231)))
                          $tb.Margin = '0,4,0,0'
                          $tb.FontSize = 13
                          $panel.Children.Add($tb)
                        }
                      }
                    })
                    $fileIndex = 0
                    foreach ($file in $filesToCopy) {
                      $syncHash.Window.Dispatcher.Invoke([Action] {
                        $syncHash.InstallStatus.Text = "Copying: $($file.Source)"
                        $syncHash.ProgressBar.Value = ($successfulCopies / $totalFiles) * 100
                      })
                      Add-Content -Path $logFile -Value "Attempting to copy: $($file.Source) to $($file.Target)"
                      try {
                        if (Test-Path $file.Source) {
                          # If target is a file, create parent directory
                          $targetDir = $file.Target
                          if ($file.Source -notlike "*Locale") {
                            $targetDir = Split-Path $file.Target -Parent
                          }
                          if (!(Test-Path $targetDir)) {
                            New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
                          }
                          # Inline copy with retry/backoff (avoid Start-Job overhead so GUI stays fast)
                          $maxAttempts = 5
                          $isGui = $false
                          try { if ($syncHash.Window) { $isGui = $true } } catch { $isGui = $false }
                          $attempt = 0
                          $copySuccess = $false
                          $jobResult = $null
                          while ($attempt -lt $maxAttempts -and -not $copySuccess) {
                            try {
                              Copy-Item -Path $file.Source -Destination $file.Target -Force -Recurse -ErrorAction Stop
                              $copySuccess = $true
                              $jobResult = @{ Success = $true; Attempt = $attempt + 1 }
                            } catch {
                              $err = $_.Exception.Message
                              if ($err -match 'being used by another process' -or $err -match 'The process cannot access the file') {
                                if ($isGui) { Start-Sleep -Milliseconds 25 } else { Start-Sleep -Seconds (2 * ($attempt + 1)) }
                                $attempt++
                                continue
                              } else {
                                $jobResult = @{ Success = $false; Error = $err; Attempt = $attempt + 1 }
                                break
                              }
                            }
                          }
                          if (-not $copySuccess -and $jobResult -eq $null) {
                            $jobResult = @{ Success = $false; Error = 'MaxAttemptsReached'; Attempt = $attempt }
                          }
                          if ($jobResult.Success) {
                            $successfulCopies++
                            $msg = "Successfully copied: $($file.Source) (attempt $($jobResult.Attempt))"
                            Add-Content -Path $logFile -Value $msg
                            $srcName = [System.IO.Path]::GetFileName($file.Source)
                            $syncHash.Window.Dispatcher.Invoke([Action] {
                              $syncHash.InstallStatus.Text = $msg
                              $syncHash.ProgressBar.Value = ($successfulCopies / $totalFiles) * 100
                              $panel = $syncHash.Window.FindName('FileListPanel')
                              if ($panel -and $panel.Children.Count -gt $fileIndex) {
                                $tb = $panel.Children[$fileIndex]
                                $tb.Text = "✔ $srcName"
                                $tb.Foreground = (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(88,172,61)))
                              }
                            })
                          } else {
                            $failedCopies++
                            $errText = $jobResult.Error
                            $msg = "Failed copying $($file.Source): $errText (attempt $($jobResult.Attempt))"
                            Add-Content -Path $logFile -Value $msg
                            $srcName = [System.IO.Path]::GetFileName($file.Source)
                            $syncHash.Window.Dispatcher.Invoke([Action] {
                              $syncHash.InstallStatus.Text = $msg
                              $panel = $syncHash.Window.FindName('FileListPanel')
                              if ($panel -and $panel.Children.Count -gt $fileIndex) {
                                $tb = $panel.Children[$fileIndex]
                                $tb.Text = "✖ $srcName"
                                $tb.Foreground = (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(224,36,94)))
                              }
                            })
                          }
                        } else {
                          $failedCopies++
                          $msg = "Error: Source file not found: $($file.Source)"
                          Add-Content -Path $logFile -Value $msg
                          $srcName = [System.IO.Path]::GetFileName($file.Source)
                          $syncHash.Window.Dispatcher.Invoke([Action] {
                            $syncHash.InstallStatus.Text = $msg
                            $panel = $syncHash.Window.FindName('FileListPanel')
                            if ($panel -and $panel.Children.Count -gt $fileIndex) {
                              $tb = $panel.Children[$fileIndex]
                              $tb.Text = "✖ $srcName"
                              $tb.Foreground = (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(224,36,94)))
                            }
                          })
                        }
                        $fileIndex++
                      } catch {
                        $failedCopies++
                        $errorMsg = $_.Exception.Message
                        $msg = "Error copying $($file.Source): $errorMsg"
                        Add-Content -Path $logFile -Value $msg
                        $syncHash.Window.Dispatcher.Invoke([Action] {
                          $syncHash.InstallStatus.Text = $msg
                        })
                        # removed artificial pause to keep GUI responsive and speed up no-GUI runs
                      }
                    }
                    $syncHash.Window.Dispatcher.Invoke([Action] {
                        if ($failedCopies -eq 0) {
                          $syncHash.CompleteStatus.Text = "The Stream Deck plugin has been successfully installed to your OBS Studio installation."
                        }
                        else {
                          $syncHash.CompleteStatus.Text = "Installation completed with some issues. Successfully copied $successfulCopies out of $totalFiles files."
                        }
                        # Copy the file list into the results panel so the user can review final statuses
                        $w = $syncHash.Window
                        $src = $w.FindName('FileListPanel')
                        $dst = $w.FindName('ResultFileListPanel')
                        if ($dst -and $src) {
                          $dst.Children.Clear()
                          foreach ($child in $src.Children) {
                            try {
                              $text = $child.Text
                              $brush = $child.Foreground
                              $ntb = New-Object System.Windows.Controls.TextBlock
                              $ntb.Text = $text
                              $ntb.Foreground = $brush
                              $ntb.Margin = '0,4,0,0'
                              $ntb.FontSize = 13
                              $dst.Children.Add($ntb)
                            } catch { continue }
                          }
                        }
                        # Directly switch pages to avoid cross-runspace function resolution issues
                        $w.FindName('WelcomePage').Visibility = 'Collapsed'
                        $w.FindName('LocationPage').Visibility = 'Collapsed'
                        $w.FindName('InstallPage').Visibility = 'Collapsed'
                        $w.FindName('CompletePage').Visibility = 'Visible'
                        $w.FindName('BackButton').Visibility = 'Collapsed'
                        $w.FindName('CancelButton').Visibility = 'Collapsed'
                            $w.FindName('NextButton').Content = 'Finish'
                            $w.FindName('NextButton').Visibility = 'Visible'
                        $syncHash.CurrentPage = 'Complete'
                      })
                }
                catch {
                  $syncHash.Window.Dispatcher.Invoke([Action] {
                      $syncHash.CompleteStatus.Text = "Installation failed: $($_.Exception.Message)"
                      # Ensure UI shows completion page for errors
                      $w = $syncHash.Window
                      # copy file list if available
                      $src = $w.FindName('FileListPanel')
                      $dst = $w.FindName('ResultFileListPanel')
                      if ($dst -and $src) {
                        $dst.Children.Clear()
                        foreach ($child in $src.Children) {
                          try {
                            $text = $child.Text
                            $brush = $child.Foreground
                            $ntb = New-Object System.Windows.Controls.TextBlock
                            $ntb.Text = $text
                            $ntb.Foreground = $brush
                            $ntb.Margin = '0,4,0,0'
                            $ntb.FontSize = 13
                            $dst.Children.Add($ntb)
                          } catch { continue }
                        }
                      }
                      $w.FindName('WelcomePage').Visibility = 'Collapsed'
                      $w.FindName('LocationPage').Visibility = 'Collapsed'
                      $w.FindName('InstallPage').Visibility = 'Collapsed'
                      $w.FindName('CompletePage').Visibility = 'Visible'
                      $w.FindName('BackButton').Visibility = 'Collapsed'
                      $w.FindName('CancelButton').Visibility = 'Collapsed'
                      $w.FindName('NextButton').Content = 'Finish'
                      $w.FindName('NextButton').Visibility = 'Visible'
                      $syncHash.CurrentPage = 'Complete'
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
}

StreamDeckOBSPortableInstaller -NoGui:$NoGui