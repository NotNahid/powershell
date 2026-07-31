<#
    .SYNOPSIS
    Quick Navigator v1.0 - Hierarchical Explorer Speed Dial
    Customizable folder navigation from the Windows Right-Click Menu.
#>

param (
    [Parameter(Mandatory=$false)]
    [ValidateSet("Open", "Settings", "Apply", "Uninstall")]
    [String]$Action,

    [Parameter(Mandatory=$false)]
    [String]$Path
)

# --- HIDE CONSOLE WINDOW ---
# This hides the background terminal when the GUI is running
if ($Action -eq $null -or $Action -eq "Settings") {
    $win32Type = Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);' -Name "Win32ShowWindow" -Namespace Win32 -PassThru -ErrorAction SilentlyContinue
    $hwnd = (Get-Process -Id $pid).MainWindowHandle
    if ($hwnd -ne 0) {
        $win32Type::ShowWindow($hwnd, 0) # 0 = SW_HIDE
    }
}

# --- AUTO-ELEVATION ---
$NeedsAdmin = ($Action -in @("Apply", "Uninstall"))
if ($NeedsAdmin -and -not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`" -Action $Action" -Verb RunAs
    exit
}

# --- Configuration & Paths ---
$ConfigFile = Join-Path (Split-Path $PSCommandPath) "QuickNavigator.json"
$AppTitle = "Quick Navigator Manager"

function Load-Config {
    if (Test-Path $ConfigFile) {
        return Get-Content $ConfigFile | ConvertFrom-Json
    }
    return @{ Categories = @() }
}

function Save-Config ($Data) {
    $Data | ConvertTo-Json -Depth 10 | Out-File $ConfigFile
}

# --- Registry Engine ---
function Update-RegistryMenu {
    $Data = Load-Config
    $ScriptPath = $PSCommandPath
    $ActionBase = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`""
    
    $RootPaths = @(
        "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\QuickNavigator",
        "Registry::HKEY_CLASSES_ROOT\DesktopBackground\shell\QuickNavigator",
        "Registry::HKEY_CLASSES_ROOT\Directory\shell\QuickNavigator",
        "Registry::HKEY_CLASSES_ROOT\*\shell\QuickNavigator"
    )

    foreach ($RootKey in $RootPaths) {
        # 1. Clean Up
        if (Test-Path $RootKey) { Remove-Item $RootKey -Recurse -Force -ErrorAction SilentlyContinue }
        
        # 2. Create Main Entry
        $MainKey = New-Item -Path $RootKey -Force
        New-ItemProperty -Path $MainKey.PSPath -Name "MUIVerb" -Value "🚀 Quick Navigator" -Force | Out-Null
        New-ItemProperty -Path $MainKey.PSPath -Name "Icon" -Value "shell32.dll,20" -Force | Out-Null
        New-ItemProperty -Path $MainKey.PSPath -Name "SubCommands" -Value "" -Force | Out-Null
        
        $ShellKey = New-Item -Path "$($MainKey.PSPath)\shell" -Force
        
        # 3. Add Categories & Folders
        foreach ($Cat in @($Data.Categories)) {
            if ($null -eq $Cat.Name) { continue }
            $CatKey = New-Item -Path "$($ShellKey.PSPath)\$($Cat.Name)" -Force
            New-ItemProperty -Path $CatKey.PSPath -Name "MUIVerb" -Value $Cat.Name -Force | Out-Null
            New-ItemProperty -Path $CatKey.PSPath -Name "SubCommands" -Value "" -Force | Out-Null
            
            $CatShell = New-Item -Path "$($CatKey.PSPath)\shell" -Force
            
            foreach ($Folder in @($Cat.Folders)) {
                if ($null -eq $Folder.Name) { continue }
                $FolderKey = New-Item -Path "$($CatShell.PSPath)\$($Folder.Name)" -Force
                New-ItemProperty -Path $FolderKey.PSPath -Name "MUIVerb" -Value $Folder.Name -Force | Out-Null
                New-Item -Path "$($FolderKey.PSPath)\command" -Value "$ActionBase -Action Open -Path `"$($Folder.Path)`"" -Force | Out-Null
            }
        }
        
        # 4. Add Settings Entry
        $SettingsKeyPath = "$($ShellKey.Name)\zQuickNavSettings"
        $SettingsKey = New-Item -Path $SettingsKeyPath -Force
        New-ItemProperty -Path $SettingsKey.PSPath -Name "MUIVerb" -Value "⚙️ Settings" -Force | Out-Null
        New-ItemProperty -Path $SettingsKey.PSPath -Name "CommandFlags" -Value 32 -Force | Out-Null
        New-Item -Path "$($SettingsKey.PSPath)\command" -Value "$ActionBase -Action Settings" -Force | Out-Null
    }
}

# --- Management GUI ---
function Show-Manager {
    try {
        Add-Type -AssemblyName PresentationFramework
        Add-Type -AssemblyName PresentationCore
        Add-Type -AssemblyName WindowsBase
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName Microsoft.VisualBasic
    } catch {}

    $Global:Config = Load-Config

    [xml]$Xaml = @"
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            Title="$AppTitle" Height="650" Width="800" WindowStartupLocation="CenterScreen" Background="#121212">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <!-- Header -->
            <Border Grid.Row="0" Background="#1E1E1E" Padding="20">
                <StackPanel>
                    <TextBlock Text="🚀 QUICK NAVIGATOR SETUP" Foreground="#0099FF" FontSize="28" FontWeight="Bold" HorizontalAlignment="Center">
                        <TextBlock.Effect>
                            <DropShadowEffect BlurRadius="5" Color="#0099FF" ShadowDepth="0" Opacity="0.5"/>
                        </TextBlock.Effect>
                    </TextBlock>
                    <TextBlock Text="Build your custom right-click speed dial in 3 easy steps." Foreground="#CCCCCC" HorizontalAlignment="Center" Margin="0,5,0,0"/>
                </StackPanel>
            </Border>

            <!-- Steps -->
            <TabControl Name="TabSteps" Grid.Row="1" Background="Transparent" BorderThickness="0" Margin="10">
                <TabControl.Resources>
                    <Style TargetType="TabItem">
                        <Setter Property="Template">
                            <Setter.Value>
                                <ControlTemplate TargetType="TabItem">
                                    <Border Name="Border" BorderBrush="#333333" BorderThickness="0,0,0,2" Margin="0,0,10,0" Padding="15,8">
                                        <ContentPresenter x:Name="ContentSite" ContentSource="Header" VerticalAlignment="Center" HorizontalAlignment="Center"/>
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsSelected" Value="True">
                                            <Setter TargetName="Border" Property="BorderBrush" Value="#0099FF"/>
                                            <Setter Property="Foreground" Value="#0099FF"/>
                                        </Trigger>
                                        <Trigger Property="IsSelected" Value="False">
                                            <Setter Property="Foreground" Value="#AAAAAA"/>
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Setter.Value>
                        </Setter>
                        <Setter Property="FontSize" Value="14"/>
                        <Setter Property="FontWeight" Value="Bold"/>
                    </Style>
                </TabControl.Resources>

                <!-- Step 1: Categories -->
                <TabItem Header="STEP 1: CATEGORIES">
                    <Grid Margin="20">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" Text="Create main groups for your shortcuts (e.g., 'Work', 'Projects', 'Social')." Foreground="#FFFFFF" TextWrapping="Wrap" Margin="0,0,0,15" FontSize="14"/>
                        <ListBox Name="ListCats" Grid.Row="1" Background="#1E1E1E" Foreground="#FFFFFF" BorderBrush="#333333" BorderThickness="1" FontSize="16" Padding="5"/>
                        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
                            <Button Name="BtnDelCat" Content="REMOVE" Width="100" Height="35" Background="#333333" Foreground="#FFFFFF" Margin="0,0,10,0"/>
                            <Button Name="BtnAddCat" Content="+ ADD CATEGORY" Width="150" Height="35" Background="#0099FF" Foreground="#FFFFFF" FontWeight="Bold"/>
                        </StackPanel>
                    </Grid>
                </TabItem>

                <!-- Step 2: Folders -->
                <TabItem Header="STEP 2: ADD FOLDERS">
                    <Grid Margin="20">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" Text="Select a category on the left, then add the folders you want inside it." Foreground="#FFFFFF" Margin="0,0,0,15" FontSize="14"/>
                        <Grid Grid.Row="1">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="220"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <ListBox Name="ListCatsFolders" Background="#181818" Foreground="#FFFFFF" BorderBrush="#333333" BorderThickness="1" Margin="0,0,10,0" Padding="5"/>
                            <ListBox Name="ListFolders" Grid.Column="1" Background="#1E1E1E" Foreground="#FFFFFF" BorderBrush="#333333" BorderThickness="1" Padding="5">
                                <ListBox.ItemTemplate>
                                    <DataTemplate>
                                        <StackPanel Margin="5">
                                            <TextBlock Text="{Binding Name}" FontWeight="Bold" Foreground="#FFFFFF" FontSize="15"/>
                                            <TextBlock Text="{Binding Path}" Foreground="#AAAAAA" FontSize="11" Margin="0,2,0,0"/>
                                        </StackPanel>
                                    </DataTemplate>
                                </ListBox.ItemTemplate>
                            </ListBox>
                        </Grid>
                        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
                            <Button Name="BtnDelFolder" Content="REMOVE" Width="100" Height="35" Background="#333333" Foreground="#FFFFFF" Margin="0,0,10,0"/>
                            <Button Name="BtnAddFolder" Content="+ ADD FOLDER" Width="150" Height="35" Background="#0099FF" Foreground="#FFFFFF" FontWeight="Bold"/>
                        </StackPanel>
                    </Grid>
                </TabItem>

                <!-- Step 3: Apply -->
                <TabItem Header="STEP 3: FINISH">
                    <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center" Margin="40">
                        <TextBlock Text="READY TO DEPLOY?" Foreground="#FFFFFF" FontSize="26" FontWeight="Bold" HorizontalAlignment="Center"/>
                        <TextBlock Text="Once you click the button below, your new Right-Click menu will be live instantly across your entire PC." 
                                   Foreground="#BBBBBB" TextWrapping="Wrap" TextAlignment="Center" Margin="0,20,0,40" Width="450" FontSize="14"/>
                        <Button Name="BtnApply" Content="🚀 APPLY CHANGES TO WINDOWS" Height="70" Width="400" Background="#0099FF" Foreground="#FFFFFF" FontWeight="Bold" FontSize="20">
                            <Button.Effect>
                                <DropShadowEffect BlurRadius="10" Color="#0099FF" ShadowDepth="0" Opacity="0.3"/>
                            </Button.Effect>
                        </Button>
                        <Button Name="BtnUninstall" Content="REMOVE ALL INTEGRATIONS" Height="35" Width="250" Background="#444444" Foreground="#CCCCCC" Margin="0,60,0,0"/>
                    </StackPanel>
                </TabItem>
            </TabControl>

            <!-- Footer -->
            <Border Grid.Row="2" Background="#0A0A0A" Padding="15">
                <TextBlock Text="Quick Navigator v1.0 • Settings saved to QuickNavigator.json" Foreground="#666666" HorizontalAlignment="Center" FontSize="10"/>
            </Border>
        </Grid>
    </Window>
"@
    
    $Reader = (New-Object System.Xml.XmlNodeReader $Xaml)
    $Window = [Windows.Markup.XamlReader]::Load($Reader)
    
    $ListCats = $Window.FindName("ListCats")
    $ListCatsFolders = $Window.FindName("ListCatsFolders")
    $ListFolders = $Window.FindName("ListFolders")
    $TabSteps = $Window.FindName("TabSteps")
    
    # Sync Lists
    function Update-UI {
        if ($null -eq $Global:Config.Categories) { $Global:Config.Categories = @() }
        
        $prevCatSelection = $ListCats.SelectedItem
        $prevCatFolderSelection = $ListCatsFolders.SelectedItem
        
        $catNames = @($Global:Config.Categories | ForEach-Object { $_.Name })
        $ListCats.ItemsSource = @($catNames)
        $ListCatsFolders.ItemsSource = @($catNames)
        
        if ($catNames -contains $prevCatSelection) { $ListCats.SelectedItem = $prevCatSelection }
        if ($catNames -contains $prevCatFolderSelection) { 
            $ListCatsFolders.SelectedItem = $prevCatFolderSelection 
        } else {
            $ListFolders.ItemsSource = $null
        }
    }
    Update-UI

    $ListCatsFolders.Add_SelectionChanged({
        if ($ListCatsFolders.SelectedItem) {
            $Cat = $Global:Config.Categories | Where-Object { $_.Name -eq $ListCatsFolders.SelectedItem } | Select-Object -First 1
            $ListFolders.ItemsSource = @($Cat.Folders)
        }
    })

    # Event: Add Category
    $Window.FindName("BtnAddCat").Add_Click({
        $Name = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Category Name (e.g., Images, Work, Tools):", "New Category")
        if ($Name) {
            $newCat = [PSCustomObject]@{ Name = $Name; Folders = @() }
            $Global:Config.Categories = @($Global:Config.Categories) + $newCat
            Update-UI
        }
    })

    # Event: Remove Category
    $Window.FindName("BtnDelCat").Add_Click({
        if ($ListCats.SelectedItem) {
            $Global:Config.Categories = $Global:Config.Categories | Where-Object { $_.Name -ne $ListCats.SelectedItem }
            Update-UI
        }
    })

    # Event: Add Folder
    $Window.FindName("BtnAddFolder").Add_Click({
        $selectedName = $ListCatsFolders.SelectedItem
        if (-not $selectedName) { 
            [System.Windows.MessageBox]::Show("Please select a category on the left first!", "Selection Required")
            return 
        }
        
        $FolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        if ($FolderDialog.ShowDialog() -eq "OK") {
            $Name = [Microsoft.VisualBasic.Interaction]::InputBox("Friendly name for this shortcut:", "Folder Name", (Split-Path $FolderDialog.SelectedPath -Leaf))
            if ($Name) {
                $Cat = $Global:Config.Categories | Where-Object { $_.Name -eq $selectedName } | Select-Object -First 1
                $Cat.Folders = @($Cat.Folders) + @{ Name = $Name; Path = $FolderDialog.SelectedPath }
                
                # Refresh List
                $ListFolders.ItemsSource = $null
                $ListFolders.ItemsSource = @($Cat.Folders)
            }
        }
    })

    # Event: Remove Folder
    $Window.FindName("BtnDelFolder").Add_Click({
        if ($ListCatsFolders.SelectedItem -and $ListFolders.SelectedItem) {
            $Cat = $Global:Config.Categories | Where-Object { $_.Name -eq $ListCatsFolders.SelectedItem } | Select-Object -First 1
            $selectedFolder = $ListFolders.SelectedItem
            $Cat.Folders = $Cat.Folders | Where-Object { $_.Name -ne $selectedFolder.Name -or $_.Path -ne $selectedFolder.Path }
            
            $ListFolders.ItemsSource = $null
            $ListFolders.ItemsSource = @($Cat.Folders)
        }
    })

    # Event: Apply
    $Window.FindName("BtnApply").Add_Click({
        Save-Config $Global:Config
        # Run the apply action completely hidden
        Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`" -Action Apply" -Verb RunAs -WindowStyle Hidden
        [System.Windows.MessageBox]::Show("🚀 Changes applied! Your new Right-Click menu is live.", "Success", "OK", "Information")
    })

    # Event: Uninstall
    $Window.FindName("BtnUninstall").Add_Click({
        $Confirm = [System.Windows.MessageBox]::Show("Are you sure you want to remove Quick Navigator from your PC?", "Confirm Uninstall", "YesNo", "Warning")
        if ($Confirm -eq "Yes") {
            Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`" -Action Uninstall" -Verb RunAs -WindowStyle Hidden
        }
    })

    $Window.ShowDialog() | Out-Null
}

# --- Router ---
if ($Action -eq "Open") {
    explorer.exe $Path
}
elseif ($Action -eq "Settings") {
    Show-Manager
}
elseif ($Action -eq "Apply") {
    Update-RegistryMenu
}
elseif ($Action -eq "Uninstall") {
    $RootPaths = @(
        "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\QuickNavigator",
        "Registry::HKEY_CLASSES_ROOT\DesktopBackground\shell\QuickNavigator",
        "Registry::HKEY_CLASSES_ROOT\Directory\shell\QuickNavigator",
        "Registry::HKEY_CLASSES_ROOT\*\shell\QuickNavigator"
    )
    foreach ($Path in $RootPaths) {
        Remove-Item $Path.Trim() -Recurse -Force -ErrorAction SilentlyContinue
    }
}
else {
    # Default: Show Manager
    Show-Manager
}
