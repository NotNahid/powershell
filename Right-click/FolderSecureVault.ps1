<#
    .SYNOPSIS
    Secure Vault v1.2 - Native Explorer Encryption Utility
    Allows locking and unlocking folders directly from the Windows Right-Click Menu.
#>

param (
    [Parameter(Mandatory=$false)]
    [ValidateSet("Lock", "Unlock", "Install", "Uninstall")]
    [String]$Action,

    [Parameter(Mandatory=$false)]
    [String]$Target
)

# --- AUTO-ELEVATION: Ensure script runs as Administrator ---
# We need Admin for Install/Uninstall actions.
$NeedsAdmin = ($Action -in @("Install", "Uninstall", $null))
if ($NeedsAdmin -and -not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $ArgList = "-ExecutionPolicy Bypass -File `"$PSCommandPath`""
    
    # Pass all bound parameters
    if ($PSBoundParameters.Count -gt 0) {
        foreach ($Key in $PSBoundParameters.Keys) {
            $ArgList += " -$Key `"$($PSBoundParameters[$Key])`""
        }
    }
    
    # Pass any unbound arguments
    if ($args.Count -gt 0) {
        $ArgList += " " + ($args -join " ")
    }

    Start-Process powershell.exe -ArgumentList $ArgList -Verb RunAs
    exit
}

# --- Configuration ---
$VaultExtension = ".vault"
$AppTitle = "Secure Vault"

# --- UI Helper: Password Prompt ---
function Get-Password {
    param([string]$Title, [string]$PromptText)

    try {
        Add-Type -AssemblyName PresentationFramework
        Add-Type -AssemblyName PresentationCore
        Add-Type -AssemblyName WindowsBase
    } catch {}

    [xml]$Xaml = @"
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            Title="$Title" Height="220" Width="400" WindowStartupLocation="CenterScreen" 
            ResizeMode="NoResize" Background="#1E1E1E" WindowStyle="None" AllowsTransparency="True" Topmost="True">
        <Border BorderBrush="#333333" BorderThickness="1" CornerRadius="8">
            <StackPanel Margin="30">
                <TextBlock Text="$PromptText" Foreground="#FFFFFF" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,20" HorizontalAlignment="Center"/>
                <PasswordBox Name="TxtPassword" Height="35" Background="#2D2D2D" Foreground="#FFFFFF" BorderBrush="#444444" VerticalContentAlignment="Center" Padding="5,0"/>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,20,0,0">
                    <Button Name="BtnOK" Content="CONFIRM" Width="120" Height="35" Background="#007ACC" Foreground="White" FontWeight="Bold" Cursor="Hand" Margin="0,0,10,0">
                        <Button.Resources>
                            <Style TargetType="Border">
                                <Setter Property="CornerRadius" Value="4"/>
                            </Style>
                        </Button.Resources>
                    </Button>
                    <Button Name="BtnCancel" Content="CANCEL" Width="120" Height="35" Background="#444444" Foreground="White" Cursor="Hand">
                        <Button.Resources>
                            <Style TargetType="Border">
                                <Setter Property="CornerRadius" Value="4"/>
                            </Style>
                        </Button.Resources>
                    </Button>
                </StackPanel>
            </StackPanel>
        </Border>
    </Window>
"@
    
    $Reader = (New-Object System.Xml.XmlNodeReader $Xaml)
    $Window = [Windows.Markup.XamlReader]::Load($Reader)
    
    $TxtPassword = $Window.FindName("TxtPassword")
    $BtnOK = $Window.FindName("BtnOK")
    $BtnCancel = $Window.FindName("BtnCancel")

    $Global:VaultPassResult = $null

    $BtnOK.Add_Click({
        $Global:VaultPassResult = $TxtPassword.Password
        $Window.Close()
    })

    $BtnCancel.Add_Click({
        $Window.Close()
    })

    $Window.ShowDialog() | Out-Null
    return $Global:VaultPassResult
}

# --- Core Cryptography ---
function New-Key {
    param([string]$Password, [byte[]]$Salt)
    $PBKDF2 = New-Object System.Security.Cryptography.Rfc2898DeriveBytes($Password, $Salt, 10000)
    return $PBKDF2.GetBytes(32) # AES-256 Key
}

function Encrypt-Folder {
    param($FolderPath, $Password)
    
    if (-not (Test-Path $FolderPath -PathType Container)) { return }

    $ZipPath = $FolderPath + ".tmp.zip"
    $OutPath = $FolderPath + $VaultExtension
    
    try {
        # 1. Zip the folder
        Compress-Archive -Path "$FolderPath\*" -DestinationPath $ZipPath -Force
        
        # 2. Setup AES
        $Aes = New-Object System.Security.Cryptography.AesManaged
        $Aes.KeySize = 256
        $Aes.BlockSize = 128
        $Aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
        $Aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
        
        $Aes.GenerateIV()
        $Salt = New-Object byte[] 32
        [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($Salt)
        
        $Aes.Key = New-Key -Password $Password -Salt $Salt
        
        $Encryptor = $Aes.CreateEncryptor()
        
        # 3. Write Vault File [Salt(32) + IV(16) + Data]
        $fsOut = New-Object System.IO.FileStream($OutPath, [System.IO.FileMode]::Create)
        $fsOut.Write($Salt, 0, $Salt.Length)
        $fsOut.Write($Aes.IV, 0, $Aes.IV.Length)
        
        $cs = New-Object System.Security.Cryptography.CryptoStream($fsOut, $Encryptor, [System.Security.Cryptography.CryptoStreamMode]::Write)
        $fsIn = New-Object System.IO.FileStream($ZipPath, [System.IO.FileMode]::Open)
        
        $fsIn.CopyTo($cs)
        
        $fsIn.Close()
        $cs.Close()
        $fsOut.Close()
        
        Remove-Item $ZipPath -Force
        
        # 4. Cleanup choice
        $Choice = [System.Windows.MessageBox]::Show("Folder Locked Successfully! Do you want to delete the original folder?", "Security Choice", "YesNo", "Question")
        if ($Choice -eq "Yes") {
            Remove-Item $FolderPath -Recurse -Force
        }
    } catch {
        [System.Windows.MessageBox]::Show("Encryption Failed: $($_.Exception.Message)", "Error", "OK", "Error")
    }
}

function Decrypt-Vault {
    param($VaultPath, $Password)
    
    if (-not (Test-Path $VaultPath)) { return }

    $ParentDir = Split-Path $VaultPath -Parent
    $BaseName = [System.IO.Path]::GetFileNameWithoutExtension($VaultPath)
    $ZipPath = Join-Path $ParentDir ($BaseName + ".tmp.zip")
    $OutDir = Join-Path $ParentDir $BaseName
    
    try {
        $fsIn = New-Object System.IO.FileStream($VaultPath, [System.IO.FileMode]::Open)
        
        $Salt = New-Object byte[] 32
        $fsIn.Read($Salt, 0, 32) | Out-Null
        
        $IV = New-Object byte[] 16
        $fsIn.Read($IV, 0, 16) | Out-Null
        
        $Aes = New-Object System.Security.Cryptography.AesManaged
        $Aes.KeySize = 256
        $Aes.BlockSize = 128
        $Aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
        $Aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
        
        $Aes.IV = $IV
        $Aes.Key = New-Key -Password $Password -Salt $Salt
        
        $Decryptor = $Aes.CreateDecryptor()
        
        $fsOut = New-Object System.IO.FileStream($ZipPath, [System.IO.FileMode]::Create)
        $cs = New-Object System.Security.Cryptography.CryptoStream($fsOut, $Decryptor, [System.Security.Cryptography.CryptoStreamMode]::Write)
        
        $fsIn.CopyTo($cs)
        
        $cs.Close()
        $fsOut.Close()
        $fsIn.Close()
        
        # Extract Zip
        if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
        Expand-Archive -Path $ZipPath -DestinationPath $OutDir -Force
        Remove-Item $ZipPath -Force
        
        [System.Windows.MessageBox]::Show("Vault Unlocked Successfully!", "Success", "OK", "Information")
    } catch {
        if (Test-Path $ZipPath) { Remove-Item $ZipPath -Force }
        [System.Windows.MessageBox]::Show("Decryption Failed! Check your password.", "Error", "OK", "Error")
    }
}

# --- Installation Logic ---
function Install-Integration {
    $ScriptPath = $PSCommandPath
    if (-not $ScriptPath) { $ScriptPath = (Get-Item $MyInvocation.MyCommand.Path).FullName }

    # We use -WindowStyle Normal so the password prompt is definitely visible
    $ActionBase = "powershell.exe -ExecutionPolicy Bypass -File `"$ScriptPath`""
    
    try {
        # 1. Folder Lock
        $FolderKey = "Registry::HKEY_CLASSES_ROOT\Directory\shell\SecureVault.Lock"
        New-Item -Path $FolderKey -Force | Out-Null
        New-ItemProperty -Path $FolderKey -Name "MUIVerb" -Value "🛡️ Lock Folder (Secure Vault)" -Force | Out-Null
        New-ItemProperty -Path $FolderKey -Name "Icon" -Value "shell32.dll,47" -Force | Out-Null
        New-Item -Path "$FolderKey\command" -Value "$ActionBase -Action Lock -Target `"%1`"" -Force | Out-Null

        # 2. Vault Unlock
        $VaultKey = "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\$VaultExtension\shell\SecureVault.Unlock"
        New-Item -Path $VaultKey -Force | Out-Null
        New-ItemProperty -Path $VaultKey -Name "MUIVerb" -Value "🔓 Unlock Vault" -Force | Out-Null
        New-ItemProperty -Path $VaultKey -Name "Icon" -Value "shell32.dll,46" -Force | Out-Null
        New-Item -Path "$VaultKey\command" -Value "$ActionBase -Action Unlock -Target `"%1`"" -Force | Out-Null

        [System.Windows.MessageBox]::Show("Secure Vault Integration Installed Successfully!", "Success")
    } catch {
        [System.Windows.MessageBox]::Show("Installation Failed. Please run as Administrator.", "Error")
    }
}

function Uninstall-Integration {
    try {
        Remove-Item "Registry::HKEY_CLASSES_ROOT\Directory\shell\SecureVault.Lock" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\$VaultExtension\shell\SecureVault.Unlock" -Recurse -Force -ErrorAction SilentlyContinue
        [System.Windows.MessageBox]::Show("Secure Vault Integration Removed.", "Success")
    } catch {
        [System.Windows.MessageBox]::Show("Uninstallation Failed. Please run as Administrator.", "Error")
    }
}

# --- Main Logic Router ---
if ($Action -eq "Lock") {
    $Pass = Get-Password -Title "Locking Folder" -PromptText "Enter Vault Password"
    if ($Pass) { Encrypt-Folder -FolderPath $Target -Password $Pass }
}
elseif ($Action -eq "Unlock") {
    $Pass = Get-Password -Title "Unlocking Vault" -PromptText "Enter Vault Password"
    if ($Pass) { Decrypt-Vault -VaultPath $Target -Password $Pass }
}
elseif ($Action -eq "Uninstall") {
    Uninstall-Integration
}
else {
    # Default GUI for Initial Setup
    try {
        Add-Type -AssemblyName PresentationFramework
    } catch {}
    
    [xml]$MainXaml = @"
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            Title="Secure Vault Manager" Height="250" Width="400" WindowStartupLocation="CenterScreen" Background="#1E1E1E">
        <StackPanel Margin="20" VerticalAlignment="Center">
            <TextBlock Text="🛡️ SECURE VAULT" Foreground="White" FontSize="24" FontWeight="Bold" HorizontalAlignment="Center" Margin="0,0,0,30"/>
            <Button Name="BtnInstall" Content="INSTALL EXPLORER MENU" Height="40" Background="#007ACC" Foreground="White" FontWeight="Bold" Margin="0,5"/>
            <Button Name="BtnUninstall" Content="REMOVE INTEGRATION" Height="40" Background="#444444" Foreground="White" Margin="0,5"/>
        </StackPanel>
    </Window>
"@
    $Reader = (New-Object System.Xml.XmlNodeReader $MainXaml)
    $MainWindow = [Windows.Markup.XamlReader]::Load($Reader)
    
    $MainWindow.FindName("BtnInstall").Add_Click({ Install-Integration })
    $MainWindow.FindName("BtnUninstall").Add_Click({ Uninstall-Integration })
    
    $MainWindow.ShowDialog() | Out-Null
}
