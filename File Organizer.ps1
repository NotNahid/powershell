Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationCore

# ========================== MODERN FOLDER PICKER (COM Shell) ==========================
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

[ComImport, Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")]
public class FileOpenDialogCOM {}

[ComImport, Guid("42F85136-DB7E-439C-85F1-E4075D135FC8"),
 InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IFileOpenDialog {
    void Show(IntPtr parent);
    void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
    void SetFileTypeIndex(uint iFileType);
    void GetFileTypeIndex(out uint piFileType);
    void Advise(IntPtr pfde, out uint pdwCookie);
    void Unadvise(uint dwCookie);
    void SetOptions(uint fos);
    void GetOptions(out uint pfos);
    void SetDefaultFolder(IShellItem psi);
    void SetFolder(IShellItem psi);
    void GetFolder(out IShellItem ppsi);
    void GetCurrentSelection(out IShellItem ppsi);
    void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
    void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
    void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
    void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
    void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
    void GetResult(out IShellItem ppsi);
    void AddPlace(IShellItem psi, int fdap);
    void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
    void Close(int hr);
    void SetClientGuid([In] ref Guid guid);
    void ClearClientData();
    void SetFilter(IntPtr pFilter);
    void GetResults(out IntPtr ppenum);
    void GetSelectedItems(out IntPtr ppsai);
}

[ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"),
 InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IShellItem {
    void BindToHandler(IntPtr pbc, [In] ref Guid bhid, [In] ref Guid riid, out IntPtr ppv);
    void GetParent(out IShellItem ppsi);
    void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
    void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
    void Compare(IShellItem psi, uint hint, out int piOrder);
}

public static class ModernFolderPicker {
    [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    private static extern void SHCreateItemFromParsingName(
        [MarshalAs(UnmanagedType.LPWStr)] string pszPath,
        IntPtr pbc,
        [In, MarshalAs(UnmanagedType.LPStruct)] Guid riid,
        out IShellItem ppv);

    private const uint FOS_PICKFOLDERS = 0x00000020;
    private const uint FOS_FORCEFILESYSTEM = 0x00000040;
    private const uint FOS_NOCHANGEDIR = 0x00000008;

    public static string ShowDialog(IntPtr ownerHandle, string title, string initialPath) {
        IFileOpenDialog dialog = (IFileOpenDialog)new FileOpenDialogCOM();

        uint options;
        dialog.GetOptions(out options);
        dialog.SetOptions(options | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_NOCHANGEDIR);

        if (!string.IsNullOrEmpty(title))
            dialog.SetTitle(title);

        dialog.SetOkButtonLabel("Select Folder");

        if (!string.IsNullOrEmpty(initialPath) && System.IO.Directory.Exists(initialPath)) {
            Guid shellItemGuid = typeof(IShellItem).GUID;
            IShellItem folder;
            SHCreateItemFromParsingName(initialPath, IntPtr.Zero, shellItemGuid, out folder);
            if (folder != null)
                dialog.SetFolder(folder);
        }

        try {
            dialog.Show(ownerHandle);
            IShellItem item;
            dialog.GetResult(out item);
            string path;
            item.GetDisplayName(0x80058000, out path);
            return path;
        }
        catch {
            // Fallback to classic dialog if COM fails
            FolderBrowserDialog fallback = new FolderBrowserDialog();
            fallback.Description = title;
            fallback.ShowNewFolderButton = false;
            if (!string.IsNullOrEmpty(initialPath) && System.IO.Directory.Exists(initialPath))
                fallback.SelectedPath = initialPath;
            if (fallback.ShowDialog() == DialogResult.OK)
                return fallback.SelectedPath;
            return null;
        }
    }
}
"@ -ReferencedAssemblies System.Windows.Forms

# Rest of the script remains exactly the same...

# ========================== CONFIG / PERSISTENCE ==========================
$global:AppDataFolder = Join-Path $env:APPDATA "FileOrganizerPro"
$global:ConfigFile = Join-Path $global:AppDataFolder "config.json"
$global:StatsFile = Join-Path $global:AppDataFolder "stats.json"
$global:LogFolder = Join-Path $global:AppDataFolder "logs"
$global:ErrorLogFile = Join-Path $global:LogFolder "error_log.txt"

if (-not (Test-Path $global:AppDataFolder)) { New-Item -Path $global:AppDataFolder -ItemType Directory -Force | Out-Null }
if (-not (Test-Path $global:LogFolder)) { New-Item -Path $global:LogFolder -ItemType Directory -Force | Out-Null }

# ========================== ERROR LOGGING ==========================
function Write-ErrorLog {
    param([string]$Message, [string]$Context = "General")
    try {
        $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        Add-Content -Path $global:ErrorLogFile -Value "[$timestamp] [$Context] $Message" -Force
    }
    catch {}
}

# ========================== CONFIG FUNCTIONS ==========================
function Load-Config {
    $default = @{
        Theme            = "Dark"
        RecentFolders    = @()
        CustomCategories = @{}
        RecursiveMode    = $false
        DryRunMode       = $false
        ConflictAction   = "Rename"
        WindowWidth      = 1100
        WindowHeight     = 850
        LastFolder       = ""
        MaxPreviewFiles  = 5
        ShowHiddenFiles  = $false
    }
    if (Test-Path $global:ConfigFile) {
        try {
            $loaded = Get-Content $global:ConfigFile -Raw | ConvertFrom-Json
            $config = @{}
            foreach ($prop in $default.Keys) {
                if ($loaded.PSObject.Properties.Name -contains $prop) {
                    $val = $loaded.$prop
                    if ($val -is [System.Management.Automation.PSCustomObject]) {
                        $ht = @{}; $val.PSObject.Properties | ForEach-Object { $ht[$_.Name] = @($_.Value) }
                        $config[$prop] = $ht
                    }
                    elseif ($val -is [System.Object[]]) { $config[$prop] = @($val) }
                    else { $config[$prop] = $val }
                }
                else { $config[$prop] = $default[$prop] }
            }
            return $config
        }
        catch { Write-ErrorLog "Config load failed: $_" "Config"; return $default }
    }
    return $default
}

function Save-Config {
    param([hashtable]$Config)
    try { $Config | ConvertTo-Json -Depth 5 | Set-Content $global:ConfigFile -Force }
    catch { Write-ErrorLog "Config save failed: $_" "Config" }
}

function Load-Stats {
    $default = @{
        TotalFilesOrganized = 0; TotalSizeMoved = 0; TotalOperations = 0
        TotalUndoOperations = 0; CategoryCounts = @{}; ExtensionCounts = @{}
        FirstUsed = (Get-Date).ToString("yyyy-MM-dd HH:mm")
        LastUsed  = (Get-Date).ToString("yyyy-MM-dd HH:mm")
    }
    if (Test-Path $global:StatsFile) {
        try {
            $loaded = Get-Content $global:StatsFile -Raw | ConvertFrom-Json
            $stats = @{}
            foreach ($prop in $default.Keys) {
                if ($loaded.PSObject.Properties.Name -contains $prop) {
                    $val = $loaded.$prop
                    if ($val -is [System.Management.Automation.PSCustomObject]) {
                        $ht = @{}; $val.PSObject.Properties | ForEach-Object { $ht[$_.Name] = $_.Value }
                        $stats[$prop] = $ht
                    }
                    else { $stats[$prop] = $val }
                }
                else { $stats[$prop] = $default[$prop] }
            }
            return $stats
        }
        catch { Write-ErrorLog "Stats load failed: $_" "Stats"; return $default }
    }
    return $default
}

function Save-Stats {
    param([hashtable]$Stats)
    try { $Stats | ConvertTo-Json -Depth 5 | Set-Content $global:StatsFile -Force }
    catch { Write-ErrorLog "Stats save failed: $_" "Stats" }
}

$global:Config = Load-Config
$global:Stats = Load-Stats

# ========================== FILE CATEGORIES ==========================
$global:FileCategories = [ordered]@{
    "Images"      = @(".jpg",".jpeg",".png",".gif",".bmp",".svg",".webp",".ico",".tiff",".tif",".raw",".heic",".heif")
    "Videos"      = @(".mp4",".avi",".mkv",".mov",".wmv",".flv",".webm",".m4v",".mpg",".mpeg",".3gp")
    "Audio"       = @(".mp3",".wav",".flac",".aac",".ogg",".wma",".m4a",".opus",".aiff",".alac")
    "Documents"   = @(".pdf",".doc",".docx",".txt",".rtf",".odt",".xls",".xlsx",".ppt",".pptx",".csv",".epub",".pages",".numbers",".key")
    "Archives"    = @(".zip",".rar",".7z",".tar",".gz",".bz2",".xz",".iso",".cab")
    "Code"        = @(".py",".js",".ts",".html",".htm",".css",".java",".cpp",".c",".cs",".ps1",".sh",".bash",".json",".xml",".yaml",".yml",".sql",".php",".rb",".go",".rs",".swift",".kt",".r",".lua",".pl",".md",".ini",".cfg",".conf",".toml",".log")
    "Executables" = @(".exe",".msi",".bat",".cmd",".com",".scr",".appx",".msix")
    "Fonts"       = @(".ttf",".otf",".woff",".woff2",".eot",".fon")
    "3D Models"   = @(".obj",".fbx",".stl",".blend",".3ds",".dae",".gltf",".glb")
    "Design"      = @(".psd",".ai",".xd",".fig",".sketch",".indd",".cdr")
    "Torrents"    = @(".torrent")
}

# O(1) reverse lookup
$global:ExtensionLookup = @{}
foreach ($category in $global:FileCategories.GetEnumerator()) {
    foreach ($ext in $category.Value) { $global:ExtensionLookup[$ext.ToLower()] = $category.Key }
}

if ($global:Config.CustomCategories -and $global:Config.CustomCategories.Count -gt 0) {
    foreach ($entry in $global:Config.CustomCategories.GetEnumerator()) {
        $global:FileCategories[$entry.Key] = @($entry.Value)
        foreach ($ext in $entry.Value) { $global:ExtensionLookup[$ext.ToLower()] = $entry.Key }
    }
}

# ========================== THEMES ==========================
$global:Themes = @{
    "Dark" = @{
        Bg=[System.Drawing.Color]::FromArgb(22,22,28); Panel=[System.Drawing.Color]::FromArgb(32,32,40)
        Input=[System.Drawing.Color]::FromArgb(42,42,52); Accent=[System.Drawing.Color]::FromArgb(0,122,204)
        Green=[System.Drawing.Color]::FromArgb(46,160,67); Red=[System.Drawing.Color]::FromArgb(200,60,60)
        Orange=[System.Drawing.Color]::FromArgb(220,160,40); TextPrimary=[System.Drawing.Color]::White
        TextDim=[System.Drawing.Color]::FromArgb(160,160,170); TextMuted=[System.Drawing.Color]::FromArgb(110,110,120)
        Border=[System.Drawing.Color]::FromArgb(55,55,65); ListBg=[System.Drawing.Color]::FromArgb(28,28,35)
        FolderText=[System.Drawing.Color]::FromArgb(100,180,255); DriveText=[System.Drawing.Color]::FromArgb(255,200,80)
        QuickText=[System.Drawing.Color]::FromArgb(120,220,150); Success=[System.Drawing.Color]::FromArgb(0,200,120)
        HighlightBg=[System.Drawing.Color]::FromArgb(45,45,58); HintBg=[System.Drawing.Color]::FromArgb(36,36,46)
        DryRun=[System.Drawing.Color]::FromArgb(180,120,255)
    }
    "Light" = @{
        Bg=[System.Drawing.Color]::FromArgb(242,242,247); Panel=[System.Drawing.Color]::FromArgb(255,255,255)
        Input=[System.Drawing.Color]::FromArgb(245,245,250); Accent=[System.Drawing.Color]::FromArgb(0,102,184)
        Green=[System.Drawing.Color]::FromArgb(36,140,57); Red=[System.Drawing.Color]::FromArgb(190,50,50)
        Orange=[System.Drawing.Color]::FromArgb(200,140,20); TextPrimary=[System.Drawing.Color]::FromArgb(30,30,35)
        TextDim=[System.Drawing.Color]::FromArgb(100,100,110); TextMuted=[System.Drawing.Color]::FromArgb(140,140,150)
        Border=[System.Drawing.Color]::FromArgb(200,200,212); ListBg=[System.Drawing.Color]::FromArgb(250,250,253)
        FolderText=[System.Drawing.Color]::FromArgb(0,90,180); DriveText=[System.Drawing.Color]::FromArgb(180,130,0)
        QuickText=[System.Drawing.Color]::FromArgb(30,150,80); Success=[System.Drawing.Color]::FromArgb(0,160,90)
        HighlightBg=[System.Drawing.Color]::FromArgb(228,233,245); HintBg=[System.Drawing.Color]::FromArgb(238,240,248)
        DryRun=[System.Drawing.Color]::FromArgb(130,80,200)
    }
}

$global:CurrentTheme = $global:Config.Theme
function Get-Theme { return $global:Themes[$global:CurrentTheme] }

# ========================== HELPERS ==========================
function Get-CategoryForFile {
    param([string]$Extension)
    $ext = $Extension.ToLower()
    if ($global:ExtensionLookup.ContainsKey($ext)) { return $global:ExtensionLookup[$ext] }
    return "Other"
}

function Get-UniqueFileName {
    param([string]$DestinationFolder, [string]$FileName)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $extension = [System.IO.Path]::GetExtension($FileName)
    $finalPath = Join-Path $DestinationFolder $FileName
    $counter = 1
    while (Test-Path $finalPath) {
        $finalPath = Join-Path $DestinationFolder "${baseName} ($counter)${extension}"
        $counter++
    }
    return $finalPath
}

function Get-FolderSizeText {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    elseif ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    elseif ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    else { return "$Bytes B" }
}

function Add-RecentFolder {
    param([string]$Path)
    $recents = [System.Collections.ArrayList]@()
    if ($global:Config.RecentFolders) { foreach ($r in $global:Config.RecentFolders) { [void]$recents.Add($r) } }
    if ($recents.Contains($Path)) { $recents.Remove($Path) }
    $recents.Insert(0, $Path)
    if ($recents.Count -gt 15) { $recents = [System.Collections.ArrayList]@($recents[0..14]) }
    $global:Config.RecentFolders = @($recents)
    Save-Config -Config $global:Config
}

function Export-MoveLog {
    param([array]$Log, [string]$FolderPath)
    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $logFile = Join-Path $global:LogFolder "organize_log_$timestamp.csv"
    $Log | Select-Object FileName,Category,OriginalPath,DestinationPath,FileSize,Timestamp |
        Export-Csv -Path $logFile -NoTypeInformation -Force
    return $logFile
}

function Select-FolderModern {
    param([string]$Title = "Select a folder to organize", [string]$InitialPath = "")
    try {
        $path = [ModernFolderPicker]::ShowDialog($form.Handle, $Title, $InitialPath)
        return $path
    }
    catch {
        # Fallback to classic dialog if COM fails
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = $Title
        $dialog.ShowNewFolderButton = $false
        if ($InitialPath -ne "" -and (Test-Path $InitialPath)) { $dialog.SelectedPath = $InitialPath }
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dialog.SelectedPath }
        return $null
    }
}

# ========================== GLOBAL STATE ==========================
$global:SelectedPath = ""
$global:AnalysisResults = @{}
$global:AnalysisCache = @{}
$global:TotalFiles = 0
$global:MoveLog = @()
$global:AllFiles = @()
$global:DryRunMode = [bool]$global:Config.DryRunMode
$global:RecursiveMode = [bool]$global:Config.RecursiveMode
$global:ConflictAction = $global:Config.ConflictAction
$global:IsOrganizing = $false

# ========================== TOOLTIP ==========================
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 12000; $toolTip.InitialDelay = 300
$toolTip.ReshowDelay = 100; $toolTip.ShowAlways = $true

# ========================== MAIN FORM ==========================
$form = New-Object System.Windows.Forms.Form
$form.Text = "File Organizer Pro v2.0"
$form.Size = New-Object System.Drawing.Size($global:Config.WindowWidth, $global:Config.WindowHeight)
$form.MinimumSize = New-Object System.Drawing.Size(850, 650)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.KeyPreview = $true; $form.AllowDrop = $true

$form.Add_FormClosing({
    $global:Config.WindowWidth = $form.Width; $global:Config.WindowHeight = $form.Height
    $global:Config.Theme = $global:CurrentTheme; $global:Config.DryRunMode = $global:DryRunMode
    $global:Config.RecursiveMode = $global:RecursiveMode; $global:Config.ConflictAction = $global:ConflictAction
    $global:Config.LastFolder = $global:SelectedPath; $global:Config.ShowHiddenFiles = $chkShowHidden.Checked
    Save-Config -Config $global:Config
    $global:Stats.LastUsed = (Get-Date).ToString("yyyy-MM-dd HH:mm")
    Save-Stats -Stats $global:Stats
})

# Drag and Drop
$form.Add_DragEnter({
    if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $paths = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        if ($paths.Count -gt 0 -and (Test-Path $paths[0] -PathType Container)) {
            $_.Effect = [System.Windows.Forms.DragDropEffects]::Link
        } else { $_.Effect = [System.Windows.Forms.DragDropEffects]::None }
    }
})
$form.Add_DragDrop({
    $paths = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    if ($paths.Count -gt 0 -and (Test-Path $paths[0] -PathType Container)) { Set-SelectedFolder -Path $paths[0] }
})

# ========================== MAIN LAYOUT ==========================
$mainPanel = New-Object System.Windows.Forms.TableLayoutPanel
$mainPanel.Dock = "Fill"; $mainPanel.Padding = New-Object System.Windows.Forms.Padding(14)
$mainPanel.RowCount = 8; $mainPanel.ColumnCount = 1
$mainPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 100))) | Out-Null
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 44))) | Out-Null   # 0: Header
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 70))) | Out-Null   # 1: Folder selection
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 34))) | Out-Null   # 2: Mode bar
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 48))) | Out-Null   # 3: Filters
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 50))) | Out-Null    # 4: Preview
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 28))) | Out-Null   # 5: Progress
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 50))) | Out-Null   # 6: Buttons
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 50))) | Out-Null    # 7: Bottom
$form.Controls.Add($mainPanel)

# ========================== ROW 0: HEADER ==========================
$headerPanel = New-Object System.Windows.Forms.TableLayoutPanel
$headerPanel.Dock = "Fill"; $headerPanel.ColumnCount = 6; $headerPanel.RowCount = 1
$headerPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 100))) | Out-Null
$headerPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 100))) | Out-Null
$headerPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 90))) | Out-Null
$headerPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 100))) | Out-Null
$headerPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 100))) | Out-Null
$headerPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 100))) | Out-Null
$headerPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 180))) | Out-Null

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "FILE ORGANIZER PRO"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
$titleLabel.Dock = "Fill"; $titleLabel.TextAlign = "MiddleLeft"
$headerPanel.Controls.Add($titleLabel, 0, 0)

$btnStats = New-Object System.Windows.Forms.Button
$btnStats.Text = "Stats"; $btnStats.Dock = "Fill"; $btnStats.FlatStyle = "Flat"
$btnStats.Font = New-Object System.Drawing.Font("Segoe UI", 8.5); $btnStats.Cursor = "Hand"
$btnStats.Margin = New-Object System.Windows.Forms.Padding(2,7,2,7)
$toolTip.SetToolTip($btnStats, "Lifetime statistics (Ctrl+I)")
$headerPanel.Controls.Add($btnStats, 1, 0)

$btnCustomCat = New-Object System.Windows.Forms.Button
$btnCustomCat.Text = "Categories"; $btnCustomCat.Dock = "Fill"; $btnCustomCat.FlatStyle = "Flat"
$btnCustomCat.Font = New-Object System.Drawing.Font("Segoe UI", 8.5); $btnCustomCat.Cursor = "Hand"
$btnCustomCat.Margin = New-Object System.Windows.Forms.Padding(2,7,2,7)
$toolTip.SetToolTip($btnCustomCat, "Custom categories (Ctrl+K)")
$headerPanel.Controls.Add($btnCustomCat, 2, 0)

$btnRecent = New-Object System.Windows.Forms.Button
$btnRecent.Text = "Recent"; $btnRecent.Dock = "Fill"; $btnRecent.FlatStyle = "Flat"
$btnRecent.Font = New-Object System.Drawing.Font("Segoe UI", 8.5); $btnRecent.Cursor = "Hand"
$btnRecent.Margin = New-Object System.Windows.Forms.Padding(2,7,2,7)
$toolTip.SetToolTip($btnRecent, "Recent folders (Ctrl+E)")
$headerPanel.Controls.Add($btnRecent, 3, 0)

$themeToggleBtn = New-Object System.Windows.Forms.Button
$themeToggleBtn.Text = "Light Mode"; $themeToggleBtn.Dock = "Fill"; $themeToggleBtn.FlatStyle = "Flat"
$themeToggleBtn.Font = New-Object System.Drawing.Font("Segoe UI", 8.5); $themeToggleBtn.Cursor = "Hand"
$themeToggleBtn.Margin = New-Object System.Windows.Forms.Padding(2,7,2,7)
$toolTip.SetToolTip($themeToggleBtn, "Toggle theme (Ctrl+T)")
$headerPanel.Controls.Add($themeToggleBtn, 4, 0)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Ready"; $statusLabel.Dock = "Fill"; $statusLabel.TextAlign = "MiddleRight"
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$headerPanel.Controls.Add($statusLabel, 5, 0)

$mainPanel.Controls.Add($headerPanel, 0, 0)

# ========================== ROW 1: FOLDER SELECTION ==========================
$folderPanel = New-Object System.Windows.Forms.TableLayoutPanel
$folderPanel.Dock = "Fill"; $folderPanel.ColumnCount = 4; $folderPanel.RowCount = 2
$folderPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 30))) | Out-Null
$folderPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 100))) | Out-Null
$folderPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 150))) | Out-Null
$folderPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 100))) | Out-Null
$folderPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 70))) | Out-Null
$folderPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 75))) | Out-Null

$selectFolderBtn = New-Object System.Windows.Forms.Button
$selectFolderBtn.Text = "  Browse Folder..."
$selectFolderBtn.Dock = "Fill"; $selectFolderBtn.FlatStyle = "Flat"
$selectFolderBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$selectFolderBtn.Cursor = "Hand"; $selectFolderBtn.TextAlign = "MiddleCenter"
$selectFolderBtn.Margin = New-Object System.Windows.Forms.Padding(0,2,6,2)
$toolTip.SetToolTip($selectFolderBtn, "Open full Windows Explorer folder picker (Ctrl+B)")
$folderPanel.Controls.Add($selectFolderBtn, 0, 0)
$folderPanel.SetRowSpan($selectFolderBtn, 2)

$folderPathBox = New-Object System.Windows.Forms.TextBox
$folderPathBox.Dock = "Fill"; $folderPathBox.BorderStyle = "FixedSingle"
$folderPathBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$folderPathBox.Margin = New-Object System.Windows.Forms.Padding(0,4,4,0)
$toolTip.SetToolTip($folderPathBox, "Type or paste a path, press Enter (Ctrl+L to focus)")
$folderPanel.Controls.Add($folderPathBox, 1, 0)

$btnGo = New-Object System.Windows.Forms.Button
$btnGo.Text = "Go"; $btnGo.Dock = "Fill"; $btnGo.FlatStyle = "Flat"
$btnGo.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnGo.Cursor = "Hand"; $btnGo.Margin = New-Object System.Windows.Forms.Padding(0,4,2,0)
$folderPanel.Controls.Add($btnGo, 2, 0)

$btnOpenExplorer = New-Object System.Windows.Forms.Button
$btnOpenExplorer.Text = "Open"; $btnOpenExplorer.Dock = "Fill"; $btnOpenExplorer.FlatStyle = "Flat"
$btnOpenExplorer.Font = New-Object System.Drawing.Font("Segoe UI", 9); $btnOpenExplorer.Cursor = "Hand"
$btnOpenExplorer.Margin = New-Object System.Windows.Forms.Padding(2,4,0,0)
$toolTip.SetToolTip($btnOpenExplorer, "Open in Explorer (Ctrl+O)")
$folderPanel.Controls.Add($btnOpenExplorer, 3, 0)

$folderInfoLabel = New-Object System.Windows.Forms.Label
$folderInfoLabel.Text = "No folder selected  --  Click 'Browse Folder' or drag and drop a folder here"
$folderInfoLabel.Dock = "Fill"; $folderInfoLabel.TextAlign = "MiddleLeft"
$folderInfoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$folderPanel.Controls.Add($folderInfoLabel, 1, 1)
$folderPanel.SetColumnSpan($folderInfoLabel, 3)

$mainPanel.Controls.Add($folderPanel, 0, 1)

# ========================== ROW 2: MODE BAR ==========================
$modePanel = New-Object System.Windows.Forms.TableLayoutPanel
$modePanel.Dock = "Fill"; $modePanel.ColumnCount = 7; $modePanel.RowCount = 1
$modePanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 100))) | Out-Null
$modePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 105))) | Out-Null
$modePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 115))) | Out-Null
$modePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 130))) | Out-Null
$modePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 16))) | Out-Null
$modePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 90))) | Out-Null
$modePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 115))) | Out-Null
$modePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 100))) | Out-Null

$chkDryRun = New-Object System.Windows.Forms.CheckBox
$chkDryRun.Text = "Dry Run"; $chkDryRun.Checked = $global:DryRunMode
$chkDryRun.AutoSize = $true; $chkDryRun.Dock = "Fill"
$chkDryRun.Font = New-Object System.Drawing.Font("Segoe UI", 9); $chkDryRun.Cursor = "Hand"
$toolTip.SetToolTip($chkDryRun, "Simulate without moving (Ctrl+D)")
$modePanel.Controls.Add($chkDryRun, 0, 0)

$chkRecursive = New-Object System.Windows.Forms.CheckBox
$chkRecursive.Text = "Recursive"; $chkRecursive.Checked = $global:RecursiveMode
$chkRecursive.AutoSize = $true; $chkRecursive.Dock = "Fill"
$chkRecursive.Font = New-Object System.Drawing.Font("Segoe UI", 9); $chkRecursive.Cursor = "Hand"
$toolTip.SetToolTip($chkRecursive, "Include subfolders (Ctrl+R)")
$modePanel.Controls.Add($chkRecursive, 1, 0)

$chkShowHidden = New-Object System.Windows.Forms.CheckBox
$chkShowHidden.Text = "Hidden Files"; $chkShowHidden.Checked = [bool]$global:Config.ShowHiddenFiles
$chkShowHidden.AutoSize = $true; $chkShowHidden.Dock = "Fill"
$chkShowHidden.Font = New-Object System.Drawing.Font("Segoe UI", 9); $chkShowHidden.Cursor = "Hand"
$toolTip.SetToolTip($chkShowHidden, "Include hidden/system files")
$modePanel.Controls.Add($chkShowHidden, 2, 0)

$sepLabel = New-Object System.Windows.Forms.Label
$sepLabel.Text = "|"; $sepLabel.Dock = "Fill"; $sepLabel.TextAlign = "MiddleCenter"
$modePanel.Controls.Add($sepLabel, 3, 0)

$conflictLabel = New-Object System.Windows.Forms.Label
$conflictLabel.Text = "Conflict:"; $conflictLabel.Dock = "Fill"; $conflictLabel.TextAlign = "MiddleLeft"
$conflictLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$modePanel.Controls.Add($conflictLabel, 4, 0)

$conflictCombo = New-Object System.Windows.Forms.ComboBox
$conflictCombo.Dock = "Fill"; $conflictCombo.DropDownStyle = "DropDownList"
$conflictCombo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$conflictCombo.Items.AddRange(@("Rename","Skip","Overwrite"))
$conflictCombo.SelectedItem = $global:ConflictAction
$conflictCombo.Margin = New-Object System.Windows.Forms.Padding(0,4,4,4)
$modePanel.Controls.Add($conflictCombo, 5, 0)

$modeStatusLabel = New-Object System.Windows.Forms.Label
$modeStatusLabel.Text = ""; $modeStatusLabel.Dock = "Fill"; $modeStatusLabel.TextAlign = "MiddleRight"
$modeStatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$modePanel.Controls.Add($modeStatusLabel, 6, 0)

$mainPanel.Controls.Add($modePanel, 0, 2)

$chkDryRun.Add_CheckedChanged({ $global:DryRunMode = $chkDryRun.Checked; Update-ModeStatus; if ($global:SelectedPath -ne "") { Invoke-FolderAnalysis -Path $global:SelectedPath -UseCache } })
$chkRecursive.Add_CheckedChanged({ $global:RecursiveMode = $chkRecursive.Checked; Update-ModeStatus; $global:AnalysisCache = @{}; if ($global:SelectedPath -ne "") { Invoke-FolderAnalysis -Path $global:SelectedPath } })
$chkShowHidden.Add_CheckedChanged({ $global:AnalysisCache = @{}; if ($global:SelectedPath -ne "") { Invoke-FolderAnalysis -Path $global:SelectedPath } })
$conflictCombo.Add_SelectedIndexChanged({ $global:ConflictAction = $conflictCombo.SelectedItem })

function Update-ModeStatus {
    $t = Get-Theme; $parts = @()
    if ($global:DryRunMode) { $parts += "DRY RUN" }
    if ($global:RecursiveMode) { $parts += "RECURSIVE" }
    if ($parts.Count -eq 0) { $modeStatusLabel.Text = ""; return }
    $modeStatusLabel.Text = "[ $($parts -join ' | ') ]"
    $modeStatusLabel.ForeColor = if ($global:DryRunMode) { $t.DryRun } else { $t.Orange }
}

# ========================== ROW 3: CATEGORY FILTERS ==========================
$filterPanel = New-Object System.Windows.Forms.Panel
$filterPanel.Dock = "Fill"; $filterPanel.Padding = New-Object System.Windows.Forms.Padding(0,2,0,2)

$filterFlow = New-Object System.Windows.Forms.FlowLayoutPanel
$filterFlow.Dock = "Fill"; $filterFlow.WrapContents = $true; $filterFlow.AutoScroll = $true
$filterFlow.Padding = New-Object System.Windows.Forms.Padding(4,2,4,2)

$chkAll = New-Object System.Windows.Forms.CheckBox
$chkAll.Text = "All"; $chkAll.Checked = $true; $chkAll.AutoSize = $true
$chkAll.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$chkAll.Margin = New-Object System.Windows.Forms.Padding(4,4,10,4); $chkAll.Cursor = "Hand"
$toolTip.SetToolTip($chkAll, "Toggle all (Ctrl+A)")
$filterFlow.Controls.Add($chkAll)

$global:CategoryCheckboxes = @{}
foreach ($cat in $global:FileCategories.Keys) {
    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Text = $cat; $chk.Checked = $true; $chk.AutoSize = $true
    $chk.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $chk.Margin = New-Object System.Windows.Forms.Padding(3,4,3,4); $chk.Cursor = "Hand"
    $extList = ($global:FileCategories[$cat] -join "  ")
    $toolTip.SetToolTip($chk, "$cat`n$extList")
    $filterFlow.Controls.Add($chk)
    $global:CategoryCheckboxes[$cat] = $chk
}

$chkOther = New-Object System.Windows.Forms.CheckBox
$chkOther.Text = "Other"; $chkOther.Checked = $true; $chkOther.AutoSize = $true
$chkOther.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$chkOther.Margin = New-Object System.Windows.Forms.Padding(3,4,3,4); $chkOther.Cursor = "Hand"
$toolTip.SetToolTip($chkOther, "Uncategorized files")
$filterFlow.Controls.Add($chkOther)
$global:CategoryCheckboxes["Other"] = $chkOther
$filterPanel.Controls.Add($filterFlow)
$mainPanel.Controls.Add($filterPanel, 0, 3)

$global:UpdatingCheckboxes = $false
$chkAll.Add_CheckedChanged({
    if ($global:UpdatingCheckboxes) { return }; $global:UpdatingCheckboxes = $true
    foreach ($chk in $global:CategoryCheckboxes.Values) { $chk.Checked = $chkAll.Checked }
    $global:UpdatingCheckboxes = $false
    if ($global:SelectedPath -ne "") { Invoke-FolderAnalysis -Path $global:SelectedPath -UseCache }
})
foreach ($chk in $global:CategoryCheckboxes.Values) {
    $chk.Add_CheckedChanged({
        if ($global:UpdatingCheckboxes) { return }; $global:UpdatingCheckboxes = $true
        $allOn = $true; foreach ($c in $global:CategoryCheckboxes.Values) { if (-not $c.Checked) { $allOn = $false; break } }
        $chkAll.Checked = $allOn; $global:UpdatingCheckboxes = $false
        if ($global:SelectedPath -ne "") { Invoke-FolderAnalysis -Path $global:SelectedPath -UseCache }
    })
}

# ========================== ROW 4: ANALYSIS PREVIEW ==========================
$previewOuterPanel = New-Object System.Windows.Forms.Panel
$previewOuterPanel.Dock = "Fill"; $previewOuterPanel.Padding = New-Object System.Windows.Forms.Padding(0,2,0,2)
$previewGroupBox = New-Object System.Windows.Forms.GroupBox
$previewGroupBox.Text = "  Analysis Preview  "; $previewGroupBox.Dock = "Fill"
$previewGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$previewOuterPanel.Controls.Add($previewGroupBox)

$previewBox = New-Object System.Windows.Forms.RichTextBox
$previewBox.Dock = "Fill"; $previewBox.ReadOnly = $true; $previewBox.BorderStyle = "None"
$previewBox.Font = New-Object System.Drawing.Font("Cascadia Code, Consolas, Courier New", 9)
$previewBox.Text = "Select a folder to see the analysis..."
$previewGroupBox.Controls.Add($previewBox)
$mainPanel.Controls.Add($previewOuterPanel, 0, 4)

# ========================== ROW 5: PROGRESS ==========================
$progressPanel = New-Object System.Windows.Forms.TableLayoutPanel
$progressPanel.Dock = "Fill"; $progressPanel.ColumnCount = 2; $progressPanel.RowCount = 1
$progressPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 100))) | Out-Null
$progressPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 130))) | Out-Null

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Dock = "Fill"; $progressBar.Style = "Continuous"; $progressBar.Minimum = 0; $progressBar.Value = 0
$progressBar.Margin = New-Object System.Windows.Forms.Padding(0,4,4,4)
$progressPanel.Controls.Add($progressBar, 0, 0)

$progressLabel = New-Object System.Windows.Forms.Label
$progressLabel.Text = ""; $progressLabel.Dock = "Fill"; $progressLabel.TextAlign = "MiddleRight"
$progressLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$progressPanel.Controls.Add($progressLabel, 1, 0)
$mainPanel.Controls.Add($progressPanel, 0, 5)

# ========================== ROW 6: BUTTONS ==========================
$buttonPanel = New-Object System.Windows.Forms.TableLayoutPanel
$buttonPanel.Dock = "Fill"; $buttonPanel.ColumnCount = 5; $buttonPanel.RowCount = 1
$buttonPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 100))) | Out-Null
$buttonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 32))) | Out-Null
$buttonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 22))) | Out-Null
$buttonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 18))) | Out-Null
$buttonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 16))) | Out-Null
$buttonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 12))) | Out-Null

$organizeBtn = New-Object System.Windows.Forms.Button
$organizeBtn.Text = "ORGANIZE FILES"; $organizeBtn.Dock = "Fill"; $organizeBtn.FlatStyle = "Flat"
$organizeBtn.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$organizeBtn.FlatAppearance.BorderSize = 0; $organizeBtn.Cursor = "Hand"; $organizeBtn.Enabled = $false
$organizeBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,3,0)
$toolTip.SetToolTip($organizeBtn, "Organize files (Ctrl+Enter)")
$buttonPanel.Controls.Add($organizeBtn, 0, 0)

$undoLastBtn = New-Object System.Windows.Forms.Button
$undoLastBtn.Text = "Undo Category"; $undoLastBtn.Dock = "Fill"; $undoLastBtn.FlatStyle = "Flat"
$undoLastBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$undoLastBtn.FlatAppearance.BorderSize = 0; $undoLastBtn.Cursor = "Hand"; $undoLastBtn.Enabled = $false
$undoLastBtn.Margin = New-Object System.Windows.Forms.Padding(3,0,3,0)
$toolTip.SetToolTip($undoLastBtn, "Undo category (Ctrl+U)")
$buttonPanel.Controls.Add($undoLastBtn, 1, 0)

$undoAllBtn = New-Object System.Windows.Forms.Button
$undoAllBtn.Text = "Undo All"; $undoAllBtn.Dock = "Fill"; $undoAllBtn.FlatStyle = "Flat"
$undoAllBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$undoAllBtn.FlatAppearance.BorderSize = 0; $undoAllBtn.Cursor = "Hand"; $undoAllBtn.Enabled = $false
$undoAllBtn.Margin = New-Object System.Windows.Forms.Padding(3,0,3,0)
$toolTip.SetToolTip($undoAllBtn, "Undo all (Ctrl+Z)")
$buttonPanel.Controls.Add($undoAllBtn, 2, 0)

$exportLogBtn = New-Object System.Windows.Forms.Button
$exportLogBtn.Text = "Export Log"; $exportLogBtn.Dock = "Fill"; $exportLogBtn.FlatStyle = "Flat"
$exportLogBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9); $exportLogBtn.Cursor = "Hand"
$exportLogBtn.Enabled = $false; $exportLogBtn.Margin = New-Object System.Windows.Forms.Padding(3,0,3,0)
$toolTip.SetToolTip($exportLogBtn, "Export CSV log (Ctrl+S)")
$buttonPanel.Controls.Add($exportLogBtn, 3, 0)

$refreshBtn = New-Object System.Windows.Forms.Button
$refreshBtn.Text = "Refresh"; $refreshBtn.Dock = "Fill"; $refreshBtn.FlatStyle = "Flat"
$refreshBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9); $refreshBtn.Cursor = "Hand"
$refreshBtn.Margin = New-Object System.Windows.Forms.Padding(3,0,0,0)
$toolTip.SetToolTip($refreshBtn, "Re-analyze (F5)")
$buttonPanel.Controls.Add($refreshBtn, 4, 0)
$mainPanel.Controls.Add($buttonPanel, 0, 6)

# ========================== ROW 7: BOTTOM SPLIT ==========================
$bottomSplit = New-Object System.Windows.Forms.TableLayoutPanel
$bottomSplit.Dock = "Fill"; $bottomSplit.ColumnCount = 2; $bottomSplit.RowCount = 1
$bottomSplit.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 100))) | Out-Null
$bottomSplit.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 50))) | Out-Null
$bottomSplit.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 50))) | Out-Null

$hintsGroupBox = New-Object System.Windows.Forms.GroupBox
$hintsGroupBox.Text = "  Shortcuts & Tips  "; $hintsGroupBox.Dock = "Fill"
$hintsGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$hintsBox = New-Object System.Windows.Forms.RichTextBox
$hintsBox.Dock = "Fill"; $hintsBox.ReadOnly = $true; $hintsBox.BorderStyle = "None"
$hintsBox.Font = New-Object System.Drawing.Font("Cascadia Code, Consolas, Courier New", 8)
$hintsBox.ScrollBars = "Vertical"
$hintsGroupBox.Controls.Add($hintsBox)
$bottomSplit.Controls.Add($hintsGroupBox, 0, 0)

$filePreviewGroupBox = New-Object System.Windows.Forms.GroupBox
$filePreviewGroupBox.Text = "  Folder Info  "; $filePreviewGroupBox.Dock = "Fill"
$filePreviewGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$filePreviewPanel = New-Object System.Windows.Forms.Panel; $filePreviewPanel.Dock = "Fill"

$filePreviewPicture = New-Object System.Windows.Forms.PictureBox
$filePreviewPicture.Dock = "Fill"; $filePreviewPicture.SizeMode = "Zoom"; $filePreviewPicture.Visible = $false
$filePreviewPanel.Controls.Add($filePreviewPicture)

$filePreviewText = New-Object System.Windows.Forms.RichTextBox
$filePreviewText.Dock = "Fill"; $filePreviewText.ReadOnly = $true; $filePreviewText.BorderStyle = "None"
$filePreviewText.Font = New-Object System.Drawing.Font("Cascadia Code, Consolas, Courier New", 8.5)
$filePreviewText.ScrollBars = "Vertical"; $filePreviewText.WordWrap = $true; $filePreviewText.Visible = $true
$filePreviewText.Text = "Folder details will appear here after selecting a folder."
$filePreviewPanel.Controls.Add($filePreviewText)
$filePreviewGroupBox.Controls.Add($filePreviewPanel)
$bottomSplit.Controls.Add($filePreviewGroupBox, 1, 0)
$mainPanel.Controls.Add($bottomSplit, 0, 7)

# ========================== FOLDER INFO PREVIEW ==========================
$global:ImageExtensions = @(".jpg",".jpeg",".png",".gif",".bmp",".ico",".tiff",".tif",".webp")
$global:TextExtensions = @(".txt",".md",".log",".csv",".json",".xml",".yaml",".yml",".ini",".cfg",".conf",".toml",
    ".py",".js",".ts",".html",".htm",".css",".java",".cpp",".c",".cs",".ps1",".sh",".bash",
    ".sql",".php",".rb",".go",".rs",".swift",".kt",".r",".lua",".pl",".bat",".cmd")

function Show-FolderPreview {
    param([string]$FolderPath)
    $t = Get-Theme
    $filePreviewPicture.Visible = $false; $filePreviewText.Visible = $true
    $filePreviewText.SuspendLayout(); $filePreviewText.Clear()

    if (-not $FolderPath -or -not (Test-Path $FolderPath)) {
        $filePreviewText.SelectionColor = $t.TextDim
        $filePreviewText.AppendText("Select a folder to see details.")
        $filePreviewText.ResumeLayout(); return
    }

    try {
        $info = Get-Item $FolderPath
        $allItems = Get-ChildItem $FolderPath -ErrorAction SilentlyContinue
        $files = $allItems | Where-Object { -not $_.PSIsContainer }
        $folders = $allItems | Where-Object { $_.PSIsContainer }
        $fileCount = if ($files) { @($files).Count } else { 0 }
        $folderCount = if ($folders) { @($folders).Count } else { 0 }
        $totalSize = if ($files) { ($files | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum } else { 0 }
        if (-not $totalSize) { $totalSize = 0 }

        $filePreviewText.SelectionColor = $t.Accent
        $filePreviewText.AppendText("FOLDER DETAILS`n")
        $filePreviewText.SelectionColor = $t.TextMuted
        $filePreviewText.AppendText(([string]::new([char]0x2500, 38)) + "`n`n")

        $filePreviewText.SelectionColor = $t.TextPrimary
        $filePreviewText.AppendText("  Name:       $($info.Name)`n")
        $filePreviewText.AppendText("  Path:       $($info.FullName)`n")
        $filePreviewText.AppendText("  Files:      $fileCount`n")
        $filePreviewText.AppendText("  Folders:    $folderCount`n")
        $filePreviewText.AppendText("  Size:       $(Get-FolderSizeText $totalSize)`n")
        $filePreviewText.AppendText("  Created:    $($info.CreationTime.ToString('yyyy-MM-dd HH:mm'))`n")
        $filePreviewText.AppendText("  Modified:   $($info.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))`n`n")

        # Extension breakdown
        if ($fileCount -gt 0) {
            $filePreviewText.SelectionColor = $t.Accent
            $filePreviewText.AppendText("FILE TYPES`n")
            $filePreviewText.SelectionColor = $t.TextMuted
            $filePreviewText.AppendText(([string]::new([char]0x2500, 38)) + "`n`n")

            $extGroups = $files | Group-Object { $_.Extension.ToLower() } | Sort-Object Count -Descending | Select-Object -First 12
            $maxCount = ($extGroups | Select-Object -First 1).Count
            foreach ($eg in $extGroups) {
                $extName = if ($eg.Name -eq "") { "(none)" } else { $eg.Name }
                $barLen = if ($maxCount -gt 0) { [math]::Max(1, [math]::Floor(($eg.Count / $maxCount) * 12)) } else { 1 }
                $bar = [string]::new([char]0x2588, $barLen)
                $extSize = ($eg.Group | Measure-Object -Property Length -Sum).Sum

                $filePreviewText.SelectionColor = $t.Success
                $filePreviewText.AppendText("  $($extName.PadRight(8))")
                $filePreviewText.SelectionColor = $t.FolderText
                $filePreviewText.AppendText("$bar ")
                $filePreviewText.SelectionColor = $t.TextDim
                $filePreviewText.AppendText("$($eg.Count) ($(Get-FolderSizeText $extSize))`n")
            }

            if (($files | Group-Object { $_.Extension.ToLower() }).Count -gt 12) {
                $filePreviewText.SelectionColor = $t.TextMuted
                $filePreviewText.AppendText("`n  ... and more types`n")
            }

            # Largest files
            $filePreviewText.AppendText("`n")
            $filePreviewText.SelectionColor = $t.Accent
            $filePreviewText.AppendText("LARGEST FILES`n")
            $filePreviewText.SelectionColor = $t.TextMuted
            $filePreviewText.AppendText(([string]::new([char]0x2500, 38)) + "`n`n")

            $largest = $files | Sort-Object Length -Descending | Select-Object -First 5
            foreach ($lf in $largest) {
                $fname = $lf.Name
                if ($fname.Length -gt 28) { $fname = $fname.Substring(0, 25) + "..." }
                $filePreviewText.SelectionColor = $t.TextPrimary
                $filePreviewText.AppendText("  $($fname.PadRight(30))")
                $filePreviewText.SelectionColor = $t.TextDim
                $filePreviewText.AppendText("$(Get-FolderSizeText $lf.Length)`n")
            }
        }
    }
    catch {
        $filePreviewText.SelectionColor = $t.Red
        $filePreviewText.AppendText("Cannot read folder: $_")
    }
    $filePreviewText.SelectionStart = 0; $filePreviewText.ScrollToCaret()
    $filePreviewText.ResumeLayout()
}

# ========================== FOLDER SELECTION & ANALYSIS ==========================
function Set-SelectedFolder {
    param([string]$Path)
    $t = Get-Theme
    if (-not (Test-Path $Path -PathType Container)) {
        $statusLabel.Text = "Invalid folder"; $statusLabel.ForeColor = $t.Red; return
    }
    $global:SelectedPath = $Path; $global:AnalysisCache = @{}
    $folderPathBox.Text = $Path; $global:Config.LastFolder = $Path
    Save-Config -Config $global:Config

    try {
        $items = Get-ChildItem -Path $Path -ErrorAction SilentlyContinue
        $fileCount = ($items | Where-Object { -not $_.PSIsContainer }).Count
        $folderCount = ($items | Where-Object { $_.PSIsContainer }).Count
        $totalSize = ($items | Where-Object { -not $_.PSIsContainer } | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
        if (-not $totalSize) { $totalSize = 0 }
        $folderInfoLabel.Text = "$fileCount files  |  $folderCount folders  |  $(Get-FolderSizeText $totalSize)  |  $Path"
        $folderInfoLabel.ForeColor = $t.TextDim
    }
    catch { $folderInfoLabel.Text = "Access restricted: $Path"; $folderInfoLabel.ForeColor = $t.Red }

    Invoke-FolderAnalysis -Path $Path
    Show-FolderPreview -FolderPath $Path
    $statusLabel.Text = "Loaded: $([System.IO.Path]::GetFileName($Path))"; $statusLabel.ForeColor = $t.Success
}

function Invoke-FolderAnalysis {
    param([string]$Path, [switch]$UseCache)
    $t = Get-Theme; $global:AnalysisResults = @{}; $global:TotalFiles = 0; $global:AllFiles = @()

    if ([string]::IsNullOrEmpty($Path) -or -not (Test-Path $Path)) {
        $previewBox.Clear(); $previewBox.SelectionColor = $t.TextDim
        $previewBox.AppendText("Select a folder to analyze."); $organizeBtn.Enabled = $false; return
    }

    $cacheKey = "$Path|$($global:RecursiveMode)|$($chkShowHidden.Checked)"
    $files = $null
    if ($UseCache -and $global:AnalysisCache.ContainsKey($cacheKey)) { $files = $global:AnalysisCache[$cacheKey] }
    else {
        $gciParams = @{ Path = $Path; File = $true; ErrorAction = "SilentlyContinue" }
        if ($global:RecursiveMode) { $gciParams["Recurse"] = $true }
        if ($chkShowHidden.Checked) { $gciParams["Force"] = $true }
        $files = @(Get-ChildItem @gciParams)
        $global:AnalysisCache[$cacheKey] = $files
    }

    if (-not $files -or $files.Count -eq 0) {
        $previewBox.Clear(); $previewBox.SelectionColor = $t.TextDim
        $previewBox.AppendText("No files found$(if ($global:RecursiveMode) { ' (recursive)' }).")
        $organizeBtn.Enabled = $false; return
    }

    $selectedCategories = @()
    foreach ($entry in $global:CategoryCheckboxes.GetEnumerator()) { if ($entry.Value.Checked) { $selectedCategories += $entry.Key } }
    if ($selectedCategories.Count -eq 0) {
        $previewBox.Clear(); $previewBox.SelectionColor = $t.Orange
        $previewBox.AppendText("No categories selected."); $organizeBtn.Enabled = $false; return
    }

    $totalSize = [long]0
    foreach ($file in $files) {
        $category = Get-CategoryForFile -Extension $file.Extension
        if ($selectedCategories -contains $category) {
            if (-not $global:AnalysisResults.ContainsKey($category)) { $global:AnalysisResults[$category] = @() }
            $global:AnalysisResults[$category] += $file; $global:AllFiles += $file
            $global:TotalFiles++; $totalSize += $file.Length
        }
    }

    $previewBox.SuspendLayout(); $previewBox.Clear()
    if ($global:TotalFiles -eq 0) {
        $previewBox.SelectionColor = $t.TextDim; $previewBox.AppendText("No matching files.")
        $organizeBtn.Enabled = $false; $previewBox.ResumeLayout(); return
    }

    $previewBox.SelectionColor = if ($global:DryRunMode) { $t.DryRun } else { $t.Success }
    $previewBox.AppendText("ANALYSIS$(if ($global:DryRunMode) { '  [DRY RUN]' })`n")
    $previewBox.SelectionColor = $t.TextMuted
    $previewBox.AppendText(([string]::new([char]0x2500, 60)) + "`n`n")

    $folderName = [System.IO.Path]::GetFileName($Path)
    $previewBox.SelectionColor = $t.TextPrimary
    $previewBox.AppendText("  Folder:        $folderName`n")
    $previewBox.SelectionColor = $t.TextDim
    $previewBox.AppendText("  Path:          $Path`n")
    $previewBox.SelectionColor = $t.TextPrimary
    $previewBox.AppendText("  Mode:          $(if ($global:RecursiveMode) { 'Recursive' } else { 'Top-level' })`n")
    $previewBox.AppendText("  Files:         $($global:TotalFiles)`n")
    $previewBox.AppendText("  Size:          $(Get-FolderSizeText $totalSize)`n")
    $previewBox.AppendText("  Categories:    $($global:AnalysisResults.Count)`n")
    $previewBox.AppendText("  Conflict:      $($global:ConflictAction)`n")

    $skipped = $files.Count - $global:TotalFiles
    if ($skipped -gt 0) { $previewBox.SelectionColor = $t.TextMuted; $previewBox.AppendText("  Filtered:      $skipped`n") }
    $previewBox.AppendText("`n"); $previewBox.SelectionColor = $t.TextMuted
    $previewBox.AppendText(([string]::new([char]0x2500, 60)) + "`n`n")

    $maxPreview = if ($global:Config.MaxPreviewFiles) { [int]$global:Config.MaxPreviewFiles } else { 5 }
    $sorted = $global:AnalysisResults.GetEnumerator() | Sort-Object { $_.Value.Count } -Descending

    foreach ($cat in $sorted) {
        $catName = $cat.Key; $catFiles = $cat.Value; $count = $catFiles.Count
        $catSize = ($catFiles | Measure-Object -Property Length -Sum).Sum
        $pct = [math]::Round(($count / $global:TotalFiles) * 100, 1)
        $barLen = [math]::Max(1, [math]::Floor($pct / 4))
        $bar = ([string]::new([char]0x2588, $barLen)).PadRight(25)

        $previewBox.SelectionColor = $t.FolderText
        $previewBox.AppendText("  [$catName]`n")
        $previewBox.SelectionColor = $t.TextDim
        $previewBox.AppendText("    $count file(s)  |  $(Get-FolderSizeText $catSize)  |  $pct%`n")
        $previewBox.SelectionColor = $t.Success
        $previewBox.AppendText("    $bar`n")

        $show = [Math]::Min($maxPreview, $catFiles.Count)
        for ($i = 0; $i -lt $show; $i++) {
            $fname = $catFiles[$i].Name
            if ($fname.Length -gt 45) { $fname = $fname.Substring(0, 42) + "..." }
            $previewBox.SelectionColor = $t.TextMuted
            $previewBox.AppendText("      $fname  ($(Get-FolderSizeText $catFiles[$i].Length))`n")
        }
        if ($catFiles.Count -gt $maxPreview) {
            $previewBox.SelectionColor = [System.Drawing.Color]::FromArgb(90,90,100)
            $previewBox.AppendText("      ... and $($catFiles.Count - $maxPreview) more`n")
        }
        $previewBox.AppendText("`n")
    }

    $previewBox.SelectionColor = $t.TextMuted
    $previewBox.AppendText(([string]::new([char]0x2500, 60)) + "`n")
    $previewBox.SelectionColor = $t.TextDim
    $previewBox.AppendText("  Press ORGANIZE FILES (Ctrl+Enter) to proceed.`n")
    if ($global:DryRunMode) { $previewBox.SelectionColor = $t.DryRun; $previewBox.AppendText("  DRY RUN ON -- no files will move.`n") }

    $previewBox.SelectionStart = 0; $previewBox.ScrollToCaret(); $previewBox.ResumeLayout()
    $organizeBtn.Enabled = $true
    $statusLabel.Text = "Ready  |  $($global:TotalFiles) files  |  $(Get-FolderSizeText $totalSize)"
    $statusLabel.ForeColor = $t.Success
}

# ========================== HINTS ==========================
function Update-HintsPanel {
    $t = Get-Theme; $hintsBox.SuspendLayout(); $hintsBox.Clear()
    $shortcuts = @(
        @("Ctrl+B","Browse folder"), @("Ctrl+Enter","Organize"), @("Ctrl+Z","Undo all"),
        @("Ctrl+U","Undo category"), @("Ctrl+S","Export log"), @("Ctrl+L","Focus path"),
        @("Ctrl+T","Toggle theme"), @("Ctrl+O","Open Explorer"), @("Ctrl+A","Toggle all categories"),
        @("Ctrl+D","Toggle Dry Run"), @("Ctrl+R","Toggle Recursive"), @("Ctrl+I","Statistics"),
        @("Ctrl+K","Categories"), @("Ctrl+E","Recent folders"), @("F5","Refresh"), @("Escape","Clear status")
    )
    $hintsBox.SelectionColor = $t.Accent; $hintsBox.AppendText("KEYBOARD SHORTCUTS`n")
    $hintsBox.SelectionColor = $t.TextMuted; $hintsBox.AppendText(([string]::new([char]0x2500, 42)) + "`n")
    foreach ($s in $shortcuts) {
        $hintsBox.SelectionColor = $t.Success; $hintsBox.AppendText("  $($s[0].PadRight(15))")
        $hintsBox.SelectionColor = $t.TextDim; $hintsBox.AppendText("$($s[1])`n")
    }
    $hintsBox.SelectionColor = $t.TextMuted; $hintsBox.AppendText("`n" + ([string]::new([char]0x2500, 42)) + "`n")
    $hintsBox.SelectionColor = $t.Accent; $hintsBox.AppendText("TIPS`n")
    $hintsBox.SelectionColor = $t.TextMuted; $hintsBox.AppendText(([string]::new([char]0x2500, 42)) + "`n")
    $tips = @(
        "Drag and drop a folder onto the window.",
        "Browse opens the full Windows Explorer dialog.",
        "Dry Run simulates without moving files.",
        "Recursive mode includes subfolders.",
        "Create custom categories for your needs.",
        "Export Log saves a CSV of all moves.",
        "Files are never deleted, only moved.",
        "Empty folders cleaned up on undo.",
        "Analysis is cached for instant filtering."
    )
    foreach ($tip in $tips) { $hintsBox.SelectionColor = $t.TextDim; $hintsBox.AppendText("  * $tip`n") }
    $hintsBox.SelectionStart = 0; $hintsBox.ScrollToCaret(); $hintsBox.ResumeLayout()
}

# ========================== APPLY THEME ==========================
function Apply-Theme {
    $t = Get-Theme
    $form.BackColor = $t.Bg; $form.ForeColor = $t.TextPrimary; $mainPanel.BackColor = $t.Bg
    $headerPanel.BackColor = $t.Bg; $titleLabel.ForeColor = $t.Accent; $statusLabel.ForeColor = $t.TextDim
    foreach ($btn in @($btnStats,$btnCustomCat,$btnRecent,$themeToggleBtn)) {
        $btn.BackColor = $t.Input; $btn.ForeColor = $t.TextPrimary
        $btn.FlatAppearance.BorderColor = $t.Border; $btn.FlatAppearance.BorderSize = 1
    }
    $themeToggleBtn.Text = if ($global:CurrentTheme -eq "Dark") { "Light Mode" } else { "Dark Mode" }

    $folderPanel.BackColor = $t.Panel
    $selectFolderBtn.BackColor = $t.Accent; $selectFolderBtn.ForeColor = [System.Drawing.Color]::White; $selectFolderBtn.FlatAppearance.BorderSize = 0
    $folderPathBox.BackColor = $t.Input; $folderPathBox.ForeColor = $t.TextPrimary
    $btnGo.BackColor = $t.Green; $btnGo.ForeColor = [System.Drawing.Color]::White; $btnGo.FlatAppearance.BorderSize = 0
    $btnOpenExplorer.BackColor = $t.Input; $btnOpenExplorer.ForeColor = $t.TextPrimary
    $btnOpenExplorer.FlatAppearance.BorderColor = $t.Border; $btnOpenExplorer.FlatAppearance.BorderSize = 1
    $folderInfoLabel.ForeColor = $t.TextDim; $folderInfoLabel.BackColor = $t.Panel

    $modePanel.BackColor = $t.Panel
    foreach ($ctrl in @($chkDryRun,$chkRecursive,$chkShowHidden)) { $ctrl.ForeColor = $t.TextPrimary; $ctrl.BackColor = $t.Panel }
    $sepLabel.ForeColor = $t.TextMuted; $sepLabel.BackColor = $t.Panel
    $conflictLabel.ForeColor = $t.TextDim; $conflictLabel.BackColor = $t.Panel
    $conflictCombo.BackColor = $t.Input; $conflictCombo.ForeColor = $t.TextPrimary
    $modeStatusLabel.BackColor = $t.Panel

    $filterPanel.BackColor = $t.Bg; $filterFlow.BackColor = $t.Panel
    $chkAll.ForeColor = $t.Accent; $chkAll.BackColor = $t.Panel
    foreach ($chk in $global:CategoryCheckboxes.Values) { $chk.ForeColor = $t.TextPrimary; $chk.BackColor = $t.Panel }

    $previewGroupBox.BackColor = $t.Panel; $previewGroupBox.ForeColor = $t.TextDim
    $previewBox.BackColor = $t.ListBg; $previewBox.ForeColor = $t.TextPrimary
    $progressPanel.BackColor = $t.Bg; $progressLabel.ForeColor = $t.TextDim; $progressLabel.BackColor = $t.Bg

    $organizeBtn.BackColor = $t.Green; $organizeBtn.ForeColor = [System.Drawing.Color]::White
    $undoLastBtn.BackColor = $t.Orange; $undoLastBtn.ForeColor = [System.Drawing.Color]::White
    $undoAllBtn.BackColor = $t.Red; $undoAllBtn.ForeColor = [System.Drawing.Color]::White
    $exportLogBtn.BackColor = $t.Input; $exportLogBtn.ForeColor = $t.TextPrimary
    $exportLogBtn.FlatAppearance.BorderColor = $t.Border; $exportLogBtn.FlatAppearance.BorderSize = 1
    $refreshBtn.BackColor = $t.Input; $refreshBtn.ForeColor = $t.TextPrimary
    $refreshBtn.FlatAppearance.BorderColor = $t.Border; $refreshBtn.FlatAppearance.BorderSize = 1

    $bottomSplit.BackColor = $t.Bg
    $hintsGroupBox.BackColor = $t.Panel; $hintsGroupBox.ForeColor = $t.TextDim
    $hintsBox.BackColor = $t.HintBg; $hintsBox.ForeColor = $t.TextPrimary
    $filePreviewGroupBox.BackColor = $t.Panel; $filePreviewGroupBox.ForeColor = $t.TextDim
    $filePreviewPanel.BackColor = $t.ListBg; $filePreviewPicture.BackColor = $t.ListBg
    $filePreviewText.BackColor = $t.ListBg; $filePreviewText.ForeColor = $t.TextPrimary

    Update-HintsPanel; Update-ModeStatus
}

# ========================== BUTTON EVENTS ==========================
$selectFolderBtn.Add_Click({
    $initial = if ($global:SelectedPath -ne "") { $global:SelectedPath } elseif ($global:Config.LastFolder -ne "") { $global:Config.LastFolder } else { "" }
    $path = Select-FolderModern -Title "Select a folder to organize" -InitialPath $initial
    if ($path) { Set-SelectedFolder -Path $path }
})

$btnGo.Add_Click({
    $t = Get-Theme; $path = $folderPathBox.Text.Trim()
    if ($path -ne "" -and (Test-Path $path -PathType Container)) { Set-SelectedFolder -Path $path }
    else { $statusLabel.Text = "Invalid path"; $statusLabel.ForeColor = $t.Red }
})
$folderPathBox.Add_KeyDown({ if ($_.KeyCode -eq "Enter") { $btnGo.PerformClick(); $_.SuppressKeyPress = $true } })

$btnOpenExplorer.Add_Click({
    if ($global:SelectedPath -ne "" -and (Test-Path $global:SelectedPath)) { Start-Process "explorer.exe" $global:SelectedPath }
})

$refreshBtn.Add_Click({
    $global:AnalysisCache = @{}
    if ($global:SelectedPath -ne "" -and (Test-Path $global:SelectedPath)) { Set-SelectedFolder -Path $global:SelectedPath }
})

$themeToggleBtn.Add_Click({
    $global:CurrentTheme = if ($global:CurrentTheme -eq "Dark") { "Light" } else { "Dark" }
    Apply-Theme
    if ($global:SelectedPath -ne "") { Invoke-FolderAnalysis -Path $global:SelectedPath -UseCache }
})

# Recent Folders
$btnRecent.Add_Click({
    $t = Get-Theme
    if (-not $global:Config.RecentFolders -or $global:Config.RecentFolders.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No recent folders yet.", "Recent", "OK", "Information"); return
    }
    $rf = New-Object System.Windows.Forms.Form
    $rf.Text = "Recent Folders"; $rf.Size = New-Object System.Drawing.Size(620, 440)
    $rf.StartPosition = "CenterParent"; $rf.BackColor = $t.Bg; $rf.ForeColor = $t.TextPrimary
    $rf.FormBorderStyle = "FixedDialog"; $rf.MaximizeBox = $false; $rf.MinimizeBox = $false
    $rf.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $rfLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $rfLayout.Dock = "Fill"; $rfLayout.Padding = New-Object System.Windows.Forms.Padding(12)
    $rfLayout.RowCount = 3; $rfLayout.ColumnCount = 1
    $rfLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 100))) | Out-Null
    $rfLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 28))) | Out-Null
    $rfLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 100))) | Out-Null
    $rfLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 44))) | Out-Null
    $rf.Controls.Add($rfLayout)

    $rfLbl = New-Object System.Windows.Forms.Label
    $rfLbl.Text = "Double-click to select:"; $rfLbl.Dock = "Fill"; $rfLbl.ForeColor = $t.TextDim; $rfLbl.TextAlign = "MiddleLeft"
    $rfLayout.Controls.Add($rfLbl, 0, 0)

    $rfList = New-Object System.Windows.Forms.ListBox
    $rfList.Dock = "Fill"; $rfList.BackColor = $t.ListBg; $rfList.ForeColor = $t.TextPrimary
    $rfList.BorderStyle = "FixedSingle"; $rfList.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    foreach ($p in $global:Config.RecentFolders) { if (Test-Path $p) { $rfList.Items.Add($p) | Out-Null } }
    if ($rfList.Items.Count -gt 0) { $rfList.SelectedIndex = 0 }
    $rfLayout.Controls.Add($rfList, 0, 1)

    $rfBtn = New-Object System.Windows.Forms.Button
    $rfBtn.Text = "Open Selected"; $rfBtn.Dock = "Fill"; $rfBtn.FlatStyle = "Flat"
    $rfBtn.BackColor = $t.Accent; $rfBtn.ForeColor = [System.Drawing.Color]::White
    $rfBtn.FlatAppearance.BorderSize = 0; $rfBtn.Cursor = "Hand"
    $rfBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $rfBtn.Margin = New-Object System.Windows.Forms.Padding(0,4,0,0)
    $rfLayout.Controls.Add($rfBtn, 0, 2)

    $selectAction = { if ($rfList.SelectedItem) { $script:selectedRecent = $rfList.SelectedItem.ToString(); $rf.Close() } }
    $rfBtn.Add_Click($selectAction); $rfList.Add_DoubleClick($selectAction)
    $script:selectedRecent = $null; $rf.ShowDialog()
    if ($script:selectedRecent) { Set-SelectedFolder -Path $script:selectedRecent }
})

# Statistics
$btnStats.Add_Click({
    $t = Get-Theme
    $sf = New-Object System.Windows.Forms.Form
    $sf.Text = "Statistics"; $sf.Size = New-Object System.Drawing.Size(550, 540)
    $sf.StartPosition = "CenterParent"; $sf.BackColor = $t.Bg; $sf.ForeColor = $t.TextPrimary
    $sf.FormBorderStyle = "FixedDialog"; $sf.MaximizeBox = $false; $sf.MinimizeBox = $false
    $sf.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $sr = New-Object System.Windows.Forms.RichTextBox
    $sr.Dock = "Fill"; $sr.ReadOnly = $true; $sr.BorderStyle = "None"
    $sr.BackColor = $t.ListBg; $sr.ForeColor = $t.TextPrimary
    $sr.Font = New-Object System.Drawing.Font("Cascadia Code, Consolas", 9.5)
    $sr.Padding = New-Object System.Windows.Forms.Padding(12)

    $sr.SuspendLayout()
    $sr.SelectionColor = $t.Accent; $sr.AppendText("LIFETIME STATISTICS`n")
    $sr.SelectionColor = $t.TextMuted; $sr.AppendText(([string]::new([char]0x2500, 48)) + "`n`n")
    $sr.SelectionColor = $t.TextPrimary
    $sr.AppendText("  Files organized:   $($global:Stats.TotalFilesOrganized)`n")
    $sr.AppendText("  Data moved:        $(Get-FolderSizeText $global:Stats.TotalSizeMoved)`n")
    $sr.AppendText("  Operations:        $($global:Stats.TotalOperations)`n")
    $sr.AppendText("  Undo operations:   $($global:Stats.TotalUndoOperations)`n")
    $sr.AppendText("  First used:        $($global:Stats.FirstUsed)`n")
    $sr.AppendText("  Last used:         $($global:Stats.LastUsed)`n`n")

    $sr.SelectionColor = $t.TextMuted; $sr.AppendText(([string]::new([char]0x2500, 48)) + "`n")
    $sr.SelectionColor = $t.Accent; $sr.AppendText("BY CATEGORY`n")
    $sr.SelectionColor = $t.TextMuted; $sr.AppendText(([string]::new([char]0x2500, 48)) + "`n`n")
    if ($global:Stats.CategoryCounts -and $global:Stats.CategoryCounts.Count -gt 0) {
        $sorted = $global:Stats.CategoryCounts.GetEnumerator() | Sort-Object Value -Descending
        $maxVal = ($sorted | Select-Object -First 1).Value
        foreach ($entry in $sorted) {
            $barLen = if ($maxVal -gt 0) { [math]::Max(1, [math]::Floor(($entry.Value / $maxVal) * 15)) } else { 1 }
            $sr.SelectionColor = $t.FolderText; $sr.AppendText("  $($entry.Key.PadRight(16))")
            $sr.SelectionColor = $t.Success; $sr.AppendText("$([string]::new([char]0x2588, $barLen)) ")
            $sr.SelectionColor = $t.TextDim; $sr.AppendText("$($entry.Value)`n")
        }
    } else { $sr.SelectionColor = $t.TextDim; $sr.AppendText("  No data yet.`n") }

    $sr.AppendText("`n"); $sr.SelectionColor = $t.TextMuted; $sr.AppendText(([string]::new([char]0x2500, 48)) + "`n")
    $sr.SelectionColor = $t.Accent; $sr.AppendText("TOP EXTENSIONS`n")
    $sr.SelectionColor = $t.TextMuted; $sr.AppendText(([string]::new([char]0x2500, 48)) + "`n`n")
    if ($global:Stats.ExtensionCounts -and $global:Stats.ExtensionCounts.Count -gt 0) {
        $sortedExt = $global:Stats.ExtensionCounts.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 20
        $maxExt = ($sortedExt | Select-Object -First 1).Value
        foreach ($entry in $sortedExt) {
            $barLen = if ($maxExt -gt 0) { [math]::Max(1, [math]::Floor(($entry.Value / $maxExt) * 12)) } else { 1 }
            $sr.SelectionColor = $t.Success; $sr.AppendText("  $($entry.Key.PadRight(10))")
            $sr.SelectionColor = $t.Accent; $sr.AppendText("$([string]::new([char]0x2588, $barLen)) ")
            $sr.SelectionColor = $t.TextDim; $sr.AppendText("$($entry.Value)`n")
        }
    } else { $sr.SelectionColor = $t.TextDim; $sr.AppendText("  No data yet.`n") }

    $sr.SelectionStart = 0; $sr.ScrollToCaret(); $sr.ResumeLayout()
    $sf.Controls.Add($sr); $sf.ShowDialog()
})

# Custom Categories
$btnCustomCat.Add_Click({
    $t = Get-Theme
    $cf = New-Object System.Windows.Forms.Form
    $cf.Text = "Custom Categories"; $cf.Size = New-Object System.Drawing.Size(580, 500)
    $cf.StartPosition = "CenterParent"; $cf.BackColor = $t.Bg; $cf.ForeColor = $t.TextPrimary
    $cf.FormBorderStyle = "FixedDialog"; $cf.MaximizeBox = $false; $cf.MinimizeBox = $false
    $cf.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $cfLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $cfLayout.Dock = "Fill"; $cfLayout.Padding = New-Object System.Windows.Forms.Padding(12)
    $cfLayout.RowCount = 5; $cfLayout.ColumnCount = 1
    $cfLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 100))) | Out-Null
    $cfLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 28))) | Out-Null
    $cfLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 100))) | Out-Null
    $cfLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 34))) | Out-Null
    $cfLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 34))) | Out-Null
    $cfLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 44))) | Out-Null
    $cf.Controls.Add($cfLayout)

    $cfInfo = New-Object System.Windows.Forms.Label
    $cfInfo.Text = "Add categories with extensions (e.g. .dat .sav .custom)"
    $cfInfo.Dock = "Fill"; $cfInfo.ForeColor = $t.TextDim; $cfInfo.TextAlign = "MiddleLeft"
    $cfLayout.Controls.Add($cfInfo, 0, 0)

    $cfList = New-Object System.Windows.Forms.ListView
    $cfList.Dock = "Fill"; $cfList.View = "Details"; $cfList.FullRowSelect = $true
    $cfList.BackColor = $t.ListBg; $cfList.ForeColor = $t.TextPrimary; $cfList.BorderStyle = "FixedSingle"
    $cfList.Columns.Add("Category", 160) | Out-Null; $cfList.Columns.Add("Extensions", 370) | Out-Null
    $cfList.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    if ($global:Config.CustomCategories -and $global:Config.CustomCategories.Count -gt 0) {
        foreach ($entry in $global:Config.CustomCategories.GetEnumerator()) {
            $item = New-Object System.Windows.Forms.ListViewItem($entry.Key)
            $item.SubItems.Add(($entry.Value -join "  ")) | Out-Null; $cfList.Items.Add($item) | Out-Null
        }
    }
    $cfLayout.Controls.Add($cfList, 0, 1)

    $cfNP = New-Object System.Windows.Forms.TableLayoutPanel
    $cfNP.Dock = "Fill"; $cfNP.ColumnCount = 2; $cfNP.RowCount = 1
    $cfNP.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 115))) | Out-Null
    $cfNP.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 100))) | Out-Null
    $cfNL = New-Object System.Windows.Forms.Label; $cfNL.Text = "Category:"; $cfNL.Dock = "Fill"; $cfNL.TextAlign = "MiddleLeft"; $cfNL.ForeColor = $t.TextDim
    $cfNP.Controls.Add($cfNL, 0, 0)
    $cfNB = New-Object System.Windows.Forms.TextBox; $cfNB.Dock = "Fill"; $cfNB.BackColor = $t.Input; $cfNB.ForeColor = $t.TextPrimary
    $cfNB.BorderStyle = "FixedSingle"; $cfNB.Margin = New-Object System.Windows.Forms.Padding(0,4,0,4)
    $cfNP.Controls.Add($cfNB, 1, 0); $cfLayout.Controls.Add($cfNP, 0, 2)

    $cfEP = New-Object System.Windows.Forms.TableLayoutPanel
    $cfEP.Dock = "Fill"; $cfEP.ColumnCount = 2; $cfEP.RowCount = 1
    $cfEP.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 115))) | Out-Null
    $cfEP.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 100))) | Out-Null
    $cfEL = New-Object System.Windows.Forms.Label; $cfEL.Text = "Extensions:"; $cfEL.Dock = "Fill"; $cfEL.TextAlign = "MiddleLeft"; $cfEL.ForeColor = $t.TextDim
    $cfEP.Controls.Add($cfEL, 0, 0)
    $cfEB = New-Object System.Windows.Forms.TextBox; $cfEB.Dock = "Fill"; $cfEB.BackColor = $t.Input; $cfEB.ForeColor = $t.TextPrimary
    $cfEB.BorderStyle = "FixedSingle"; $cfEB.Margin = New-Object System.Windows.Forms.Padding(0,4,0,4)
    $cfEP.Controls.Add($cfEB, 1, 0); $cfLayout.Controls.Add($cfEP, 0, 3)

    $cfBP = New-Object System.Windows.Forms.TableLayoutPanel
    $cfBP.Dock = "Fill"; $cfBP.ColumnCount = 2; $cfBP.RowCount = 1
    $cfBP.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 50))) | Out-Null
    $cfBP.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 50))) | Out-Null

    $cfAdd = New-Object System.Windows.Forms.Button
    $cfAdd.Text = "Add"; $cfAdd.Dock = "Fill"; $cfAdd.FlatStyle = "Flat"
    $cfAdd.BackColor = $t.Green; $cfAdd.ForeColor = [System.Drawing.Color]::White; $cfAdd.FlatAppearance.BorderSize = 0
    $cfAdd.Cursor = "Hand"; $cfAdd.Margin = New-Object System.Windows.Forms.Padding(0,2,4,2)
    $cfBP.Controls.Add($cfAdd, 0, 0)

    $cfRem = New-Object System.Windows.Forms.Button
    $cfRem.Text = "Remove"; $cfRem.Dock = "Fill"; $cfRem.FlatStyle = "Flat"
    $cfRem.BackColor = $t.Red; $cfRem.ForeColor = [System.Drawing.Color]::White; $cfRem.FlatAppearance.BorderSize = 0
    $cfRem.Cursor = "Hand"; $cfRem.Margin = New-Object System.Windows.Forms.Padding(4,2,0,2)
    $cfBP.Controls.Add($cfRem, 1, 0); $cfLayout.Controls.Add($cfBP, 0, 4)

    $cfAdd.Add_Click({
        $name = $cfNB.Text.Trim(); $exts = $cfEB.Text.Trim()
        if ($name -eq "" -or $exts -eq "") { [System.Windows.Forms.MessageBox]::Show("Enter name and extensions.","Missing","OK","Warning"); return }
        $extArray = @($exts -split '\s+|,|;' | Where-Object { $_ -ne "" } | ForEach-Object { $e = $_.Trim().ToLower(); if (-not $e.StartsWith(".")) { $e = ".$e" }; $e })
        if ($extArray.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Invalid extensions.","Error","OK","Warning"); return }

        if (-not $global:Config.CustomCategories) { $global:Config.CustomCategories = @{} }
        $global:Config.CustomCategories[$name] = $extArray; Save-Config -Config $global:Config
        $global:FileCategories[$name] = $extArray
        foreach ($ext in $extArray) { $global:ExtensionLookup[$ext.ToLower()] = $name }

        if (-not $global:CategoryCheckboxes.ContainsKey($name)) {
            $chk = New-Object System.Windows.Forms.CheckBox
            $chk.Text = $name; $chk.Checked = $true; $chk.AutoSize = $true
            $chk.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
            $chk.Margin = New-Object System.Windows.Forms.Padding(3,4,3,4); $chk.Cursor = "Hand"
            $chk.ForeColor = $t.TextPrimary; $chk.BackColor = $t.Panel
            $toolTip.SetToolTip($chk, "$name`n$($extArray -join '  ')")
            $chk.Add_CheckedChanged({
                if ($global:UpdatingCheckboxes) { return }; $global:UpdatingCheckboxes = $true
                $allOn = $true; foreach ($c in $global:CategoryCheckboxes.Values) { if (-not $c.Checked) { $allOn = $false; break } }
                $chkAll.Checked = $allOn; $global:UpdatingCheckboxes = $false
                if ($global:SelectedPath -ne "") { Invoke-FolderAnalysis -Path $global:SelectedPath -UseCache }
            })
            $filterFlow.Controls.Add($chk); $global:CategoryCheckboxes[$name] = $chk
        }
        $item = New-Object System.Windows.Forms.ListViewItem($name)
        $item.SubItems.Add(($extArray -join "  ")) | Out-Null; $cfList.Items.Add($item) | Out-Null
        $cfNB.Text = ""; $cfEB.Text = ""
        $global:AnalysisCache = @{}; if ($global:SelectedPath -ne "") { Invoke-FolderAnalysis -Path $global:SelectedPath }
    })

    $cfRem.Add_Click({
        if ($cfList.SelectedItems.Count -eq 0) { return }
        $selName = $cfList.SelectedItems[0].Text
        $global:Config.CustomCategories.Remove($selName); Save-Config -Config $global:Config
        $global:FileCategories.Remove($selName)
        $keysToRemove = @($global:ExtensionLookup.GetEnumerator() | Where-Object { $_.Value -eq $selName } | ForEach-Object { $_.Key })
        foreach ($k in $keysToRemove) { $global:ExtensionLookup.Remove($k) }
        if ($global:CategoryCheckboxes.ContainsKey($selName)) {
            $filterFlow.Controls.Remove($global:CategoryCheckboxes[$selName]); $global:CategoryCheckboxes.Remove($selName)
        }
        $cfList.Items.Remove($cfList.SelectedItems[0])
        $global:AnalysisCache = @{}; if ($global:SelectedPath -ne "") { Invoke-FolderAnalysis -Path $global:SelectedPath }
    })
    $cf.ShowDialog()
})

# Export Log
$exportLogBtn.Add_Click({
    if ($global:MoveLog.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No moves to export.","Export","OK","Information"); return }
    $logFile = Export-MoveLog -Log $global:MoveLog -FolderPath $global:SelectedPath
    $result = [System.Windows.Forms.MessageBox]::Show("Log saved to:`n$logFile`n`nOpen it?","Exported",
        [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information)
    if ($result -eq "Yes") { Start-Process $logFile }
})

# ========================== ORGANIZE ==========================
$organizeBtn.Add_Click({
    $t = Get-Theme
    if ([string]::IsNullOrEmpty($global:SelectedPath) -or $global:TotalFiles -eq 0 -or $global:IsOrganizing) { return }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "$(if ($global:DryRunMode) {'Simulate'} else {'Move'}) $($global:TotalFiles) file(s)?`n`nFolder: $($global:SelectedPath)`nConflict: $($global:ConflictAction)$(if ($global:DryRunMode) {"`n`nDRY RUN: No files will move."})",
        "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne "Yes") { return }

    $global:IsOrganizing = $true
    $organizeBtn.Enabled = $false; $undoLastBtn.Enabled = $false; $undoAllBtn.Enabled = $false
    $exportLogBtn.Enabled = $false; $selectFolderBtn.Enabled = $false; $refreshBtn.Enabled = $false
    $global:MoveLog = @()
    $progressBar.Maximum = $global:TotalFiles; $progressBar.Value = 0
    $moved = 0; $skipped = 0; $errors = 0
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

    foreach ($catEntry in $global:AnalysisResults.GetEnumerator()) {
        $category = $catEntry.Key; $catFiles = $catEntry.Value
        $destFolder = Join-Path $global:SelectedPath $category

        if (-not $global:DryRunMode -and -not (Test-Path $destFolder)) {
            try { New-Item -Path $destFolder -ItemType Directory -Force | Out-Null }
            catch { Write-ErrorLog "Cannot create $destFolder : $_" "Organize"; $errors += $catFiles.Count; continue }
        }

        foreach ($file in $catFiles) {
            $destPath = Join-Path $destFolder $file.Name
            if (-not $global:DryRunMode -and (Test-Path $destPath)) {
                switch ($global:ConflictAction) {
                    "Skip" { $skipped++; $progressBar.Value = [Math]::Min($progressBar.Maximum, $moved+$skipped+$errors)
                        $progressLabel.Text = "$($moved+$skipped+$errors) / $($global:TotalFiles)"; $form.Refresh(); continue }
                    "Rename" { $destPath = Get-UniqueFileName -DestinationFolder $destFolder -FileName $file.Name }
                    "Overwrite" {}
                }
            }

            if ($global:DryRunMode) {
                $global:MoveLog += [PSCustomObject]@{ OriginalPath=$file.FullName; DestinationPath=$destPath; Category=$category; FileName=$file.Name; FileSize=$file.Length; Timestamp=$timestamp }
                $moved++
            } else {
                try {
                    Move-Item -Path $file.FullName -Destination $destPath -Force -ErrorAction Stop
                    $global:MoveLog += [PSCustomObject]@{ OriginalPath=$file.FullName; DestinationPath=$destPath; Category=$category; FileName=$file.Name; FileSize=$file.Length; Timestamp=$timestamp }
                    $moved++
                } catch { $errors++; Write-ErrorLog "Move failed $($file.FullName): $_" "Organize" }
            }
            $progressBar.Value = [Math]::Min($progressBar.Maximum, $moved+$skipped+$errors)
            $progressLabel.Text = "$($moved+$skipped+$errors) / $($global:TotalFiles)"
            $statusLabel.Text = "$(if ($global:DryRunMode) {'Simulating'} else {'Moving'})..."
            $statusLabel.ForeColor = if ($global:DryRunMode) { $t.DryRun } else { $t.Orange }
            $form.Refresh()
        }
    }

    $progressBar.Value = $progressBar.Maximum
    if (-not $global:DryRunMode -and $moved -gt 0) {
        $global:Stats.TotalFilesOrganized += $moved
        $global:Stats.TotalSizeMoved += ($global:MoveLog | Measure-Object -Property FileSize -Sum).Sum
        $global:Stats.TotalOperations++
        foreach ($entry in $global:MoveLog) {
            if (-not $global:Stats.CategoryCounts.ContainsKey($entry.Category)) { $global:Stats.CategoryCounts[$entry.Category] = 0 }
            $global:Stats.CategoryCounts[$entry.Category] = [int]$global:Stats.CategoryCounts[$entry.Category] + 1
            $extKey = [System.IO.Path]::GetExtension($entry.FileName).ToLower()
            if ($extKey -ne "") {
                if (-not $global:Stats.ExtensionCounts.ContainsKey($extKey)) { $global:Stats.ExtensionCounts[$extKey] = 0 }
                $global:Stats.ExtensionCounts[$extKey] = [int]$global:Stats.ExtensionCounts[$extKey] + 1
            }
        }
        Save-Stats -Stats $global:Stats; Add-RecentFolder -Path $global:SelectedPath
    }

    $totalMovedSize = ($global:MoveLog | Measure-Object -Property FileSize -Sum).Sum
    $previewBox.SuspendLayout(); $previewBox.Clear()
    $previewBox.SelectionColor = if ($global:DryRunMode) { $t.DryRun } else { $t.Success }
    $previewBox.AppendText("$(if ($global:DryRunMode) {'DRY RUN COMPLETE'} else {'ORGANIZED'})`n")
    $previewBox.SelectionColor = $t.TextMuted; $previewBox.AppendText(([string]::new([char]0x2500, 60)) + "`n`n")
    $previewBox.SelectionColor = $t.TextPrimary
    $previewBox.AppendText("  Folder:      $($global:SelectedPath)`n  Moved:       $moved`n")
    if ($skipped -gt 0) { $previewBox.SelectionColor = $t.Orange; $previewBox.AppendText("  Skipped:     $skipped`n") }
    if ($errors -gt 0) { $previewBox.SelectionColor = $t.Red; $previewBox.AppendText("  Errors:      $errors`n") }
    $previewBox.SelectionColor = $t.TextDim; $previewBox.AppendText("  Size:        $(Get-FolderSizeText $totalMovedSize)`n`n")

    $createdFolders = $global:MoveLog | Group-Object Category | Sort-Object Count -Descending
    $previewBox.SelectionColor = $t.FolderText; $previewBox.AppendText("  Folders:`n")
    foreach ($g in $createdFolders) {
        $gs = ($g.Group | Measure-Object -Property FileSize -Sum).Sum
        $previewBox.SelectionColor = $t.TextDim; $previewBox.AppendText("    [$($g.Name)]  $($g.Count) file(s)  |  $(Get-FolderSizeText $gs)`n")
    }
    $previewBox.AppendText("`n"); $previewBox.SelectionColor = $t.TextMuted
    if ($global:DryRunMode) { $previewBox.AppendText("  No files moved. Disable Dry Run to organize.`n") }
    else { $previewBox.AppendText("  Undo Category (Ctrl+U) | Undo All (Ctrl+Z) | Export (Ctrl+S)`n") }
    $previewBox.SelectionStart = 0; $previewBox.ScrollToCaret(); $previewBox.ResumeLayout()

    $statusLabel.Text = "$(if ($global:DryRunMode) {'Simulated'} else {'Done'})  |  $moved moved  |  $skipped skipped  |  $errors errors"
    $statusLabel.ForeColor = if ($global:DryRunMode) { $t.DryRun } else { $t.Success }
    $progressLabel.Text = "Complete"

    $global:IsOrganizing = $false; $selectFolderBtn.Enabled = $true; $refreshBtn.Enabled = $true
    if ($global:MoveLog.Count -gt 0 -and -not $global:DryRunMode) { $undoLastBtn.Enabled = $true; $undoAllBtn.Enabled = $true; $exportLogBtn.Enabled = $true }
    elseif ($global:DryRunMode -and $global:MoveLog.Count -gt 0) { $exportLogBtn.Enabled = $true }
    $global:AnalysisCache = @{}

    # Update folder preview
    if (-not $global:DryRunMode) { Show-FolderPreview -FolderPath $global:SelectedPath }
})

# ========================== UNDO CATEGORY ==========================
$undoLastBtn.Add_Click({
    $t = Get-Theme
    if ($global:MoveLog.Count -eq 0) { return }
    $categories = $global:MoveLog | Select-Object -ExpandProperty Category -Unique

    $uf = New-Object System.Windows.Forms.Form
    $uf.Text = "Undo Category"; $uf.Size = New-Object System.Drawing.Size(480, 400)
    $uf.StartPosition = "CenterParent"; $uf.BackColor = $t.Bg; $uf.ForeColor = $t.TextPrimary
    $uf.FormBorderStyle = "FixedDialog"; $uf.MaximizeBox = $false; $uf.MinimizeBox = $false
    $uf.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $ufL = New-Object System.Windows.Forms.TableLayoutPanel
    $ufL.Dock = "Fill"; $ufL.Padding = New-Object System.Windows.Forms.Padding(12)
    $ufL.RowCount = 3; $ufL.ColumnCount = 1
    $ufL.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 100))) | Out-Null
    $ufL.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 28))) | Out-Null
    $ufL.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 100))) | Out-Null
    $ufL.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 44))) | Out-Null
    $uf.Controls.Add($ufL)

    $uLbl = New-Object System.Windows.Forms.Label
    $uLbl.Text = "Select category to restore:"; $uLbl.Dock = "Fill"; $uLbl.ForeColor = $t.TextDim; $uLbl.TextAlign = "MiddleLeft"
    $ufL.Controls.Add($uLbl, 0, 0)

    $uList = New-Object System.Windows.Forms.ListBox
    $uList.Dock = "Fill"; $uList.BackColor = $t.ListBg; $uList.ForeColor = $t.TextPrimary
    $uList.BorderStyle = "FixedSingle"; $uList.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    foreach ($cat in $categories) {
        $ce = $global:MoveLog | Where-Object { $_.Category -eq $cat }
        $cs = ($ce | Measure-Object -Property FileSize -Sum).Sum
        $uList.Items.Add("$cat  --  $($ce.Count) file(s)  |  $(Get-FolderSizeText $cs)") | Out-Null
    }
    if ($uList.Items.Count -gt 0) { $uList.SelectedIndex = 0 }
    $ufL.Controls.Add($uList, 0, 1)

    $uBtn = New-Object System.Windows.Forms.Button
    $uBtn.Text = "Undo Selected"; $uBtn.Dock = "Fill"; $uBtn.FlatStyle = "Flat"
    $uBtn.BackColor = $t.Orange; $uBtn.ForeColor = [System.Drawing.Color]::White; $uBtn.FlatAppearance.BorderSize = 0
    $uBtn.Cursor = "Hand"; $uBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $uBtn.Margin = New-Object System.Windows.Forms.Padding(0,4,0,0)
    $ufL.Controls.Add($uBtn, 0, 2)

    $uBtn.Add_Click({
        if ($uList.SelectedIndex -lt 0) { return }
        $selCat = $uList.SelectedItem.ToString().Split("  ")[0].Trim()
        $entries = $global:MoveLog | Where-Object { $_.Category -eq $selCat }
        $restored = 0; $ue = 0
        foreach ($entry in $entries) {
            try {
                if (Test-Path $entry.DestinationPath) {
                    $parentDir = Split-Path $entry.OriginalPath -Parent
                    if (-not (Test-Path $parentDir)) { New-Item -Path $parentDir -ItemType Directory -Force | Out-Null }
                    Move-Item -Path $entry.DestinationPath -Destination $entry.OriginalPath -Force -ErrorAction Stop; $restored++
                }
            } catch { $ue++; Write-ErrorLog "Undo failed $($entry.FileName): $_" "UndoCat" }
        }
        $catFolder = Join-Path $global:SelectedPath $selCat
        if ((Test-Path $catFolder) -and ((Get-ChildItem $catFolder -ErrorAction SilentlyContinue).Count -eq 0)) { Remove-Item $catFolder -Force -ErrorAction SilentlyContinue }
        $global:MoveLog = @($global:MoveLog | Where-Object { $_.Category -ne $selCat })
        $global:Stats.TotalUndoOperations++; Save-Stats -Stats $global:Stats
        [System.Windows.Forms.MessageBox]::Show("Restored $restored from '$selCat'.$(if ($ue -gt 0) {" Errors: $ue"})","Done","OK","Information")
        $uf.Close()
    })
    $uf.ShowDialog()
    if ($global:MoveLog.Count -eq 0) { $undoLastBtn.Enabled = $false; $undoAllBtn.Enabled = $false; $exportLogBtn.Enabled = $false }
    $global:AnalysisCache = @{}
    if ($global:SelectedPath -ne "" -and (Test-Path $global:SelectedPath)) { Invoke-FolderAnalysis -Path $global:SelectedPath; Show-FolderPreview -FolderPath $global:SelectedPath }
})

# ========================== UNDO ALL ==========================
$undoAllBtn.Add_Click({
    $t = Get-Theme
    if ($global:MoveLog.Count -eq 0) { return }
    $totalUndoSize = ($global:MoveLog | Measure-Object -Property FileSize -Sum).Sum
    $confirm = [System.Windows.Forms.MessageBox]::Show("Restore all $($global:MoveLog.Count) file(s)?`n$(Get-FolderSizeText $totalUndoSize)",
        "Undo All", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($confirm -ne "Yes") { return }

    $undoAllBtn.Enabled = $false; $undoLastBtn.Enabled = $false; $exportLogBtn.Enabled = $false; $selectFolderBtn.Enabled = $false
    $restored = 0; $ue = 0; $progressBar.Maximum = $global:MoveLog.Count; $progressBar.Value = 0

    foreach ($entry in $global:MoveLog) {
        try {
            if (Test-Path $entry.DestinationPath) {
                $parentDir = Split-Path $entry.OriginalPath -Parent
                if (-not (Test-Path $parentDir)) { New-Item -Path $parentDir -ItemType Directory -Force | Out-Null }
                Move-Item -Path $entry.DestinationPath -Destination $entry.OriginalPath -Force -ErrorAction Stop; $restored++
            }
        } catch { $ue++; Write-ErrorLog "Undo failed $($entry.FileName): $_" "UndoAll" }
        $progressBar.Value = [Math]::Min($progressBar.Maximum, $restored+$ue)
        $progressLabel.Text = "$($restored+$ue) / $($global:MoveLog.Count)"
        $statusLabel.Text = "Undoing..."; $statusLabel.ForeColor = $t.Orange; $form.Refresh()
    }

    foreach ($cat in ($global:MoveLog | Select-Object -ExpandProperty Category -Unique)) {
        $catFolder = Join-Path $global:SelectedPath $cat
        if ((Test-Path $catFolder) -and ((Get-ChildItem $catFolder -ErrorAction SilentlyContinue).Count -eq 0)) { Remove-Item $catFolder -Force -ErrorAction SilentlyContinue }
    }
    $global:Stats.TotalUndoOperations++; Save-Stats -Stats $global:Stats
    $statusLabel.Text = "Restored: $restored  |  Errors: $ue"; $statusLabel.ForeColor = $t.Success
    $progressLabel.Text = "Undo complete"; $global:MoveLog = @(); $selectFolderBtn.Enabled = $true
    $global:AnalysisCache = @{}
    if ($global:SelectedPath -ne "" -and (Test-Path $global:SelectedPath)) { Invoke-FolderAnalysis -Path $global:SelectedPath; Show-FolderPreview -FolderPath $global:SelectedPath }
})

# ========================== KEYBOARD SHORTCUTS ==========================
$form.Add_KeyDown({
    $t = Get-Theme
    if ($_.Control -and $_.KeyCode -eq "B") { $selectFolderBtn.PerformClick(); $_.Handled = $true; $_.SuppressKeyPress = $true }
    elseif ($_.Control -and $_.KeyCode -eq "Return") { if ($organizeBtn.Enabled) { $organizeBtn.PerformClick() }; $_.Handled = $true }
    elseif ($_.Control -and $_.KeyCode -eq "Z") { if ($undoAllBtn.Enabled) { $undoAllBtn.PerformClick() }; $_.Handled = $true }
    elseif ($_.Control -and $_.KeyCode -eq "U") { if ($undoLastBtn.Enabled) { $undoLastBtn.PerformClick() }; $_.Handled = $true }
    elseif ($_.Control -and $_.KeyCode -eq "S") { if ($exportLogBtn.Enabled) { $exportLogBtn.PerformClick() }; $_.Handled = $true; $_.SuppressKeyPress = $true }
    elseif ($_.Control -and $_.KeyCode -eq "L") { $folderPathBox.Focus(); $folderPathBox.SelectAll(); $_.Handled = $true; $_.SuppressKeyPress = $true }
    elseif ($_.Control -and $_.KeyCode -eq "T") { $themeToggleBtn.PerformClick(); $_.Handled = $true; $_.SuppressKeyPress = $true }
    elseif ($_.Control -and $_.KeyCode -eq "O") { $btnOpenExplorer.PerformClick(); $_.Handled = $true; $_.SuppressKeyPress = $true }
    elseif ($_.Control -and $_.KeyCode -eq "A") {
        if ($form.ActiveControl -isnot [System.Windows.Forms.TextBox]) { $chkAll.Checked = -not $chkAll.Checked; $_.Handled = $true; $_.SuppressKeyPress = $true }
    }
    elseif ($_.Control -and $_.KeyCode -eq "D") { $chkDryRun.Checked = -not $chkDryRun.Checked; $_.Handled = $true; $_.SuppressKeyPress = $true }
    elseif ($_.Control -and $_.KeyCode -eq "R") { $chkRecursive.Checked = -not $chkRecursive.Checked; $_.Handled = $true; $_.SuppressKeyPress = $true }
    elseif ($_.Control -and $_.KeyCode -eq "I") { $btnStats.PerformClick(); $_.Handled = $true; $_.SuppressKeyPress = $true }
    elseif ($_.Control -and $_.KeyCode -eq "K") { $btnCustomCat.PerformClick(); $_.Handled = $true; $_.SuppressKeyPress = $true }
    elseif ($_.Control -and $_.KeyCode -eq "E") { $btnRecent.PerformClick(); $_.Handled = $true; $_.SuppressKeyPress = $true }
    elseif ($_.KeyCode -eq "F5") { $refreshBtn.PerformClick(); $_.Handled = $true }
    elseif ($_.KeyCode -eq "Escape") { $statusLabel.Text = "Ready"; $statusLabel.ForeColor = $t.TextDim; $progressBar.Value = 0; $progressLabel.Text = ""; $_.Handled = $true }
})

# ========================== INIT ==========================
Apply-Theme
if ($global:Config.LastFolder -ne "" -and (Test-Path $global:Config.LastFolder)) { Set-SelectedFolder -Path $global:Config.LastFolder }
[void]$form.ShowDialog()
