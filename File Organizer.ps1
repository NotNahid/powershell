Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationCore

# ========================== CONFIG / PERSISTENCE ==========================
$global:AppDataFolder = Join-Path $env:APPDATA "FileOrganizerPro"
$global:ConfigFile = Join-Path $global:AppDataFolder "config.json"
$global:StatsFile = Join-Path $global:AppDataFolder "stats.json"
$global:LogFolder = Join-Path $global:AppDataFolder "logs"

if (-not (Test-Path $global:AppDataFolder)) { New-Item -Path $global:AppDataFolder -ItemType Directory -Force | Out-Null }
if (-not (Test-Path $global:LogFolder)) { New-Item -Path $global:LogFolder -ItemType Directory -Force | Out-Null }

function Load-Config {
    $default = @{
        Theme           = "Dark"
        RecentFolders   = @()
        CustomCategories = @{}
        RecursiveMode   = $false
        DryRunMode      = $false
        ConflictAction  = "Rename"
        WindowWidth     = 1060
        WindowHeight    = 820
    }
    if (Test-Path $global:ConfigFile) {
        try {
            $loaded = Get-Content $global:ConfigFile -Raw | ConvertFrom-Json
            $config = @{}
            foreach ($prop in $default.Keys) {
                if ($loaded.PSObject.Properties.Name -contains $prop) {
                    $val = $loaded.$prop
                    if ($val -is [System.Management.Automation.PSCustomObject]) {
                        $ht = @{}
                        $val.PSObject.Properties | ForEach-Object { $ht[$_.Name] = @($_.Value) }
                        $config[$prop] = $ht
                    }
                    elseif ($val -is [System.Object[]]) {
                        $config[$prop] = @($val)
                    }
                    else {
                        $config[$prop] = $val
                    }
                }
                else { $config[$prop] = $default[$prop] }
            }
            return $config
        }
        catch { return $default }
    }
    return $default
}

function Save-Config {
    param([hashtable]$Config)
    try { $Config | ConvertTo-Json -Depth 5 | Set-Content $global:ConfigFile -Force } catch {}
}

function Load-Stats {
    $default = @{
        TotalFilesOrganized = 0
        TotalSizeMoved      = 0
        TotalOperations      = 0
        CategoryCounts       = @{}
        ExtensionCounts      = @{}
        FirstUsed            = (Get-Date).ToString("yyyy-MM-dd HH:mm")
        LastUsed             = (Get-Date).ToString("yyyy-MM-dd HH:mm")
    }
    if (Test-Path $global:StatsFile) {
        try {
            $loaded = Get-Content $global:StatsFile -Raw | ConvertFrom-Json
            $stats = @{}
            foreach ($prop in $default.Keys) {
                if ($loaded.PSObject.Properties.Name -contains $prop) {
                    $val = $loaded.$prop
                    if ($val -is [System.Management.Automation.PSCustomObject]) {
                        $ht = @{}
                        $val.PSObject.Properties | ForEach-Object { $ht[$_.Name] = $_.Value }
                        $stats[$prop] = $ht
                    }
                    else { $stats[$prop] = $val }
                }
                else { $stats[$prop] = $default[$prop] }
            }
            return $stats
        }
        catch { return $default }
    }
    return $default
}

function Save-Stats {
    param([hashtable]$Stats)
    try { $Stats | ConvertTo-Json -Depth 5 | Set-Content $global:StatsFile -Force } catch {}
}

$global:Config = Load-Config
$global:Stats = Load-Stats

# ========================== FILE CATEGORIES ==========================
$global:FileCategories = [ordered]@{
    "Images"      = @(".jpg", ".jpeg", ".png", ".gif", ".bmp", ".svg", ".webp", ".ico", ".tiff", ".tif", ".raw", ".heic", ".heif")
    "Videos"      = @(".mp4", ".avi", ".mkv", ".mov", ".wmv", ".flv", ".webm", ".m4v", ".mpg", ".mpeg", ".3gp")
    "Audio"       = @(".mp3", ".wav", ".flac", ".aac", ".ogg", ".wma", ".m4a", ".opus", ".aiff", ".alac")
    "Documents"   = @(".pdf", ".doc", ".docx", ".txt", ".rtf", ".odt", ".xls", ".xlsx", ".ppt", ".pptx", ".csv", ".epub", ".pages", ".numbers", ".key")
    "Archives"    = @(".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz", ".iso", ".cab")
    "Code"        = @(".py", ".js", ".ts", ".html", ".htm", ".css", ".java", ".cpp", ".c", ".cs", ".ps1", ".sh", ".bash", ".json", ".xml", ".yaml", ".yml", ".sql", ".php", ".rb", ".go", ".rs", ".swift", ".kt", ".r", ".lua", ".pl", ".md", ".ini", ".cfg", ".conf", ".toml", ".log")
    "Executables" = @(".exe", ".msi", ".bat", ".cmd", ".com", ".scr", ".appx", ".msix")
    "Fonts"       = @(".ttf", ".otf", ".woff", ".woff2", ".eot", ".fon")
    "3D Models"   = @(".obj", ".fbx", ".stl", ".blend", ".3ds", ".dae", ".gltf", ".glb")
    "Design"      = @(".psd", ".ai", ".xd", ".fig", ".sketch", ".indd", ".cdr")
    "Torrents"    = @(".torrent")
}

# Load custom categories from config
if ($global:Config.CustomCategories -and $global:Config.CustomCategories.Count -gt 0) {
    foreach ($entry in $global:Config.CustomCategories.GetEnumerator()) {
        $global:FileCategories[$entry.Key] = @($entry.Value)
    }
}

# ========================== THEME DEFINITIONS ==========================
$global:Themes = @{
    "Dark" = @{
        Bg          = [System.Drawing.Color]::FromArgb(25, 25, 30)
        Panel       = [System.Drawing.Color]::FromArgb(35, 35, 42)
        Input       = [System.Drawing.Color]::FromArgb(45, 45, 55)
        Accent      = [System.Drawing.Color]::FromArgb(0, 122, 204)
        Green       = [System.Drawing.Color]::FromArgb(46, 160, 67)
        Red         = [System.Drawing.Color]::FromArgb(200, 60, 60)
        Orange      = [System.Drawing.Color]::FromArgb(220, 160, 40)
        TextPrimary = [System.Drawing.Color]::White
        TextDim     = [System.Drawing.Color]::FromArgb(160, 160, 170)
        TextMuted   = [System.Drawing.Color]::FromArgb(110, 110, 120)
        Border      = [System.Drawing.Color]::FromArgb(60, 60, 70)
        ListBg      = [System.Drawing.Color]::FromArgb(30, 30, 38)
        FolderText  = [System.Drawing.Color]::FromArgb(100, 180, 255)
        DriveText   = [System.Drawing.Color]::FromArgb(255, 200, 80)
        QuickText   = [System.Drawing.Color]::FromArgb(120, 220, 150)
        Success     = [System.Drawing.Color]::FromArgb(0, 200, 120)
        HighlightBg = [System.Drawing.Color]::FromArgb(50, 50, 60)
        HintBg      = [System.Drawing.Color]::FromArgb(40, 40, 50)
        DryRun      = [System.Drawing.Color]::FromArgb(180, 120, 255)
    }
    "Light" = @{
        Bg          = [System.Drawing.Color]::FromArgb(240, 240, 245)
        Panel       = [System.Drawing.Color]::FromArgb(255, 255, 255)
        Input       = [System.Drawing.Color]::FromArgb(245, 245, 248)
        Accent      = [System.Drawing.Color]::FromArgb(0, 102, 184)
        Green       = [System.Drawing.Color]::FromArgb(36, 140, 57)
        Red         = [System.Drawing.Color]::FromArgb(190, 50, 50)
        Orange      = [System.Drawing.Color]::FromArgb(200, 140, 20)
        TextPrimary = [System.Drawing.Color]::FromArgb(30, 30, 35)
        TextDim     = [System.Drawing.Color]::FromArgb(100, 100, 110)
        TextMuted   = [System.Drawing.Color]::FromArgb(140, 140, 150)
        Border      = [System.Drawing.Color]::FromArgb(200, 200, 210)
        ListBg      = [System.Drawing.Color]::FromArgb(250, 250, 252)
        FolderText  = [System.Drawing.Color]::FromArgb(0, 90, 180)
        DriveText   = [System.Drawing.Color]::FromArgb(180, 130, 0)
        QuickText   = [System.Drawing.Color]::FromArgb(30, 150, 80)
        Success     = [System.Drawing.Color]::FromArgb(0, 160, 90)
        HighlightBg = [System.Drawing.Color]::FromArgb(230, 235, 245)
        HintBg      = [System.Drawing.Color]::FromArgb(235, 238, 245)
        DryRun      = [System.Drawing.Color]::FromArgb(130, 80, 200)
    }
}

$global:CurrentTheme = $global:Config.Theme
function Get-Theme { return $global:Themes[$global:CurrentTheme] }

# ========================== HELPER FUNCTIONS ==========================
function Get-CategoryForFile {
    param([string]$Extension)
    $ext = $Extension.ToLower()
    foreach ($category in $global:FileCategories.GetEnumerator()) {
        if ($category.Value -contains $ext) { return $category.Key }
    }
    return "Other"
}

function Get-UniqueFileName {
    param([string]$DestinationFolder, [string]$FileName)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $extension = [System.IO.Path]::GetExtension($FileName)
    $finalPath = Join-Path $DestinationFolder $FileName
    $counter = 1
    while (Test-Path $finalPath) {
        $newName = "${baseName} ($counter)${extension}"
        $finalPath = Join-Path $DestinationFolder $newName
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
    $recents = [System.Collections.ArrayList]@($global:Config.RecentFolders)
    $recents.Remove($Path)
    $recents.Insert(0, $Path)
    if ($recents.Count -gt 10) { $recents = [System.Collections.ArrayList]@($recents[0..9]) }
    $global:Config.RecentFolders = @($recents)
    Save-Config -Config $global:Config
}

function Export-MoveLog {
    param([array]$Log, [string]$FolderPath)
    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $logFile = Join-Path $global:LogFolder "organize_log_$timestamp.csv"
    $Log | Select-Object FileName, Category, OriginalPath, DestinationPath, FileSize | Export-Csv -Path $logFile -NoTypeInformation -Force
    return $logFile
}

# ========================== GLOBAL STATE ==========================
$global:SelectedPath = ""
$global:AnalysisResults = @{}
$global:TotalFiles = 0
$global:MoveLog = @()
$global:AllFiles = @()
$global:NavHistory = @()
$global:NavIndex = -1
$global:DryRunMode = [bool]$global:Config.DryRunMode
$global:RecursiveMode = [bool]$global:Config.RecursiveMode
$global:ConflictAction = $global:Config.ConflictAction

function Push-NavHistory {
    param([string]$Path)
    if ($global:NavIndex -lt $global:NavHistory.Count - 1) {
        $global:NavHistory = @($global:NavHistory[0..$global:NavIndex])
    }
    $global:NavHistory += $Path
    $global:NavIndex = $global:NavHistory.Count - 1
}

# ========================== TOOLTIP ==========================
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 10000
$toolTip.InitialDelay = 350
$toolTip.ReshowDelay = 150
$toolTip.ShowAlways = $true

# ========================== MAIN FORM ==========================
$form = New-Object System.Windows.Forms.Form
$form.Text = "File Organizer Pro"
$form.Size = New-Object System.Drawing.Size($global:Config.WindowWidth, $global:Config.WindowHeight)
$form.MinimumSize = New-Object System.Drawing.Size(820, 620)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.KeyPreview = $true
$form.AllowDrop = $true

# Save window size on close
$form.Add_FormClosing({
    $global:Config.WindowWidth = $form.Width
    $global:Config.WindowHeight = $form.Height
    $global:Config.Theme = $global:CurrentTheme
    $global:Config.DryRunMode = $global:DryRunMode
    $global:Config.RecursiveMode = $global:RecursiveMode
    $global:Config.ConflictAction = $global:ConflictAction
    Save-Config -Config $global:Config
    $global:Stats.LastUsed = (Get-Date).ToString("yyyy-MM-dd HH:mm")
    Save-Stats -Stats $global:Stats
})

# ========================== DRAG AND DROP ==========================
$form.Add_DragEnter({
    if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $paths = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        if ($paths.Count -gt 0 -and (Test-Path $paths[0] -PathType Container)) {
            $_.Effect = [System.Windows.Forms.DragDropEffects]::Link
        }
        else { $_.Effect = [System.Windows.Forms.DragDropEffects]::None }
    }
})

$form.Add_DragDrop({
    $paths = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    if ($paths.Count -gt 0 -and (Test-Path $paths[0] -PathType Container)) {
        Navigate-To -Path $paths[0]
    }
})

# ========================== MAIN TABLE LAYOUT ==========================
$mainPanel = New-Object System.Windows.Forms.TableLayoutPanel
$mainPanel.Dock = "Fill"
$mainPanel.Padding = New-Object System.Windows.Forms.Padding(12)
$mainPanel.RowCount = 9
$mainPanel.ColumnCount = 1
$mainPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 100))) | Out-Null
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 42))) | Out-Null   # 0: Header
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 32))) | Out-Null   # 1: Mode bar
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 38))) | Out-Null    # 2: Browser
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 44))) | Out-Null   # 3: Category filters
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 34))) | Out-Null    # 4: Preview
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 24))) | Out-Null   # 5: Progress
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 46))) | Out-Null   # 6: Buttons
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 24))) | Out-Null   # 7: Status bar
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 28))) | Out-Null    # 8: Hints + preview
$form.Controls.Add($mainPanel)

# ========================== ROW 0: HEADER ==========================
$headerPanel = New-Object System.Windows.Forms.TableLayoutPanel
$headerPanel.Dock = "Fill"
$headerPanel.ColumnCount = 5
$headerPanel.RowCount = 1
$headerPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 100))) | Out-Null
$headerPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 35))) | Out-Null
$headerPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 100))) | Out-Null
$headerPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 100))) | Out-Null
$headerPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 110))) | Out-Null
$headerPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 35))) | Out-Null

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "FILE ORGANIZER PRO"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$titleLabel.Dock = "Fill"
$titleLabel.TextAlign = "MiddleLeft"
$headerPanel.Controls.Add($titleLabel, 0, 0)

$btnStats = New-Object System.Windows.Forms.Button
$btnStats.Text = "Statistics"
$btnStats.Dock = "Fill"
$btnStats.FlatStyle = "Flat"
$btnStats.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$btnStats.Cursor = "Hand"
$btnStats.Margin = New-Object System.Windows.Forms.Padding(2, 6, 2, 6)
$toolTip.SetToolTip($btnStats, "View lifetime statistics: total files organized, most common types, etc. (Shortcut: Ctrl+I)")
$headerPanel.Controls.Add($btnStats, 1, 0)

$btnCustomCat = New-Object System.Windows.Forms.Button
$btnCustomCat.Text = "Categories"
$btnCustomCat.Dock = "Fill"
$btnCustomCat.FlatStyle = "Flat"
$btnCustomCat.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$btnCustomCat.Cursor = "Hand"
$btnCustomCat.Margin = New-Object System.Windows.Forms.Padding(2, 6, 2, 6)
$toolTip.SetToolTip($btnCustomCat, "Add or manage custom file categories with your own extensions (Shortcut: Ctrl+K)")
$headerPanel.Controls.Add($btnCustomCat, 2, 0)

$themeToggleBtn = New-Object System.Windows.Forms.Button
$themeToggleBtn.Text = "Light Mode"
$themeToggleBtn.Dock = "Fill"
$themeToggleBtn.FlatStyle = "Flat"
$themeToggleBtn.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$themeToggleBtn.Cursor = "Hand"
$themeToggleBtn.Margin = New-Object System.Windows.Forms.Padding(2, 6, 2, 6)
$toolTip.SetToolTip($themeToggleBtn, "Toggle between dark and light themes (Shortcut: Ctrl+T)")
$headerPanel.Controls.Add($themeToggleBtn, 3, 0)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Ready"
$statusLabel.Dock = "Fill"
$statusLabel.TextAlign = "MiddleRight"
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$headerPanel.Controls.Add($statusLabel, 4, 0)

$mainPanel.Controls.Add($headerPanel, 0, 0)

# ========================== ROW 1: MODE BAR ==========================
$modePanel = New-Object System.Windows.Forms.TableLayoutPanel
$modePanel.Dock = "Fill"
$modePanel.ColumnCount = 6
$modePanel.RowCount = 1
$modePanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 100))) | Out-Null
$modePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 110))) | Out-Null
$modePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 120))) | Out-Null
$modePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 16))) | Out-Null
$modePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 170))) | Out-Null
$modePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 110))) | Out-Null
$modePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 100))) | Out-Null

$chkDryRun = New-Object System.Windows.Forms.CheckBox
$chkDryRun.Text = "Dry Run"
$chkDryRun.Checked = $global:DryRunMode
$chkDryRun.AutoSize = $true
$chkDryRun.Dock = "Fill"
$chkDryRun.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$chkDryRun.Cursor = "Hand"
$toolTip.SetToolTip($chkDryRun, "When enabled, the organize operation will simulate without actually moving files. Use this to preview results safely. (Shortcut: Ctrl+D)")
$modePanel.Controls.Add($chkDryRun, 0, 0)

$chkRecursive = New-Object System.Windows.Forms.CheckBox
$chkRecursive.Text = "Recursive"
$chkRecursive.Checked = $global:RecursiveMode
$chkRecursive.AutoSize = $true
$chkRecursive.Dock = "Fill"
$chkRecursive.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$chkRecursive.Cursor = "Hand"
$toolTip.SetToolTip($chkRecursive, "When enabled, files in ALL subfolders will also be organized. When disabled, only top-level files are processed. (Shortcut: Ctrl+R)")
$modePanel.Controls.Add($chkRecursive, 1, 0)

$sepLabel = New-Object System.Windows.Forms.Label
$sepLabel.Text = "|"
$sepLabel.Dock = "Fill"
$sepLabel.TextAlign = "MiddleCenter"
$modePanel.Controls.Add($sepLabel, 2, 0)

$conflictLabel = New-Object System.Windows.Forms.Label
$conflictLabel.Text = "On conflict:"
$conflictLabel.Dock = "Fill"
$conflictLabel.TextAlign = "MiddleLeft"
$conflictLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$modePanel.Controls.Add($conflictLabel, 3, 0)

$conflictCombo = New-Object System.Windows.Forms.ComboBox
$conflictCombo.Dock = "Fill"
$conflictCombo.DropDownStyle = "DropDownList"
$conflictCombo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$conflictCombo.Items.AddRange(@("Rename", "Skip", "Overwrite"))
$conflictCombo.SelectedItem = $global:ConflictAction
$conflictCombo.Margin = New-Object System.Windows.Forms.Padding(0, 3, 4, 3)
$toolTip.SetToolTip($conflictCombo, "Choose what happens when a file with the same name already exists in the destination folder:`n  Rename: Appends (1), (2), etc.`n  Skip: Leaves the file in place`n  Overwrite: Replaces the existing file")
$modePanel.Controls.Add($conflictCombo, 4, 0)

$modeStatusLabel = New-Object System.Windows.Forms.Label
$modeStatusLabel.Text = ""
$modeStatusLabel.Dock = "Fill"
$modeStatusLabel.TextAlign = "MiddleRight"
$modeStatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$modePanel.Controls.Add($modeStatusLabel, 5, 0)

$mainPanel.Controls.Add($modePanel, 0, 1)

# Mode checkbox events
$chkDryRun.Add_CheckedChanged({ $global:DryRunMode = $chkDryRun.Checked; Update-ModeStatus })
$chkRecursive.Add_CheckedChanged({
    $global:RecursiveMode = $chkRecursive.Checked
    Update-ModeStatus
    if ($global:SelectedPath -ne "") { Invoke-FolderAnalysis -Path $global:SelectedPath }
})
$conflictCombo.Add_SelectedIndexChanged({ $global:ConflictAction = $conflictCombo.SelectedItem })

function Update-ModeStatus {
    $t = Get-Theme
    $parts = @()
    if ($global:DryRunMode) { $parts += "DRY RUN"; $modeStatusLabel.ForeColor = $t.DryRun }
    if ($global:RecursiveMode) { $parts += "RECURSIVE" }
    if ($parts.Count -eq 0) { $modeStatusLabel.Text = ""; return }
    $modeStatusLabel.Text = "[ $($parts -join ' | ') ]"
    if (-not $global:DryRunMode) { $modeStatusLabel.ForeColor = $t.Orange }
}

# ========================== ROW 2: FOLDER BROWSER ==========================
$browserOuterPanel = New-Object System.Windows.Forms.Panel
$browserOuterPanel.Dock = "Fill"
$browserOuterPanel.Padding = New-Object System.Windows.Forms.Padding(0, 2, 0, 2)

$browserGroupBox = New-Object System.Windows.Forms.GroupBox
$browserGroupBox.Text = "  Folder Browser -- Double-click to navigate, drag and drop a folder onto the window  "
$browserGroupBox.Dock = "Fill"
$browserGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$browserOuterPanel.Controls.Add($browserGroupBox)

$browserInnerLayout = New-Object System.Windows.Forms.TableLayoutPanel
$browserInnerLayout.Dock = "Fill"
$browserInnerLayout.Padding = New-Object System.Windows.Forms.Padding(6, 2, 6, 4)
$browserInnerLayout.RowCount = 3
$browserInnerLayout.ColumnCount = 1
$browserInnerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 100))) | Out-Null
$browserInnerLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 32))) | Out-Null
$browserInnerLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 28))) | Out-Null
$browserInnerLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 100))) | Out-Null
$browserGroupBox.Controls.Add($browserInnerLayout)

# -- Address bar --
$addressPanel = New-Object System.Windows.Forms.TableLayoutPanel
$addressPanel.Dock = "Fill"
$addressPanel.ColumnCount = 7
$addressPanel.RowCount = 1
$addressPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 100))) | Out-Null
$addressPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 40))) | Out-Null
$addressPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 40))) | Out-Null
$addressPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 42))) | Out-Null
$addressPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 55))) | Out-Null
$addressPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 68))) | Out-Null
$addressPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 100))) | Out-Null
$addressPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 45))) | Out-Null

$btnBack = New-Object System.Windows.Forms.Button
$btnBack.Text = "<"
$btnBack.Dock = "Fill"; $btnBack.FlatStyle = "Flat"
$btnBack.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnBack.Cursor = "Hand"; $btnBack.Margin = New-Object System.Windows.Forms.Padding(0,0,1,0)
$toolTip.SetToolTip($btnBack, "Go back (Alt+Left / Backspace)")
$addressPanel.Controls.Add($btnBack, 0, 0)

$btnForward = New-Object System.Windows.Forms.Button
$btnForward.Text = ">"
$btnForward.Dock = "Fill"; $btnForward.FlatStyle = "Flat"
$btnForward.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnForward.Cursor = "Hand"; $btnForward.Margin = New-Object System.Windows.Forms.Padding(0,0,1,0)
$toolTip.SetToolTip($btnForward, "Go forward (Alt+Right)")
$addressPanel.Controls.Add($btnForward, 1, 0)

$btnUp = New-Object System.Windows.Forms.Button
$btnUp.Text = "Up"
$btnUp.Dock = "Fill"; $btnUp.FlatStyle = "Flat"
$btnUp.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$btnUp.Cursor = "Hand"; $btnUp.Margin = New-Object System.Windows.Forms.Padding(0,0,1,0)
$toolTip.SetToolTip($btnUp, "Parent folder (Alt+Up)")
$addressPanel.Controls.Add($btnUp, 2, 0)

$btnHome = New-Object System.Windows.Forms.Button
$btnHome.Text = "Home"
$btnHome.Dock = "Fill"; $btnHome.FlatStyle = "Flat"
$btnHome.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$btnHome.Cursor = "Hand"; $btnHome.Margin = New-Object System.Windows.Forms.Padding(0,0,2,0)
$toolTip.SetToolTip($btnHome, "Home -- drives and quick access (Ctrl+H)")
$addressPanel.Controls.Add($btnHome, 3, 0)

$btnRecent = New-Object System.Windows.Forms.Button
$btnRecent.Text = "Recent"
$btnRecent.Dock = "Fill"; $btnRecent.FlatStyle = "Flat"
$btnRecent.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$btnRecent.Cursor = "Hand"; $btnRecent.Margin = New-Object System.Windows.Forms.Padding(0,0,3,0)
$toolTip.SetToolTip($btnRecent, "Show recently organized folders (Shortcut: Ctrl+E)")
$addressPanel.Controls.Add($btnRecent, 4, 0)

$addressBox = New-Object System.Windows.Forms.TextBox
$addressBox.Dock = "Fill"; $addressBox.BorderStyle = "FixedSingle"
$addressBox.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$addressBox.Margin = New-Object System.Windows.Forms.Padding(0,2,3,2)
$toolTip.SetToolTip($addressBox, "Type or paste a path and press Enter (Ctrl+L to focus)")
$addressPanel.Controls.Add($addressBox, 5, 0)

$btnGo = New-Object System.Windows.Forms.Button
$btnGo.Text = "Go"
$btnGo.Dock = "Fill"; $btnGo.FlatStyle = "Flat"
$btnGo.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnGo.Cursor = "Hand"
$toolTip.SetToolTip($btnGo, "Navigate to the entered path")
$addressPanel.Controls.Add($btnGo, 6, 0)

$browserInnerLayout.Controls.Add($addressPanel, 0, 0)

# -- Search --
$searchPanel = New-Object System.Windows.Forms.TableLayoutPanel
$searchPanel.Dock = "Fill"
$searchPanel.ColumnCount = 3; $searchPanel.RowCount = 1
$searchPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 100))) | Out-Null
$searchPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 50))) | Out-Null
$searchPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 100))) | Out-Null
$searchPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 50))) | Out-Null

$searchLbl = New-Object System.Windows.Forms.Label
$searchLbl.Text = "Filter:"; $searchLbl.Dock = "Fill"; $searchLbl.TextAlign = "MiddleLeft"
$searchLbl.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$searchPanel.Controls.Add($searchLbl, 0, 0)

$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Dock = "Fill"; $searchBox.BorderStyle = "FixedSingle"
$searchBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$searchBox.Margin = New-Object System.Windows.Forms.Padding(0,2,3,2)
$toolTip.SetToolTip($searchBox, "Filter items in current folder (Shortcut: / or Ctrl+F)")
$searchPanel.Controls.Add($searchBox, 1, 0)

$btnClearSearch = New-Object System.Windows.Forms.Button
$btnClearSearch.Text = "Clear"; $btnClearSearch.Dock = "Fill"; $btnClearSearch.FlatStyle = "Flat"
$btnClearSearch.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$btnClearSearch.Cursor = "Hand"
$toolTip.SetToolTip($btnClearSearch, "Clear filter (Escape)")
$searchPanel.Controls.Add($btnClearSearch, 2, 0)

$browserInnerLayout.Controls.Add($searchPanel, 0, 1)

# -- Folder ListView --
$folderListView = New-Object System.Windows.Forms.ListView
$folderListView.Dock = "Fill"; $folderListView.View = "Details"
$folderListView.FullRowSelect = $true; $folderListView.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$folderListView.BorderStyle = "None"; $folderListView.HeaderStyle = "Nonclickable"
$folderListView.GridLines = $false; $folderListView.HideSelection = $false
$folderListView.Columns.Add("Name", 280) | Out-Null
$folderListView.Columns.Add("Type", 70) | Out-Null
$folderListView.Columns.Add("Size", 90) | Out-Null
$folderListView.Columns.Add("Modified", 140) | Out-Null
$toolTip.SetToolTip($folderListView, "Double-click folder to enter. Enter key to open. Backspace to go back.")

$browserInnerLayout.Controls.Add($folderListView, 0, 2)
$mainPanel.Controls.Add($browserOuterPanel, 0, 2)

# ========================== ROW 3: FILTERS ==========================
$filterPanel = New-Object System.Windows.Forms.Panel
$filterPanel.Dock = "Fill"; $filterPanel.Padding = New-Object System.Windows.Forms.Padding(0,2,0,2)

$filterFlow = New-Object System.Windows.Forms.FlowLayoutPanel
$filterFlow.Dock = "Fill"; $filterFlow.WrapContents = $true; $filterFlow.AutoScroll = $true
$filterFlow.Padding = New-Object System.Windows.Forms.Padding(6,2,6,2)
$toolTip.SetToolTip($filterFlow, "Check categories to include. Uncheck to skip. Hover each for its extensions.")

$chkAll = New-Object System.Windows.Forms.CheckBox
$chkAll.Text = "All"; $chkAll.Checked = $true; $chkAll.AutoSize = $true
$chkAll.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$chkAll.Margin = New-Object System.Windows.Forms.Padding(6,4,10,4); $chkAll.Cursor = "Hand"
$toolTip.SetToolTip($chkAll, "Select/deselect all categories (Ctrl+A)")
$filterFlow.Controls.Add($chkAll)

$global:CategoryCheckboxes = @{}
foreach ($cat in $global:FileCategories.Keys) {
    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Text = $cat; $chk.Checked = $true; $chk.AutoSize = $true
    $chk.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $chk.Margin = New-Object System.Windows.Forms.Padding(3,4,3,4); $chk.Cursor = "Hand"
    $chk.Tag = $cat
    $extList = ($global:FileCategories[$cat] | ForEach-Object { $_ }) -join "  "
    $toolTip.SetToolTip($chk, "$cat files`nExtensions: $extList")
    $filterFlow.Controls.Add($chk)
    $global:CategoryCheckboxes[$cat] = $chk
}

$chkOther = New-Object System.Windows.Forms.CheckBox
$chkOther.Text = "Other"; $chkOther.Checked = $true; $chkOther.AutoSize = $true
$chkOther.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$chkOther.Margin = New-Object System.Windows.Forms.Padding(3,4,3,4); $chkOther.Cursor = "Hand"
$toolTip.SetToolTip($chkOther, "Files not matching any defined category")
$filterFlow.Controls.Add($chkOther)
$global:CategoryCheckboxes["Other"] = $chkOther

$filterPanel.Controls.Add($filterFlow)
$mainPanel.Controls.Add($filterPanel, 0, 3)

$global:UpdatingCheckboxes = $false
$chkAll.Add_CheckedChanged({
    if ($global:UpdatingCheckboxes) { return }
    $global:UpdatingCheckboxes = $true
    foreach ($chk in $global:CategoryCheckboxes.Values) { $chk.Checked = $chkAll.Checked }
    $global:UpdatingCheckboxes = $false
    if ($global:SelectedPath -ne "") { Invoke-FolderAnalysis -Path $global:SelectedPath }
})

foreach ($chk in $global:CategoryCheckboxes.Values) {
    $chk.Add_CheckedChanged({
        if ($global:UpdatingCheckboxes) { return }
        $global:UpdatingCheckboxes = $true
        $allOn = $true
        foreach ($c in $global:CategoryCheckboxes.Values) { if (-not $c.Checked) { $allOn = $false; break } }
        $chkAll.Checked = $allOn
        $global:UpdatingCheckboxes = $false
        if ($global:SelectedPath -ne "") { Invoke-FolderAnalysis -Path $global:SelectedPath }
    })
}

# ========================== ROW 4: PREVIEW ==========================
$previewOuterPanel = New-Object System.Windows.Forms.Panel
$previewOuterPanel.Dock = "Fill"; $previewOuterPanel.Padding = New-Object System.Windows.Forms.Padding(0,2,0,2)

$previewGroupBox = New-Object System.Windows.Forms.GroupBox
$previewGroupBox.Text = "  Analysis Preview  "
$previewGroupBox.Dock = "Fill"; $previewGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$previewOuterPanel.Controls.Add($previewGroupBox)

$previewBox = New-Object System.Windows.Forms.RichTextBox
$previewBox.Dock = "Fill"; $previewBox.ReadOnly = $true; $previewBox.BorderStyle = "None"
$previewBox.Font = New-Object System.Drawing.Font("Cascadia Code, Consolas, Courier New", 9)
$previewBox.Text = "Select a folder to see the analysis..."
$toolTip.SetToolTip($previewBox, "Detailed breakdown of files grouped by category with sizes and percentages")
$previewGroupBox.Controls.Add($previewBox)
$mainPanel.Controls.Add($previewOuterPanel, 0, 4)

# ========================== ROW 5: PROGRESS ==========================
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Dock = "Fill"; $progressBar.Style = "Continuous"
$progressBar.Minimum = 0; $progressBar.Value = 0
$progressBar.Margin = New-Object System.Windows.Forms.Padding(0,2,0,2)
$toolTip.SetToolTip($progressBar, "Progress of current operation")
$mainPanel.Controls.Add($progressBar, 0, 5)

# ========================== ROW 6: BUTTONS ==========================
$buttonPanel = New-Object System.Windows.Forms.TableLayoutPanel
$buttonPanel.Dock = "Fill"; $buttonPanel.ColumnCount = 6; $buttonPanel.RowCount = 1
$buttonPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 100))) | Out-Null
$buttonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 30))) | Out-Null
$buttonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 18))) | Out-Null
$buttonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 14))) | Out-Null
$buttonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 14))) | Out-Null
$buttonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 12))) | Out-Null
$buttonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 12))) | Out-Null

$organizeBtn = New-Object System.Windows.Forms.Button
$organizeBtn.Text = "ORGANIZE FILES"; $organizeBtn.Dock = "Fill"; $organizeBtn.FlatStyle = "Flat"
$organizeBtn.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$organizeBtn.FlatAppearance.BorderSize = 0; $organizeBtn.Cursor = "Hand"; $organizeBtn.Enabled = $false
$organizeBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,2,0)
$toolTip.SetToolTip($organizeBtn, "Move matching files into categorized subfolders (Ctrl+Enter)")
$buttonPanel.Controls.Add($organizeBtn, 0, 0)

$undoLastBtn = New-Object System.Windows.Forms.Button
$undoLastBtn.Text = "Undo Category"; $undoLastBtn.Dock = "Fill"; $undoLastBtn.FlatStyle = "Flat"
$undoLastBtn.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$undoLastBtn.FlatAppearance.BorderSize = 0; $undoLastBtn.Cursor = "Hand"; $undoLastBtn.Enabled = $false
$undoLastBtn.Margin = New-Object System.Windows.Forms.Padding(2,0,2,0)
$toolTip.SetToolTip($undoLastBtn, "Undo a specific category (Ctrl+U)")
$buttonPanel.Controls.Add($undoLastBtn, 1, 0)

$undoAllBtn = New-Object System.Windows.Forms.Button
$undoAllBtn.Text = "Undo All"; $undoAllBtn.Dock = "Fill"; $undoAllBtn.FlatStyle = "Flat"
$undoAllBtn.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$undoAllBtn.FlatAppearance.BorderSize = 0; $undoAllBtn.Cursor = "Hand"; $undoAllBtn.Enabled = $false
$undoAllBtn.Margin = New-Object System.Windows.Forms.Padding(2,0,2,0)
$toolTip.SetToolTip($undoAllBtn, "Reverse ALL moves (Ctrl+Z)")
$buttonPanel.Controls.Add($undoAllBtn, 2, 0)

$exportLogBtn = New-Object System.Windows.Forms.Button
$exportLogBtn.Text = "Export Log"; $exportLogBtn.Dock = "Fill"; $exportLogBtn.FlatStyle = "Flat"
$exportLogBtn.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$exportLogBtn.Cursor = "Hand"; $exportLogBtn.Enabled = $false
$exportLogBtn.Margin = New-Object System.Windows.Forms.Padding(2,0,2,0)
$toolTip.SetToolTip($exportLogBtn, "Save a CSV log of all file moves (Ctrl+S)")
$buttonPanel.Controls.Add($exportLogBtn, 3, 0)

$refreshBtn = New-Object System.Windows.Forms.Button
$refreshBtn.Text = "Refresh"; $refreshBtn.Dock = "Fill"; $refreshBtn.FlatStyle = "Flat"
$refreshBtn.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$refreshBtn.Cursor = "Hand"; $refreshBtn.Margin = New-Object System.Windows.Forms.Padding(2,0,2,0)
$toolTip.SetToolTip($refreshBtn, "Reload folder and re-analyze (F5)")
$buttonPanel.Controls.Add($refreshBtn, 4, 0)

$openFolderBtn = New-Object System.Windows.Forms.Button
$openFolderBtn.Text = "Open"; $openFolderBtn.Dock = "Fill"; $openFolderBtn.FlatStyle = "Flat"
$openFolderBtn.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$openFolderBtn.Cursor = "Hand"; $openFolderBtn.Margin = New-Object System.Windows.Forms.Padding(2,0,0,0)
$toolTip.SetToolTip($openFolderBtn, "Open in Windows Explorer (Ctrl+O)")
$buttonPanel.Controls.Add($openFolderBtn, 5, 0)

$mainPanel.Controls.Add($buttonPanel, 0, 6)

# ========================== ROW 7: STATUS BAR ==========================
$statusBar = New-Object System.Windows.Forms.Label
$statusBar.Dock = "Fill"; $statusBar.TextAlign = "MiddleLeft"
$statusBar.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$statusBar.Text = "Items: 0  |  Ready  |  Drag and drop a folder to get started"
$mainPanel.Controls.Add($statusBar, 0, 7)

# ========================== ROW 8: HINTS + FILE PREVIEW ==========================
$bottomSplit = New-Object System.Windows.Forms.TableLayoutPanel
$bottomSplit.Dock = "Fill"; $bottomSplit.ColumnCount = 2; $bottomSplit.RowCount = 1
$bottomSplit.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 100))) | Out-Null
$bottomSplit.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 55))) | Out-Null
$bottomSplit.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 45))) | Out-Null

$hintsGroupBox = New-Object System.Windows.Forms.GroupBox
$hintsGroupBox.Text = "  Keyboard Shortcuts and Tips  "
$hintsGroupBox.Dock = "Fill"; $hintsGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)

$hintsBox = New-Object System.Windows.Forms.RichTextBox
$hintsBox.Dock = "Fill"; $hintsBox.ReadOnly = $true; $hintsBox.BorderStyle = "None"
$hintsBox.Font = New-Object System.Drawing.Font("Cascadia Code, Consolas, Courier New", 8)
$hintsBox.ScrollBars = "Vertical"
$hintsGroupBox.Controls.Add($hintsBox)
$bottomSplit.Controls.Add($hintsGroupBox, 0, 0)

$filePreviewGroupBox = New-Object System.Windows.Forms.GroupBox
$filePreviewGroupBox.Text = "  File Preview  "
$filePreviewGroupBox.Dock = "Fill"; $filePreviewGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)

$filePreviewPanel = New-Object System.Windows.Forms.Panel
$filePreviewPanel.Dock = "Fill"

$filePreviewPicture = New-Object System.Windows.Forms.PictureBox
$filePreviewPicture.Dock = "Fill"; $filePreviewPicture.SizeMode = "Zoom"
$filePreviewPicture.Visible = $false
$filePreviewPanel.Controls.Add($filePreviewPicture)

$filePreviewText = New-Object System.Windows.Forms.RichTextBox
$filePreviewText.Dock = "Fill"; $filePreviewText.ReadOnly = $true; $filePreviewText.BorderStyle = "None"
$filePreviewText.Font = New-Object System.Drawing.Font("Cascadia Code, Consolas, Courier New", 8.5)
$filePreviewText.ScrollBars = "Vertical"
$filePreviewText.WordWrap = $true
$filePreviewText.Visible = $true
$filePreviewText.Text = "Select a file from the browser to preview its content here.`n`nSupported preview types:`n  - Images: thumbnail preview`n  - Text/Code/Config files: content preview`n  - Other files: file information"
$filePreviewPanel.Controls.Add($filePreviewText)

$filePreviewGroupBox.Controls.Add($filePreviewPanel)
$bottomSplit.Controls.Add($filePreviewGroupBox, 1, 0)

$mainPanel.Controls.Add($bottomSplit, 0, 8)

# ========================== FILE PREVIEW LOGIC ==========================
$global:ImageExtensions = @(".jpg", ".jpeg", ".png", ".gif", ".bmp", ".ico", ".tiff", ".tif", ".webp")
$global:TextExtensions = @(".txt", ".md", ".log", ".csv", ".json", ".xml", ".yaml", ".yml", ".ini", ".cfg", ".conf", ".toml", ".py", ".js", ".ts", ".html", ".htm", ".css", ".java", ".cpp", ".c", ".cs", ".ps1", ".sh", ".bash", ".sql", ".php", ".rb", ".go", ".rs", ".swift", ".kt", ".r", ".lua", ".pl", ".bat", ".cmd", ".rtf")

function Show-FilePreview {
    param([string]$FilePath)
    $t = Get-Theme

    if (-not (Test-Path $FilePath) -or (Test-Path $FilePath -PathType Container)) {
        $filePreviewPicture.Visible = $false
        $filePreviewText.Visible = $true
        $filePreviewText.Clear()
        $filePreviewText.SelectionColor = $t.TextDim
        $filePreviewText.AppendText("Select a file to preview.")
        return
    }

    $ext = [System.IO.Path]::GetExtension($FilePath).ToLower()
    $fileInfo = Get-Item $FilePath -ErrorAction SilentlyContinue

    if ($global:ImageExtensions -contains $ext) {
        try {
            $filePreviewText.Visible = $false
            $filePreviewPicture.Visible = $true
            $stream = [System.IO.File]::OpenRead($FilePath)
            $img = [System.Drawing.Image]::FromStream($stream)
            if ($filePreviewPicture.Image) { $filePreviewPicture.Image.Dispose() }
            $filePreviewPicture.Image = $img.Clone()
            $img.Dispose()
            $stream.Close()
            $stream.Dispose()
        }
        catch {
            $filePreviewPicture.Visible = $false
            $filePreviewText.Visible = $true
            $filePreviewText.Clear()
            $filePreviewText.SelectionColor = $t.Red
            $filePreviewText.AppendText("Cannot preview this image file.")
        }
    }
    elseif ($global:TextExtensions -contains $ext) {
        $filePreviewPicture.Visible = $false
        $filePreviewText.Visible = $true
        $filePreviewText.Clear()
        try {
            $maxChars = 8000
            $content = [System.IO.File]::ReadAllText($FilePath)
            if ($content.Length -gt $maxChars) {
                $content = $content.Substring(0, $maxChars)
                $content += "`n`n--- Preview truncated at $maxChars characters ---"
            }
            $filePreviewText.SelectionColor = $t.TextPrimary
            $filePreviewText.AppendText($content)
            $filePreviewText.SelectionStart = 0
            $filePreviewText.ScrollToCaret()
        }
        catch {
            $filePreviewText.SelectionColor = $t.Red
            $filePreviewText.AppendText("Cannot read this file.")
        }
    }
    else {
        $filePreviewPicture.Visible = $false
        $filePreviewText.Visible = $true
        $filePreviewText.Clear()
        $filePreviewText.SelectionColor = $t.Accent
        $filePreviewText.AppendText("FILE INFORMATION`n")
        $filePreviewText.SelectionColor = $t.TextMuted
        $filePreviewText.AppendText(([string]::new([char]0x2500, 35)) + "`n`n")
        if ($fileInfo) {
            $cat = Get-CategoryForFile -Extension $ext
            $filePreviewText.SelectionColor = $t.TextPrimary
            $filePreviewText.AppendText("  Name:      $($fileInfo.Name)`n")
            $filePreviewText.AppendText("  Type:      $($ext.TrimStart('.').ToUpper())`n")
            $filePreviewText.AppendText("  Category:  $cat`n")
            $filePreviewText.AppendText("  Size:      $(Get-FolderSizeText $fileInfo.Length)`n")
            $filePreviewText.AppendText("  Created:   $($fileInfo.CreationTime.ToString('yyyy-MM-dd HH:mm:ss'))`n")
            $filePreviewText.AppendText("  Modified:  $($fileInfo.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))`n")
            $filePreviewText.AppendText("  Accessed:  $($fileInfo.LastAccessTime.ToString('yyyy-MM-dd HH:mm:ss'))`n")
            $filePreviewText.AppendText("  ReadOnly:  $($fileInfo.IsReadOnly)`n")
            $filePreviewText.AppendText("  Path:      $($fileInfo.DirectoryName)`n")
        }
    }
}

# ========================== FOLDER LIST FUNCTIONS ==========================
function Update-FolderList {
    param([string]$Path, [string]$Filter = "")
    $t = Get-Theme
    $folderListView.Items.Clear()
    if (-not (Test-Path $Path)) { return }
    try {
        $items = Get-ChildItem -Path $Path -ErrorAction Stop | Sort-Object { -not $_.PSIsContainer }, Name
        if ($Filter -ne "") { $items = $items | Where-Object { $_.Name -like "*$Filter*" } }
        foreach ($item in $items) {
            $lvItem = New-Object System.Windows.Forms.ListViewItem($item.Name)
            if ($item.PSIsContainer) {
                $lvItem.SubItems.Add("Folder") | Out-Null
                $lvItem.SubItems.Add("") | Out-Null
                $lvItem.ForeColor = $t.FolderText
            }
            else {
                $ext = $item.Extension.TrimStart(".")
                if ($ext -eq "") { $ext = "File" }
                $lvItem.SubItems.Add($ext.ToUpper()) | Out-Null
                $lvItem.SubItems.Add((Get-FolderSizeText $item.Length)) | Out-Null
                $lvItem.ForeColor = $t.TextPrimary
            }
            $lvItem.SubItems.Add($item.LastWriteTime.ToString("yyyy-MM-dd  HH:mm")) | Out-Null
            $lvItem.Tag = $item.FullName
            $folderListView.Items.Add($lvItem) | Out-Null
        }
        $fc = ($items | Where-Object { -not $_.PSIsContainer }).Count
        $dc = ($items | Where-Object { $_.PSIsContainer }).Count
        $statusBar.Text = "Items: $($items.Count)  |  Folders: $dc  |  Files: $fc  |  $Path"
    }
    catch {
        $statusLabel.Text = "Access denied: $Path"
        $statusLabel.ForeColor = $t.Red
    }
}

function Navigate-To {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return }
    $t = Get-Theme
    $global:SelectedPath = $Path
    $addressBox.Text = $Path
    $searchBox.Text = ""
    Push-NavHistory -Path $Path
    Update-FolderList -Path $Path
    Invoke-FolderAnalysis -Path $Path
    $statusLabel.Text = "Browsing: $Path"
    $statusLabel.ForeColor = $t.TextDim
}

function Show-DriveList {
    $folderListView.Items.Clear()
    $t = Get-Theme
    $drives = [System.IO.DriveInfo]::GetDrives() | Where-Object { $_.IsReady }
    foreach ($drive in $drives) {
        $label = if ($drive.VolumeLabel -ne "") { "$($drive.Name)  [$($drive.VolumeLabel)]" } else { $drive.Name }
        $lvItem = New-Object System.Windows.Forms.ListViewItem($label)
        $lvItem.SubItems.Add("Drive") | Out-Null
        $lvItem.SubItems.Add("$(Get-FolderSizeText ($drive.TotalSize - $drive.AvailableFreeSpace)) / $(Get-FolderSizeText $drive.TotalSize)") | Out-Null
        $lvItem.SubItems.Add($drive.DriveType.ToString()) | Out-Null
        $lvItem.Tag = $drive.Name
        $lvItem.ForeColor = $t.DriveText
        $folderListView.Items.Add($lvItem) | Out-Null
    }
    $quickPaths = @(
        @([Environment]::GetFolderPath("Desktop"), "Desktop"),
        @([Environment]::GetFolderPath("MyDocuments"), "Documents"),
        @([Environment]::GetFolderPath("UserProfile") + "\Downloads", "Downloads"),
        @([Environment]::GetFolderPath("MyPictures"), "Pictures"),
        @([Environment]::GetFolderPath("MyMusic"), "Music"),
        @([Environment]::GetFolderPath("MyVideos"), "Videos")
    )
    foreach ($qp in $quickPaths) {
        if (Test-Path $qp[0]) {
            $lvItem = New-Object System.Windows.Forms.ListViewItem("[ $($qp[1]) ]")
            $lvItem.SubItems.Add("Quick") | Out-Null
            $lvItem.SubItems.Add("") | Out-Null
            $lvItem.SubItems.Add("") | Out-Null
            $lvItem.Tag = $qp[0]
            $lvItem.ForeColor = $t.QuickText
            $folderListView.Items.Add($lvItem) | Out-Null
        }
    }
    $addressBox.Text = "My Computer"
    $global:SelectedPath = ""
    $previewBox.Clear()
    $previewBox.SelectionColor = $t.TextDim
    $previewBox.AppendText("Select a folder to analyze.`nDrag and drop a folder, use quick access, or type a path.")
    $organizeBtn.Enabled = $false
    $statusBar.Text = "Home  |  $($drives.Count) drive(s)  |  Drag and drop a folder to get started"
}

# ========================== ANALYSIS ==========================
function Invoke-FolderAnalysis {
    param([string]$Path)
    $t = Get-Theme
    $global:AnalysisResults = @{}
    $global:TotalFiles = 0
    $global:AllFiles = @()

    if ([string]::IsNullOrEmpty($Path) -or -not (Test-Path $Path)) {
        $previewBox.Clear()
        $previewBox.SelectionColor = $t.TextDim
        $previewBox.AppendText("Select a folder to analyze.")
        $organizeBtn.Enabled = $false
        return
    }

    $gciParams = @{ Path = $Path; File = $true; ErrorAction = "SilentlyContinue" }
    if ($global:RecursiveMode) { $gciParams["Recurse"] = $true }
    $files = Get-ChildItem @gciParams

    if (-not $files -or $files.Count -eq 0) {
        $previewBox.Clear()
        $previewBox.SelectionColor = $t.TextDim
        $previewBox.AppendText("No files found$(if ($global:RecursiveMode) { ' (recursive mode)' }).`nNavigate to a folder with files.")
        $organizeBtn.Enabled = $false
        return
    }

    $selectedCategories = @()
    foreach ($entry in $global:CategoryCheckboxes.GetEnumerator()) {
        if ($entry.Value.Checked) { $selectedCategories += $entry.Key }
    }

    if ($selectedCategories.Count -eq 0) {
        $previewBox.Clear()
        $previewBox.SelectionColor = $t.Orange
        $previewBox.AppendText("No categories selected. Check at least one.")
        $organizeBtn.Enabled = $false
        return
    }

    $totalSize = [long]0
    foreach ($file in $files) {
        $category = Get-CategoryForFile -Extension $file.Extension
        if ($selectedCategories -contains $category) {
            if (-not $global:AnalysisResults.ContainsKey($category)) { $global:AnalysisResults[$category] = @() }
            $global:AnalysisResults[$category] += $file
            $global:AllFiles += $file
            $global:TotalFiles++
            $totalSize += $file.Length
        }
    }

    $previewBox.Clear()
    if ($global:TotalFiles -eq 0) {
        $previewBox.SelectionColor = $t.TextDim
        $previewBox.AppendText("No matching files for selected categories.")
        $organizeBtn.Enabled = $false
        return
    }

    $previewBox.SelectionColor = $t.Success
    $previewBox.AppendText("ANALYSIS RESULTS$(if ($global:DryRunMode) { '  [DRY RUN MODE]' })`n")
    $previewBox.SelectionColor = $t.TextMuted
    $previewBox.AppendText(([string]::new([char]0x2500, 58)) + "`n`n")
    $previewBox.SelectionColor = $t.TextPrimary
    $previewBox.AppendText("  Path:          $Path`n")
    $previewBox.AppendText("  Mode:          $(if ($global:RecursiveMode) { 'Recursive (all subfolders)' } else { 'Top-level only' })`n")
    $previewBox.AppendText("  Total files:   $($global:TotalFiles)`n")
    $previewBox.AppendText("  Total size:    $(Get-FolderSizeText $totalSize)`n")
    $previewBox.AppendText("  Categories:    $($global:AnalysisResults.Count)`n")
    $previewBox.AppendText("  Conflict:      $($global:ConflictAction)`n")
    $skipped = $files.Count - $global:TotalFiles
    if ($skipped -gt 0) {
        $previewBox.SelectionColor = $t.TextMuted
        $previewBox.AppendText("  Skipped:       $skipped (unchecked categories)`n")
    }
    $previewBox.AppendText("`n")
    $previewBox.SelectionColor = $t.TextMuted
    $previewBox.AppendText(([string]::new([char]0x2500, 58)) + "`n`n")

    $sorted = $global:AnalysisResults.GetEnumerator() | Sort-Object { $_.Value.Count } -Descending
    foreach ($cat in $sorted) {
        $catName = $cat.Key
        $catFiles = $cat.Value
        $count = $catFiles.Count
        $catSize = ($catFiles | Measure-Object -Property Length -Sum).Sum
        $pct = [math]::Round(($count / $global:TotalFiles) * 100, 1)
        $barLen = [math]::Floor($pct / 5)
        $bar = ([string]::new([char]0x2588, $barLen)).PadRight(20)

        $previewBox.SelectionColor = $t.FolderText
        $previewBox.AppendText("  [$catName]`n")
        $previewBox.SelectionColor = $t.TextDim
        $previewBox.AppendText("    $count file(s)  |  $(Get-FolderSizeText $catSize)  |  $pct%`n")
        $previewBox.SelectionColor = $t.Success
        $previewBox.AppendText("    $bar`n")

        $show = [Math]::Min(3, $catFiles.Count)
        for ($i = 0; $i -lt $show; $i++) {
            $previewBox.SelectionColor = $t.TextMuted
            $previewBox.AppendText("      $($catFiles[$i].Name)  ($(Get-FolderSizeText $catFiles[$i].Length))`n")
        }
        if ($catFiles.Count -gt 3) {
            $previewBox.SelectionColor = [System.Drawing.Color]::FromArgb(80, 80, 90)
            $previewBox.AppendText("      ... and $($catFiles.Count - 3) more`n")
        }
        $previewBox.AppendText("`n")
    }

    $previewBox.SelectionColor = $t.TextMuted
    $previewBox.AppendText(([string]::new([char]0x2500, 58)) + "`n")
    $previewBox.SelectionColor = $t.TextDim
    $previewBox.AppendText("  Press ORGANIZE FILES (Ctrl+Enter) to proceed.`n")
    if ($global:DryRunMode) {
        $previewBox.SelectionColor = $t.DryRun
        $previewBox.AppendText("  DRY RUN is ON -- no files will actually be moved.`n")
    }

    $organizeBtn.Enabled = $true
    $statusLabel.Text = "Analysis  |  $($global:TotalFiles) files  |  $(Get-FolderSizeText $totalSize)"
    $statusLabel.ForeColor = $t.Success
}

# ========================== HINTS PANEL ==========================
function Update-HintsPanel {
    $t = Get-Theme
    $hintsBox.Clear()
    $shortcuts = @(
        @("Ctrl+Enter",    "Organize files"),
        @("Ctrl+Z",        "Undo all moves"),
        @("Ctrl+U",        "Undo specific category"),
        @("Ctrl+S",        "Export move log to CSV"),
        @("Ctrl+L",        "Focus address bar"),
        @("Ctrl+F  or /",  "Focus search filter"),
        @("Ctrl+H",        "Home view"),
        @("Ctrl+T",        "Toggle theme"),
        @("Ctrl+O",        "Open in Explorer"),
        @("Ctrl+A",        "Toggle all categories"),
        @("Ctrl+D",        "Toggle Dry Run mode"),
        @("Ctrl+R",        "Toggle Recursive mode"),
        @("Ctrl+I",        "View statistics"),
        @("Ctrl+K",        "Custom categories"),
        @("Ctrl+E",        "Recent folders"),
        @("F5",            "Refresh"),
        @("Alt+Left",      "Navigate back"),
        @("Alt+Right",     "Navigate forward"),
        @("Alt+Up",        "Parent folder"),
        @("Backspace",     "Go back (folder list)"),
        @("Enter",         "Open folder (folder list)"),
        @("Escape",        "Clear search")
    )
    $hintsBox.SelectionColor = $t.Accent
    $hintsBox.AppendText("KEYBOARD SHORTCUTS`n")
    $hintsBox.SelectionColor = $t.TextMuted
    $hintsBox.AppendText(([string]::new([char]0x2500, 45)) + "`n")
    foreach ($s in $shortcuts) {
        $hintsBox.SelectionColor = $t.Success
        $hintsBox.AppendText("  $($s[0].PadRight(16))")
        $hintsBox.SelectionColor = $t.TextDim
        $hintsBox.AppendText("$($s[1])`n")
    }
    $hintsBox.SelectionColor = $t.TextMuted
    $hintsBox.AppendText("`n" + ([string]::new([char]0x2500, 45)) + "`n")
    $hintsBox.SelectionColor = $t.Accent
    $hintsBox.AppendText("TIPS`n")
    $hintsBox.SelectionColor = $t.TextMuted
    $hintsBox.AppendText(([string]::new([char]0x2500, 45)) + "`n")
    $tips = @(
        "Drag and drop a folder onto the window.",
        "Hover any button for a description.",
        "Use Dry Run to simulate safely.",
        "Recursive mode organizes subfolders too.",
        "Conflict: Rename / Skip / Overwrite.",
        "Create custom categories for your workflow.",
        "Export Log saves a CSV of all moves.",
        "Statistics track your lifetime usage.",
        "Recent folders remember your history.",
        "Click a file to preview its content.",
        "Files are never deleted, only moved.",
        "Empty folders cleaned up on undo."
    )
    foreach ($tip in $tips) {
        $hintsBox.SelectionColor = $t.TextDim
        $hintsBox.AppendText("  * $tip`n")
    }
    $hintsBox.SelectionStart = 0; $hintsBox.ScrollToCaret()
}

# ========================== APPLY THEME ==========================
function Apply-Theme {
    $t = Get-Theme
    $form.BackColor = $t.Bg; $form.ForeColor = $t.TextPrimary
    $mainPanel.BackColor = $t.Bg; $headerPanel.BackColor = $t.Bg
    $titleLabel.ForeColor = $t.Accent
    $themeToggleBtn.BackColor = $t.Input; $themeToggleBtn.ForeColor = $t.TextPrimary
    $themeToggleBtn.FlatAppearance.BorderColor = $t.Border; $themeToggleBtn.FlatAppearance.BorderSize = 1
    $themeToggleBtn.Text = if ($global:CurrentTheme -eq "Dark") { "Light Mode" } else { "Dark Mode" }
    $btnStats.BackColor = $t.Input; $btnStats.ForeColor = $t.TextPrimary
    $btnStats.FlatAppearance.BorderColor = $t.Border; $btnStats.FlatAppearance.BorderSize = 1
    $btnCustomCat.BackColor = $t.Input; $btnCustomCat.ForeColor = $t.TextPrimary
    $btnCustomCat.FlatAppearance.BorderColor = $t.Border; $btnCustomCat.FlatAppearance.BorderSize = 1
    $statusLabel.ForeColor = $t.TextDim
    $modePanel.BackColor = $t.Panel
    $chkDryRun.ForeColor = $t.TextPrimary; $chkDryRun.BackColor = $t.Panel
    $chkRecursive.ForeColor = $t.TextPrimary; $chkRecursive.BackColor = $t.Panel
    $sepLabel.ForeColor = $t.TextMuted; $sepLabel.BackColor = $t.Panel
    $conflictLabel.ForeColor = $t.TextDim; $conflictLabel.BackColor = $t.Panel
    $conflictCombo.BackColor = $t.Input; $conflictCombo.ForeColor = $t.TextPrimary
    $modeStatusLabel.BackColor = $t.Panel
    $browserGroupBox.BackColor = $t.Panel; $browserGroupBox.ForeColor = $t.TextDim
    $browserInnerLayout.BackColor = $t.Panel
    foreach ($btn in @($btnBack, $btnForward, $btnUp, $btnHome, $btnRecent)) {
        $btn.BackColor = $t.Input; $btn.ForeColor = $t.TextPrimary
        $btn.FlatAppearance.BorderColor = $t.Border; $btn.FlatAppearance.BorderSize = 1
    }
    $addressBox.BackColor = $t.Input; $addressBox.ForeColor = $t.TextPrimary
    $btnGo.BackColor = $t.Accent; $btnGo.ForeColor = [System.Drawing.Color]::White; $btnGo.FlatAppearance.BorderSize = 0
    $searchLbl.ForeColor = $t.TextDim; $searchLbl.BackColor = $t.Panel
    $searchBox.BackColor = $t.Input; $searchBox.ForeColor = $t.TextPrimary
    $btnClearSearch.BackColor = $t.Input; $btnClearSearch.ForeColor = $t.TextDim
    $btnClearSearch.FlatAppearance.BorderColor = $t.Border; $btnClearSearch.FlatAppearance.BorderSize = 1
    $folderListView.BackColor = $t.ListBg; $folderListView.ForeColor = $t.TextPrimary
    $filterPanel.BackColor = $t.Bg; $filterFlow.BackColor = $t.Panel
    $chkAll.ForeColor = $t.Accent; $chkAll.BackColor = $t.Panel
    foreach ($chk in $global:CategoryCheckboxes.Values) { $chk.ForeColor = $t.TextPrimary; $chk.BackColor = $t.Panel }
    $previewGroupBox.BackColor = $t.Panel; $previewGroupBox.ForeColor = $t.TextDim
    $previewBox.BackColor = $t.ListBg; $previewBox.ForeColor = $t.TextPrimary
    $organizeBtn.BackColor = $t.Green; $organizeBtn.ForeColor = [System.Drawing.Color]::White
    $undoLastBtn.BackColor = $t.Orange; $undoLastBtn.ForeColor = [System.Drawing.Color]::White
    $undoAllBtn.BackColor = $t.Red; $undoAllBtn.ForeColor = [System.Drawing.Color]::White
    $exportLogBtn.BackColor = $t.Input; $exportLogBtn.ForeColor = $t.TextPrimary
    $exportLogBtn.FlatAppearance.BorderColor = $t.Border; $exportLogBtn.FlatAppearance.BorderSize = 1
    $refreshBtn.BackColor = $t.Input; $refreshBtn.ForeColor = $t.TextPrimary
    $refreshBtn.FlatAppearance.BorderColor = $t.Border; $refreshBtn.FlatAppearance.BorderSize = 1
    $openFolderBtn.BackColor = $t.Input; $openFolderBtn.ForeColor = $t.TextPrimary
    $openFolderBtn.FlatAppearance.BorderColor = $t.Border; $openFolderBtn.FlatAppearance.BorderSize = 1
    $statusBar.BackColor = $t.Panel; $statusBar.ForeColor = $t.TextMuted
    $bottomSplit.BackColor = $t.Bg
    $hintsGroupBox.BackColor = $t.Panel; $hintsGroupBox.ForeColor = $t.TextDim
    $hintsBox.BackColor = $t.HintBg; $hintsBox.ForeColor = $t.TextPrimary
    $filePreviewGroupBox.BackColor = $t.Panel; $filePreviewGroupBox.ForeColor = $t.TextDim
    $filePreviewPanel.BackColor = $t.ListBg
    $filePreviewPicture.BackColor = $t.ListBg
    $filePreviewText.BackColor = $t.ListBg; $filePreviewText.ForeColor = $t.TextPrimary
    Update-HintsPanel
    Update-ModeStatus
}

# ========================== NAV EVENTS ==========================
$folderListView.Add_DoubleClick({
    $sel = $folderListView.SelectedItems
    if ($sel.Count -eq 0) { return }
    $path = $sel[0].Tag
    if (Test-Path $path -PathType Container) { Navigate-To -Path $path }
})

$folderListView.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        $sel = $folderListView.SelectedItems
        if ($sel.Count -eq 0) { return }
        $path = $sel[0].Tag
        if (Test-Path $path -PathType Container) { Navigate-To -Path $path }
    }
    elseif ($_.KeyCode -eq "Back") { $btnBack.PerformClick() }
})

$folderListView.Add_SelectedIndexChanged({
    $t = Get-Theme
    if ($folderListView.SelectedItems.Count -gt 0) {
        $sel = $folderListView.SelectedItems[0]
        $statusBar.Text = "Selected: $($sel.Text)  |  $($sel.SubItems[1].Text)  |  $($sel.SubItems[2].Text)  |  $($sel.SubItems[3].Text)"
        $filePath = $sel.Tag
        if ($filePath -and (Test-Path $filePath -PathType Leaf)) {
            Show-FilePreview -FilePath $filePath
        }
        elseif ($filePath -and (Test-Path $filePath -PathType Container)) {
            $filePreviewPicture.Visible = $false; $filePreviewText.Visible = $true
            $filePreviewText.Clear()
            $filePreviewText.SelectionColor = $t.FolderText
            $filePreviewText.AppendText("FOLDER`n")
            $filePreviewText.SelectionColor = $t.TextMuted
            $filePreviewText.AppendText(([string]::new([char]0x2500, 30)) + "`n`n")
            try {
                $info = Get-Item $filePath
                $childCount = (Get-ChildItem $filePath -ErrorAction SilentlyContinue).Count
                $filePreviewText.SelectionColor = $t.TextPrimary
                $filePreviewText.AppendText("  Name:     $($info.Name)`n")
                $filePreviewText.AppendText("  Items:    $childCount`n")
                $filePreviewText.AppendText("  Created:  $($info.CreationTime.ToString('yyyy-MM-dd HH:mm'))`n")
                $filePreviewText.AppendText("  Modified: $($info.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))`n")
            }
            catch {}
        }
    }
})

$btnBack.Add_Click({
    if ($global:NavIndex -gt 0) {
        $global:NavIndex--
        $path = $global:NavHistory[$global:NavIndex]
        $global:SelectedPath = $path; $addressBox.Text = $path; $searchBox.Text = ""
        Update-FolderList -Path $path; Invoke-FolderAnalysis -Path $path
    }
})

$btnForward.Add_Click({
    if ($global:NavIndex -lt $global:NavHistory.Count - 1) {
        $global:NavIndex++
        $path = $global:NavHistory[$global:NavIndex]
        $global:SelectedPath = $path; $addressBox.Text = $path; $searchBox.Text = ""
        Update-FolderList -Path $path; Invoke-FolderAnalysis -Path $path
    }
})

$btnUp.Add_Click({
    if ([string]::IsNullOrEmpty($global:SelectedPath)) { Show-DriveList; return }
    $parent = Split-Path $global:SelectedPath -Parent
    if ([string]::IsNullOrEmpty($parent)) { Show-DriveList } else { Navigate-To -Path $parent }
})

$btnHome.Add_Click({ Show-DriveList })

$btnRecent.Add_Click({
    $t = Get-Theme
    $folderListView.Items.Clear()
    $global:SelectedPath = ""
    $addressBox.Text = "Recent Folders"
    if ($global:Config.RecentFolders.Count -eq 0) {
        $previewBox.Clear(); $previewBox.SelectionColor = $t.TextDim
        $previewBox.AppendText("No recent folders yet.`nOrganize a folder and it will appear here.")
        return
    }
    foreach ($path in $global:Config.RecentFolders) {
        if (Test-Path $path) {
            $lvItem = New-Object System.Windows.Forms.ListViewItem($path)
            $lvItem.SubItems.Add("Recent") | Out-Null
            $lvItem.SubItems.Add("") | Out-Null
            $lvItem.SubItems.Add("") | Out-Null
            $lvItem.Tag = $path
            $lvItem.ForeColor = $t.QuickText
            $folderListView.Items.Add($lvItem) | Out-Null
        }
    }
    $statusBar.Text = "Recent Folders  |  $($global:Config.RecentFolders.Count) entries"
})

$btnGo.Add_Click({
    $t = Get-Theme; $path = $addressBox.Text.Trim()
    if (Test-Path $path -PathType Container) { Navigate-To -Path $path }
    else { $statusLabel.Text = "Invalid path"; $statusLabel.ForeColor = $t.Red }
})

$addressBox.Add_KeyDown({ if ($_.KeyCode -eq "Enter") { $btnGo.PerformClick(); $_.SuppressKeyPress = $true } })

$searchBox.Add_TextChanged({
    if ($global:SelectedPath -ne "" -and (Test-Path $global:SelectedPath)) {
        Update-FolderList -Path $global:SelectedPath -Filter $searchBox.Text
    }
})

$btnClearSearch.Add_Click({ $searchBox.Text = ""; $searchBox.Focus() })

$refreshBtn.Add_Click({
    if ($global:SelectedPath -ne "" -and (Test-Path $global:SelectedPath)) { Navigate-To -Path $global:SelectedPath }
    else { Show-DriveList }
})

$openFolderBtn.Add_Click({
    if ($global:SelectedPath -ne "" -and (Test-Path $global:SelectedPath)) {
        Start-Process "explorer.exe" -ArgumentList $global:SelectedPath
    }
})

# ========================== THEME TOGGLE ==========================
$themeToggleBtn.Add_Click({
    $global:CurrentTheme = if ($global:CurrentTheme -eq "Dark") { "Light" } else { "Dark" }
    Apply-Theme
    if ($global:SelectedPath -ne "" -and (Test-Path $global:SelectedPath)) {
        Update-FolderList -Path $global:SelectedPath -Filter $searchBox.Text
        Invoke-FolderAnalysis -Path $global:SelectedPath
    }
    else { Show-DriveList }
})

# ========================== STATISTICS DIALOG ==========================
$btnStats.Add_Click({
    $t = Get-Theme
    $sf = New-Object System.Windows.Forms.Form
    $sf.Text = "Statistics Dashboard"; $sf.Size = New-Object System.Drawing.Size(520, 500)
    $sf.StartPosition = "CenterParent"; $sf.BackColor = $t.Bg; $sf.ForeColor = $t.TextPrimary
    $sf.FormBorderStyle = "FixedDialog"; $sf.MaximizeBox = $false; $sf.MinimizeBox = $false
    $sf.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $sr = New-Object System.Windows.Forms.RichTextBox
    $sr.Dock = "Fill"; $sr.ReadOnly = $true; $sr.BorderStyle = "None"
    $sr.BackColor = $t.ListBg; $sr.ForeColor = $t.TextPrimary
    $sr.Font = New-Object System.Drawing.Font("Cascadia Code, Consolas", 9.5)
    $sr.Padding = New-Object System.Windows.Forms.Padding(12)

    $sr.SelectionColor = $t.Accent
    $sr.AppendText("LIFETIME STATISTICS`n")
    $sr.SelectionColor = $t.TextMuted
    $sr.AppendText(([string]::new([char]0x2500, 45)) + "`n`n")

    $sr.SelectionColor = $t.TextPrimary
    $sr.AppendText("  Total files organized:   $($global:Stats.TotalFilesOrganized)`n")
    $sr.AppendText("  Total data moved:        $(Get-FolderSizeText $global:Stats.TotalSizeMoved)`n")
    $sr.AppendText("  Total operations:        $($global:Stats.TotalOperations)`n")
    $sr.AppendText("  First used:              $($global:Stats.FirstUsed)`n")
    $sr.AppendText("  Last used:               $($global:Stats.LastUsed)`n`n")

    $sr.SelectionColor = $t.TextMuted
    $sr.AppendText(([string]::new([char]0x2500, 45)) + "`n")
    $sr.SelectionColor = $t.Accent
    $sr.AppendText("FILES PER CATEGORY`n")
    $sr.SelectionColor = $t.TextMuted
    $sr.AppendText(([string]::new([char]0x2500, 45)) + "`n`n")

    if ($global:Stats.CategoryCounts -and $global:Stats.CategoryCounts.Count -gt 0) {
        $sorted = $global:Stats.CategoryCounts.GetEnumerator() | Sort-Object Value -Descending
        foreach ($entry in $sorted) {
            $sr.SelectionColor = $t.FolderText
            $sr.AppendText("  $($entry.Key.PadRight(18))")
            $sr.SelectionColor = $t.TextDim
            $sr.AppendText("$($entry.Value) files`n")
        }
    }
    else { $sr.SelectionColor = $t.TextDim; $sr.AppendText("  No data yet.`n") }

    $sr.AppendText("`n")
    $sr.SelectionColor = $t.TextMuted
    $sr.AppendText(([string]::new([char]0x2500, 45)) + "`n")
    $sr.SelectionColor = $t.Accent
    $sr.AppendText("TOP FILE EXTENSIONS`n")
    $sr.SelectionColor = $t.TextMuted
    $sr.AppendText(([string]::new([char]0x2500, 45)) + "`n`n")

    if ($global:Stats.ExtensionCounts -and $global:Stats.ExtensionCounts.Count -gt 0) {
        $sortedExt = $global:Stats.ExtensionCounts.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 15
        foreach ($entry in $sortedExt) {
            $sr.SelectionColor = $t.Success
            $sr.AppendText("  $($entry.Key.PadRight(12))")
            $sr.SelectionColor = $t.TextDim
            $sr.AppendText("$($entry.Value) files`n")
        }
    }
    else { $sr.SelectionColor = $t.TextDim; $sr.AppendText("  No data yet.`n") }

    $sf.Controls.Add($sr)
    $sf.ShowDialog()
})

# ========================== CUSTOM CATEGORIES DIALOG ==========================
$btnCustomCat.Add_Click({
    $t = Get-Theme
    $cf = New-Object System.Windows.Forms.Form
    $cf.Text = "Custom Categories Manager"; $cf.Size = New-Object System.Drawing.Size(560, 480)
    $cf.StartPosition = "CenterParent"; $cf.BackColor = $t.Bg; $cf.ForeColor = $t.TextPrimary
    $cf.FormBorderStyle = "FixedDialog"; $cf.MaximizeBox = $false; $cf.MinimizeBox = $false
    $cf.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $cfLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $cfLayout.Dock = "Fill"; $cfLayout.Padding = New-Object System.Windows.Forms.Padding(12)
    $cfLayout.RowCount = 5; $cfLayout.ColumnCount = 1
    $cfLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 100))) | Out-Null
    $cfLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 30))) | Out-Null
    $cfLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 100))) | Out-Null
    $cfLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 32))) | Out-Null
    $cfLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 32))) | Out-Null
    $cfLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 40))) | Out-Null
    $cf.Controls.Add($cfLayout)

    $cfInfoLbl = New-Object System.Windows.Forms.Label
    $cfInfoLbl.Text = "Add custom categories. Extensions must start with a dot (e.g., .dat .sav)"
    $cfInfoLbl.Dock = "Fill"; $cfInfoLbl.ForeColor = $t.TextDim; $cfInfoLbl.TextAlign = "MiddleLeft"
    $cfLayout.Controls.Add($cfInfoLbl, 0, 0)

    $cfList = New-Object System.Windows.Forms.ListView
    $cfList.Dock = "Fill"; $cfList.View = "Details"; $cfList.FullRowSelect = $true
    $cfList.BackColor = $t.ListBg; $cfList.ForeColor = $t.TextPrimary; $cfList.BorderStyle = "FixedSingle"
    $cfList.Columns.Add("Category", 150) | Out-Null
    $cfList.Columns.Add("Extensions", 350) | Out-Null
    $cfList.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # Show existing custom categories
    if ($global:Config.CustomCategories -and $global:Config.CustomCategories.Count -gt 0) {
        foreach ($entry in $global:Config.CustomCategories.GetEnumerator()) {
            $item = New-Object System.Windows.Forms.ListViewItem($entry.Key)
            $item.SubItems.Add(($entry.Value -join "  ")) | Out-Null
            $cfList.Items.Add($item) | Out-Null
        }
    }
    $cfLayout.Controls.Add($cfList, 0, 1)

    $cfNamePanel = New-Object System.Windows.Forms.TableLayoutPanel
    $cfNamePanel.Dock = "Fill"; $cfNamePanel.ColumnCount = 2; $cfNamePanel.RowCount = 1
    $cfNamePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 110))) | Out-Null
    $cfNamePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 100))) | Out-Null

    $cfNameLbl = New-Object System.Windows.Forms.Label
    $cfNameLbl.Text = "Category name:"; $cfNameLbl.Dock = "Fill"; $cfNameLbl.TextAlign = "MiddleLeft"
    $cfNameLbl.ForeColor = $t.TextDim
    $cfNamePanel.Controls.Add($cfNameLbl, 0, 0)

    $cfNameBox = New-Object System.Windows.Forms.TextBox
    $cfNameBox.Dock = "Fill"; $cfNameBox.BackColor = $t.Input; $cfNameBox.ForeColor = $t.TextPrimary
    $cfNameBox.BorderStyle = "FixedSingle"; $cfNameBox.Margin = New-Object System.Windows.Forms.Padding(0,3,0,3)
    $cfNamePanel.Controls.Add($cfNameBox, 1, 0)

    $cfLayout.Controls.Add($cfNamePanel, 0, 2)

    $cfExtPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $cfExtPanel.Dock = "Fill"; $cfExtPanel.ColumnCount = 2; $cfExtPanel.RowCount = 1
    $cfExtPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Absolute", 110))) | Out-Null
    $cfExtPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 100))) | Out-Null

    $cfExtLbl = New-Object System.Windows.Forms.Label
    $cfExtLbl.Text = "Extensions:"; $cfExtLbl.Dock = "Fill"; $cfExtLbl.TextAlign = "MiddleLeft"
    $cfExtLbl.ForeColor = $t.TextDim
    $cfExtPanel.Controls.Add($cfExtLbl, 0, 0)

    $cfExtBox = New-Object System.Windows.Forms.TextBox
    $cfExtBox.Dock = "Fill"; $cfExtBox.BackColor = $t.Input; $cfExtBox.ForeColor = $t.TextPrimary
    $cfExtBox.BorderStyle = "FixedSingle"; $cfExtBox.Margin = New-Object System.Windows.Forms.Padding(0,3,0,3)
    $cfExtPanel.Controls.Add($cfExtBox, 1, 0)

    $cfLayout.Controls.Add($cfExtPanel, 0, 3)

    $cfBtnPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $cfBtnPanel.Dock = "Fill"; $cfBtnPanel.ColumnCount = 2; $cfBtnPanel.RowCount = 1
    $cfBtnPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 50))) | Out-Null
    $cfBtnPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 50))) | Out-Null

    $cfAddBtn = New-Object System.Windows.Forms.Button
    $cfAddBtn.Text = "Add Category"; $cfAddBtn.Dock = "Fill"; $cfAddBtn.FlatStyle = "Flat"
    $cfAddBtn.BackColor = $t.Green; $cfAddBtn.ForeColor = [System.Drawing.Color]::White
    $cfAddBtn.FlatAppearance.BorderSize = 0; $cfAddBtn.Cursor = "Hand"
    $cfAddBtn.Margin = New-Object System.Windows.Forms.Padding(0,2,4,2)
    $cfBtnPanel.Controls.Add($cfAddBtn, 0, 0)

    $cfRemoveBtn = New-Object System.Windows.Forms.Button
    $cfRemoveBtn.Text = "Remove Selected"; $cfRemoveBtn.Dock = "Fill"; $cfRemoveBtn.FlatStyle = "Flat"
    $cfRemoveBtn.BackColor = $t.Red; $cfRemoveBtn.ForeColor = [System.Drawing.Color]::White
    $cfRemoveBtn.FlatAppearance.BorderSize = 0; $cfRemoveBtn.Cursor = "Hand"
    $cfRemoveBtn.Margin = New-Object System.Windows.Forms.Padding(4,2,0,2)
    $cfBtnPanel.Controls.Add($cfRemoveBtn, 1, 0)

    $cfLayout.Controls.Add($cfBtnPanel, 0, 4)

    $cfAddBtn.Add_Click({
        $name = $cfNameBox.Text.Trim()
        $exts = $cfExtBox.Text.Trim()
        if ($name -eq "" -or $exts -eq "") {
            [System.Windows.Forms.MessageBox]::Show("Enter both a category name and extensions.", "Missing Input", "OK", "Warning")
            return
        }
        $extArray = @($exts -split '\s+|,|;' | Where-Object { $_ -ne "" } | ForEach-Object {
            $e = $_.Trim().ToLower()
            if (-not $e.StartsWith(".")) { $e = ".$e" }
            $e
        })
        if ($extArray.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Enter valid extensions.", "Invalid", "OK", "Warning")
            return
        }

        # Add to config
        if (-not $global:Config.CustomCategories) { $global:Config.CustomCategories = @{} }
        $global:Config.CustomCategories[$name] = $extArray
        Save-Config -Config $global:Config

        # Add to runtime categories
        $global:FileCategories[$name] = $extArray

        # Add checkbox if not exists
        if (-not $global:CategoryCheckboxes.ContainsKey($name)) {
            $chk = New-Object System.Windows.Forms.CheckBox
            $chk.Text = $name; $chk.Checked = $true; $chk.AutoSize = $true
            $chk.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
            $chk.Margin = New-Object System.Windows.Forms.Padding(3,4,3,4); $chk.Cursor = "Hand"
            $chk.ForeColor = $t.TextPrimary; $chk.BackColor = $t.Panel
            $toolTip.SetToolTip($chk, "$name files`nExtensions: $($extArray -join '  ')")
            $chk.Add_CheckedChanged({
                if ($global:UpdatingCheckboxes) { return }
                $global:UpdatingCheckboxes = $true
                $allOn = $true
                foreach ($c in $global:CategoryCheckboxes.Values) { if (-not $c.Checked) { $allOn = $false; break } }
                $chkAll.Checked = $allOn
                $global:UpdatingCheckboxes = $false
                if ($global:SelectedPath -ne "") { Invoke-FolderAnalysis -Path $global:SelectedPath }
            })
            $filterFlow.Controls.Add($chk)
            $global:CategoryCheckboxes[$name] = $chk
        }

        # Update list
        $item = New-Object System.Windows.Forms.ListViewItem($name)
        $item.SubItems.Add(($extArray -join "  ")) | Out-Null
        $cfList.Items.Add($item) | Out-Null

        $cfNameBox.Text = ""; $cfExtBox.Text = ""

        if ($global:SelectedPath -ne "") { Invoke-FolderAnalysis -Path $global:SelectedPath }
    })

    $cfRemoveBtn.Add_Click({
        if ($cfList.SelectedItems.Count -eq 0) { return }
        $selName = $cfList.SelectedItems[0].Text

        $global:Config.CustomCategories.Remove($selName)
        Save-Config -Config $global:Config

        $global:FileCategories.Remove($selName)

        if ($global:CategoryCheckboxes.ContainsKey($selName)) {
            $filterFlow.Controls.Remove($global:CategoryCheckboxes[$selName])
            $global:CategoryCheckboxes.Remove($selName)
        }

        $cfList.Items.Remove($cfList.SelectedItems[0])

        if ($global:SelectedPath -ne "") { Invoke-FolderAnalysis -Path $global:SelectedPath }
    })

    $cf.ShowDialog()
})

# ========================== EXPORT LOG ==========================
$exportLogBtn.Add_Click({
    $t = Get-Theme
    if ($global:MoveLog.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No moves to export.", "Export Log", "OK", "Information")
        return
    }
    $logFile = Export-MoveLog -Log $global:MoveLog -FolderPath $global:SelectedPath
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Log exported to:`n$logFile`n`nOpen the file?", "Export Complete",
        [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information
    )
    if ($result -eq "Yes") { Start-Process $logFile }
})

# ========================== ORGANIZE ==========================
$organizeBtn.Add_Click({
    $t = Get-Theme
    if ([string]::IsNullOrEmpty($global:SelectedPath) -or $global:TotalFiles -eq 0) { return }

    $selectedCategories = @()
    foreach ($entry in $global:CategoryCheckboxes.GetEnumerator()) {
        if ($entry.Value.Checked) { $selectedCategories += $entry.Key }
    }

    $modeText = if ($global:DryRunMode) { "`n`nDRY RUN MODE: No files will actually be moved." } else { "" }
    $recurText = if ($global:RecursiveMode) { " (recursive)" } else { "" }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "This will $(if ($global:DryRunMode) { 'simulate organizing' } else { 'move' }) $($global:TotalFiles) file(s)$recurText.`n`nFolder: $($global:SelectedPath)`nConflict action: $($global:ConflictAction)$modeText`n`nProceed?",
        "Confirm Organization",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($confirm -ne "Yes") { return }

    $organizeBtn.Enabled = $false; $undoLastBtn.Enabled = $false; $undoAllBtn.Enabled = $false; $exportLogBtn.Enabled = $false
    $global:MoveLog = @()
    $progressBar.Maximum = $global:TotalFiles; $progressBar.Value = 0
    $moved = 0; $skipped = 0; $errors = 0

    foreach ($catEntry in $global:AnalysisResults.GetEnumerator()) {
        $category = $catEntry.Key
        $catFiles = $catEntry.Value
        $destFolder = Join-Path $global:SelectedPath $category

        if (-not $global:DryRunMode -and -not (Test-Path $destFolder)) {
            New-Item -Path $destFolder -ItemType Directory -Force | Out-Null
        }

        foreach ($file in $catFiles) {
            $destPath = Join-Path $destFolder $file.Name

            if ((Test-Path $destPath) -and -not $global:DryRunMode) {
                switch ($global:ConflictAction) {
                    "Skip" {
                        $skipped++
                        $progressBar.Value = [Math]::Min($progressBar.Maximum, $moved + $skipped + $errors)
                        $statusLabel.Text = "Organizing... ($($moved + $skipped + $errors) / $($global:TotalFiles))"
                        $form.Refresh()
                        continue
                    }
                    "Rename" {
                        $destPath = Get-UniqueFileName -DestinationFolder $destFolder -FileName $file.Name
                    }
                    "Overwrite" {
                        # destPath stays the same, -Force will handle it
                    }
                }
            }
            elseif (-not $global:DryRunMode) {
                # No conflict
            }

            if ($global:DryRunMode) {
                $global:MoveLog += [PSCustomObject]@{
                    OriginalPath    = $file.FullName
                    DestinationPath = $destPath
                    Category        = $category
                    FileName        = $file.Name
                    FileSize        = $file.Length
                }
                $moved++
            }
            else {
                try {
                    Move-Item -Path $file.FullName -Destination $destPath -Force -ErrorAction Stop
                    $global:MoveLog += [PSCustomObject]@{
                        OriginalPath    = $file.FullName
                        DestinationPath = $destPath
                        Category        = $category
                        FileName        = $file.Name
                        FileSize        = $file.Length
                    }
                    $moved++
                }
                catch { $errors++ }
            }

            $progressBar.Value = [Math]::Min($progressBar.Maximum, $moved + $skipped + $errors)
            $statusLabel.Text = "$(if ($global:DryRunMode) { 'Simulating' } else { 'Organizing' })... ($($moved + $skipped + $errors) / $($global:TotalFiles))"
            $statusLabel.ForeColor = if ($global:DryRunMode) { $t.DryRun } else { $t.Orange }
            $form.Refresh()
        }
    }

    $progressBar.Value = $progressBar.Maximum

    # Update stats
    if (-not $global:DryRunMode) {
        $global:Stats.TotalFilesOrganized += $moved
        $global:Stats.TotalSizeMoved += ($global:MoveLog | Measure-Object -Property FileSize -Sum).Sum
        $global:Stats.TotalOperations++
        foreach ($entry in $global:MoveLog) {
            $catKey = $entry.Category
            if (-not $global:Stats.CategoryCounts.ContainsKey($catKey)) { $global:Stats.CategoryCounts[$catKey] = 0 }
            $global:Stats.CategoryCounts[$catKey] = [int]$global:Stats.CategoryCounts[$catKey] + 1
            $extKey = [System.IO.Path]::GetExtension($entry.FileName).ToLower()
            if ($extKey -ne "") {
                if (-not $global:Stats.ExtensionCounts.ContainsKey($extKey)) { $global:Stats.ExtensionCounts[$extKey] = 0 }
                $global:Stats.ExtensionCounts[$extKey] = [int]$global:Stats.ExtensionCounts[$extKey] + 1
            }
        }
        Save-Stats -Stats $global:Stats
        Add-RecentFolder -Path $global:SelectedPath
    }

    # Summary
    $totalMovedSize = ($global:MoveLog | Measure-Object -Property FileSize -Sum).Sum
    $previewBox.Clear()
    $previewBox.SelectionColor = if ($global:DryRunMode) { $t.DryRun } else { $t.Success }
    $previewBox.AppendText("$(if ($global:DryRunMode) { 'DRY RUN SIMULATION COMPLETE' } else { 'ORGANIZATION COMPLETE' })`n")
    $previewBox.SelectionColor = $t.TextMuted
    $previewBox.AppendText(([string]::new([char]0x2500, 58)) + "`n`n")
    $previewBox.SelectionColor = $t.TextPrimary
    $previewBox.AppendText("  Folder:        $($global:SelectedPath)`n")
    $previewBox.AppendText("  Files moved:   $moved`n")
    if ($skipped -gt 0) { $previewBox.SelectionColor = $t.Orange; $previewBox.AppendText("  Skipped:       $skipped (conflicts)`n") }
    if ($errors -gt 0) { $previewBox.SelectionColor = $t.Red; $previewBox.AppendText("  Errors:        $errors`n") }
    $previewBox.SelectionColor = $t.TextDim
    $previewBox.AppendText("  Data size:     $(Get-FolderSizeText $totalMovedSize)`n`n")

    $createdFolders = $global:MoveLog | Group-Object Category | Sort-Object Count -Descending
    $previewBox.SelectionColor = $t.FolderText
    $previewBox.AppendText("  Folders$(if ($global:DryRunMode) { ' (simulated)' }):`n")
    foreach ($group in $createdFolders) {
        $grpSize = ($group.Group | Measure-Object -Property FileSize -Sum).Sum
        $previewBox.SelectionColor = $t.TextDim
        $previewBox.AppendText("    [$($group.Name)]  $($group.Count) file(s)  |  $(Get-FolderSizeText $grpSize)`n")
    }
    $previewBox.AppendText("`n")
    $previewBox.SelectionColor = $t.TextMuted
    if ($global:DryRunMode) {
        $previewBox.AppendText("  This was a dry run. No files were moved.`n")
        $previewBox.AppendText("  Disable Dry Run and run again to move files.`n")
    }
    else {
        $previewBox.AppendText("  Undo Category (Ctrl+U) | Undo All (Ctrl+Z) | Export Log (Ctrl+S)`n")
    }

    $statusLabel.Text = "$(if ($global:DryRunMode) { 'Simulation' } else { 'Done' })  |  Moved: $moved  |  Skipped: $skipped  |  Errors: $errors"
    $statusLabel.ForeColor = if ($global:DryRunMode) { $t.DryRun } else { $t.Success }

    if (-not $global:DryRunMode) {
        Update-FolderList -Path $global:SelectedPath
    }

    if ($global:MoveLog.Count -gt 0 -and -not $global:DryRunMode) {
        $undoLastBtn.Enabled = $true; $undoAllBtn.Enabled = $true; $exportLogBtn.Enabled = $true
    }
    elseif ($global:DryRunMode) {
        $exportLogBtn.Enabled = $true
    }
})

# ========================== UNDO CATEGORY ==========================
$undoLastBtn.Add_Click({
    $t = Get-Theme
    if ($global:MoveLog.Count -eq 0) { return }
    $categories = $global:MoveLog | Select-Object -ExpandProperty Category -Unique

    $uf = New-Object System.Windows.Forms.Form
    $uf.Text = "Select Category to Undo"; $uf.Size = New-Object System.Drawing.Size(440, 380)
    $uf.StartPosition = "CenterParent"; $uf.BackColor = $t.Bg; $uf.ForeColor = $t.TextPrimary
    $uf.FormBorderStyle = "FixedDialog"; $uf.MaximizeBox = $false; $uf.MinimizeBox = $false
    $uf.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $uLbl = New-Object System.Windows.Forms.Label
    $uLbl.Text = "Select a category to undo. Files will be moved back."
    $uLbl.Location = New-Object System.Drawing.Point(16, 16)
    $uLbl.Size = New-Object System.Drawing.Size(390, 35); $uLbl.ForeColor = $t.TextDim
    $uf.Controls.Add($uLbl)

    $uList = New-Object System.Windows.Forms.ListBox
    $uList.Location = New-Object System.Drawing.Point(16, 55)
    $uList.Size = New-Object System.Drawing.Size(390, 210)
    $uList.BackColor = $t.Input; $uList.ForeColor = $t.TextPrimary
    $uList.BorderStyle = "FixedSingle"; $uList.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    foreach ($cat in $categories) {
        $ce = $global:MoveLog | Where-Object { $_.Category -eq $cat }
        $cs = ($ce | Measure-Object -Property FileSize -Sum).Sum
        $uList.Items.Add("$cat  --  $($ce.Count) file(s)  |  $(Get-FolderSizeText $cs)") | Out-Null
    }
    if ($uList.Items.Count -gt 0) { $uList.SelectedIndex = 0 }
    $uf.Controls.Add($uList)

    $uBtn = New-Object System.Windows.Forms.Button
    $uBtn.Text = "Undo Selected Category"
    $uBtn.Location = New-Object System.Drawing.Point(16, 278)
    $uBtn.Size = New-Object System.Drawing.Size(390, 40)
    $uBtn.FlatStyle = "Flat"; $uBtn.BackColor = $t.Orange; $uBtn.ForeColor = [System.Drawing.Color]::White
    $uBtn.FlatAppearance.BorderSize = 0; $uBtn.Cursor = "Hand"
    $uBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $uf.Controls.Add($uBtn)

    $uBtn.Add_Click({
        if ($uList.SelectedIndex -lt 0) { return }
        $selCat = $uList.SelectedItem.ToString().Split("  ")[0].Trim()
        $entries = $global:MoveLog | Where-Object { $_.Category -eq $selCat }
        $restored = 0; $ue = 0
        foreach ($entry in $entries) {
            try {
                if (Test-Path $entry.DestinationPath) {
                    Move-Item -Path $entry.DestinationPath -Destination $entry.OriginalPath -Force -ErrorAction Stop
                    $restored++
                }
            }
            catch { $ue++ }
        }
        $catFolder = Join-Path $global:SelectedPath $selCat
        if ((Test-Path $catFolder) -and ((Get-ChildItem $catFolder).Count -eq 0)) {
            Remove-Item $catFolder -Force -ErrorAction SilentlyContinue
        }
        $global:MoveLog = @($global:MoveLog | Where-Object { $_.Category -ne $selCat })
        [System.Windows.Forms.MessageBox]::Show("Restored $restored file(s) from '$selCat'.$(if ($ue -gt 0) { " Errors: $ue" })", "Undo Complete", "OK", "Information")
        $uf.Close()
    })

    $uf.ShowDialog()
    if ($global:MoveLog.Count -eq 0) { $undoLastBtn.Enabled = $false; $undoAllBtn.Enabled = $false; $exportLogBtn.Enabled = $false }
    if ($global:SelectedPath -ne "" -and (Test-Path $global:SelectedPath)) {
        Update-FolderList -Path $global:SelectedPath; Invoke-FolderAnalysis -Path $global:SelectedPath
    }
})

# ========================== UNDO ALL ==========================
$undoAllBtn.Add_Click({
    $t = Get-Theme
    if ($global:MoveLog.Count -eq 0) { return }
    $totalUndoSize = ($global:MoveLog | Measure-Object -Property FileSize -Sum).Sum
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Restore all $($global:MoveLog.Count) file(s)?`nData: $(Get-FolderSizeText $totalUndoSize)", "Confirm Undo All",
        [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($confirm -ne "Yes") { return }

    $undoAllBtn.Enabled = $false; $undoLastBtn.Enabled = $false; $exportLogBtn.Enabled = $false
    $restored = 0; $ue = 0
    $progressBar.Maximum = $global:MoveLog.Count; $progressBar.Value = 0

    foreach ($entry in $global:MoveLog) {
        try {
            if (Test-Path $entry.DestinationPath) {
                Move-Item -Path $entry.DestinationPath -Destination $entry.OriginalPath -Force -ErrorAction Stop
                $restored++
            }
        }
        catch { $ue++ }
        $progressBar.Value = [Math]::Min($progressBar.Maximum, $restored + $ue)
        $statusLabel.Text = "Undoing... ($($restored + $ue) / $($global:MoveLog.Count))"
        $statusLabel.ForeColor = $t.Orange
        $form.Refresh()
    }

    foreach ($category in ($global:MoveLog | Select-Object -ExpandProperty Category -Unique)) {
        $catFolder = Join-Path $global:SelectedPath $category
        if ((Test-Path $catFolder) -and ((Get-ChildItem $catFolder).Count -eq 0)) {
            Remove-Item $catFolder -Force -ErrorAction SilentlyContinue
        }
    }

    $statusLabel.Text = "Undo complete  |  Restored: $restored  |  Errors: $ue"
    $statusLabel.ForeColor = $t.Success
    $global:MoveLog = @()

    if ($global:SelectedPath -ne "" -and (Test-Path $global:SelectedPath)) {
        Update-FolderList -Path $global:SelectedPath; Invoke-FolderAnalysis -Path $global:SelectedPath
    }
})

# ========================== KEYBOARD SHORTCUTS ==========================
$form.Add_KeyDown({
    $t = Get-Theme
    if ($_.Control -and $_.KeyCode -eq "Return") {
        if ($organizeBtn.Enabled) { $organizeBtn.PerformClick() }; $_.Handled = $true
    }
    elseif ($_.Control -and $_.KeyCode -eq "Z") {
        if ($undoAllBtn.Enabled) { $undoAllBtn.PerformClick() }; $_.Handled = $true
    }
    elseif ($_.Control -and $_.KeyCode -eq "U") {
        if ($undoLastBtn.Enabled) { $undoLastBtn.PerformClick() }; $_.Handled = $true
    }
    elseif ($_.Control -and $_.KeyCode -eq "S") {
        if ($exportLogBtn.Enabled) { $exportLogBtn.PerformClick() }; $_.Handled = $true; $_.SuppressKeyPress = $true
    }
    elseif ($_.Control -and $_.KeyCode -eq "L") {
        $addressBox.Focus(); $addressBox.SelectAll(); $_.Handled = $true; $_.SuppressKeyPress = $true
    }
    elseif ($_.Control -and $_.KeyCode -eq "F") {
        $searchBox.Focus(); $searchBox.SelectAll(); $_.Handled = $true; $_.SuppressKeyPress = $true
    }
    elseif ($_.KeyCode -eq "OemQuestion" -and -not $_.Shift -and -not $_.Control) {
        if ($form.ActiveControl -isnot [System.Windows.Forms.TextBox] -and $form.ActiveControl -isnot [System.Windows.Forms.ComboBox]) {
            $searchBox.Focus(); $searchBox.Text = ""; $_.Handled = $true; $_.SuppressKeyPress = $true
        }
    }
    elseif ($_.Control -and $_.KeyCode -eq "H") {
        Show-DriveList; $_.Handled = $true; $_.SuppressKeyPress = $true
    }
    elseif ($_.Control -and $_.KeyCode -eq "T") {
        $themeToggleBtn.PerformClick(); $_.Handled = $true; $_.SuppressKeyPress = $true
    }
    elseif ($_.Control -and $_.KeyCode -eq "O") {
        $openFolderBtn.PerformClick(); $_.Handled = $true; $_.SuppressKeyPress = $true
    }
    elseif ($_.Control -and $_.KeyCode -eq "A") {
        if ($form.ActiveControl -isnot [System.Windows.Forms.TextBox]) {
            $chkAll.Checked = -not $chkAll.Checked; $_.Handled = $true; $_.SuppressKeyPress = $true
        }
    }
    elseif ($_.Control -and $_.KeyCode -eq "D") {
        $chkDryRun.Checked = -not $chkDryRun.Checked
        if ($global:SelectedPath -ne "") { Invoke-FolderAnalysis -Path $global:SelectedPath }
        $_.Handled = $true; $_.SuppressKeyPress = $true
    }
    elseif ($_.Control -and $_.KeyCode -eq "R") {
        $chkRecursive.Checked = -not $chkRecursive.Checked
        $_.Handled = $true; $_.SuppressKeyPress = $true
    }
    elseif ($_.Control -and $_.KeyCode -eq "I") {
        $btnStats.PerformClick(); $_.Handled = $true; $_.SuppressKeyPress = $true
    }
    elseif ($_.Control -and $_.KeyCode -eq "K") {
        $btnCustomCat.PerformClick(); $_.Handled = $true; $_.SuppressKeyPress = $true
    }
    elseif ($_.Control -and $_.KeyCode -eq "E") {
        $btnRecent.PerformClick(); $_.Handled = $true; $_.SuppressKeyPress = $true
    }
    elseif ($_.KeyCode -eq "F5") {
        $refreshBtn.PerformClick(); $_.Handled = $true
    }
    elseif ($_.Alt -and $_.KeyCode -eq "Left") {
        $btnBack.PerformClick(); $_.Handled = $true
    }
    elseif ($_.Alt -and $_.KeyCode -eq "Right") {
        $btnForward.PerformClick(); $_.Handled = $true
    }
    elseif ($_.Alt -and $_.KeyCode -eq "Up") {
        $btnUp.PerformClick(); $_.Handled = $true
    }
    elseif ($_.KeyCode -eq "Escape") {
        $searchBox.Text = ""; $folderListView.Focus(); $_.Handled = $true
    }
})

# ========================== INITIALIZE ==========================
Apply-Theme
Show-DriveList

[void]$form.ShowDialog()
