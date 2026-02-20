Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Global Logic ---
$script:History = New-Object System.Collections.Generic.Stack[Object]
$script:fileObjects = @()
$script:CurrentPath = ""
$script:PreviewMode = $true

# --- Theme Colors ---
$bgDark      = [System.Drawing.Color]::FromArgb(30, 30, 30)
$panelDark   = [System.Drawing.Color]::FromArgb(45, 45, 48)
$textWhite   = [System.Drawing.Color]::White
$accentBlue  = [System.Drawing.Color]::FromArgb(0, 122, 204)
$btnGray     = [System.Drawing.Color]::FromArgb(60, 60, 60)
$hintYellow  = [System.Drawing.Color]::FromArgb(255, 185, 0)
$successGreen = [System.Drawing.Color]::FromArgb(16, 124, 16)
$warningOrange = [System.Drawing.Color]::FromArgb(202, 81, 0)

# --- Main Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Bulk Rename Studio Pro - Advanced Edition"
$form.Size = New-Object System.Drawing.Size(1500, 950)
$form.BackColor = $bgDark
$form.ForeColor = $textWhite
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(1200, 700)

# --- Status Bar ---
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusBar.BackColor = $panelDark
$statusBar.ForeColor = $textWhite

$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready | Select a folder to begin"
$statusLabel.Spring = $true
$statusLabel.TextAlign = "MiddleLeft"

$statusCount = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusCount.Text = "Files: 0"

$statusBar.Items.AddRange(@($statusLabel, $statusCount))
$form.Controls.Add($statusBar)

# --- 1. MAIN LAYOUT (Splitter) ---
$splitMain = New-Object System.Windows.Forms.SplitContainer
$splitMain.Dock = "Fill"
$splitMain.SplitterDistance = 320
$splitMain.BackColor = $bgDark

# --- 2. LEFT PANEL (Folder Tree) ---
$navPanel = New-Object System.Windows.Forms.Panel
$navPanel.Dock = "Fill"
$navPanel.BackColor = $panelDark

$lblNav = New-Object System.Windows.Forms.Label
$lblNav.Text = " NAVIGATION"
$lblNav.Dock = "Top"
$lblNav.Height = 35
$lblNav.TextAlign = "MiddleLeft"
$lblNav.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblNav.BackColor = $accentBlue

# Path display
$pathPanel = New-Object System.Windows.Forms.Panel
$pathPanel.Dock = "Top"
$pathPanel.Height = 60
$pathPanel.BackColor = $panelDark
$pathPanel.Padding = 8

$lblPath = New-Object System.Windows.Forms.Label
$lblPath.Text = "Current Path:"
$lblPath.Location = New-Object System.Drawing.Point(8, 5)
$lblPath.AutoSize = $true
$lblPath.Font = New-Object System.Drawing.Font("Segoe UI", 8)

$txtPath = New-Object System.Windows.Forms.TextBox
$txtPath.Location = New-Object System.Drawing.Point(8, 25)
$txtPath.Width = 280
$txtPath.BackColor = $bgDark
$txtPath.ForeColor = $textWhite
$txtPath.ReadOnly = $true
$txtPath.BorderStyle = "FixedSingle"

$pathPanel.Controls.AddRange(@($lblPath, $txtPath))

$treeView = New-Object System.Windows.Forms.TreeView
$treeView.Dock = "Fill"
$treeView.BackColor = $panelDark
$treeView.ForeColor = $textWhite
$treeView.LineColor = $accentBlue
$treeView.BorderStyle = "None"
$treeView.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$navPanel.Controls.AddRange(@($treeView, $pathPanel, $lblNav))
$splitMain.Panel1.Controls.Add($navPanel)

# --- 3. RIGHT PANEL (Workspace) ---
$workspacePanel = New-Object System.Windows.Forms.Panel
$workspacePanel.Dock = "Fill"

# Dashboard (Top Panel) - Increased height and better organization
$dashPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$dashPanel.Dock = "Top"
$dashPanel.Height = 260
$dashPanel.BackColor = $panelDark
$dashPanel.Padding = 10
$dashPanel.AutoScroll = $true

function Create-Group($text, $width, $height) {
    $gb = New-Object System.Windows.Forms.GroupBox
    $gb.Text = $text
    $gb.Size = New-Object System.Drawing.Size($width, $height)
    $gb.ForeColor = $textWhite
    $gb.FlatStyle = "Flat"
    $gb.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    return $gb
}

function Create-HintButton($parent, $x, $y, $hintText) {
    $btnHint = New-Object System.Windows.Forms.Button
    $btnHint.Text = "?"
    $btnHint.Size = New-Object System.Drawing.Size(20, 20)
    $btnHint.Location = New-Object System.Drawing.Point($x, $y)
    $btnHint.BackColor = $hintYellow
    $btnHint.ForeColor = [System.Drawing.Color]::Black
    $btnHint.FlatStyle = "Flat"
    $btnHint.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $btnHint.Cursor = [System.Windows.Forms.Cursors]::Hand
    
    $tooltip = New-Object System.Windows.Forms.ToolTip
    $tooltip.SetToolTip($btnHint, $hintText)
    
    $btnHint.Add_Click({
        [System.Windows.Forms.MessageBox]::Show($hintText, "Hint", "OK", "Information")
    })
    
    $parent.Controls.Add($btnHint)
    return $btnHint
}

# Module: Settings
$grpSettings = Create-Group "OPTIONS" 180 230
$chkRecurse = New-Object System.Windows.Forms.CheckBox
$chkRecurse.Text = "Include Subfolders"
$chkRecurse.Location = New-Object System.Drawing.Point(15, 30)
$chkRecurse.AutoSize = $true
$chkRecurse.ForeColor = $textWhite

$chkHidden = New-Object System.Windows.Forms.CheckBox
$chkHidden.Text = "Show Hidden Files"
$chkHidden.Location = New-Object System.Drawing.Point(15, 60)
$chkHidden.AutoSize = $true
$chkHidden.ForeColor = $textWhite

$chkFolders = New-Object System.Windows.Forms.CheckBox
$chkFolders.Text = "Include Folders"
$chkFolders.Location = New-Object System.Drawing.Point(15, 90)
$chkFolders.AutoSize = $true
$chkFolders.ForeColor = $textWhite

$chkPreview = New-Object System.Windows.Forms.CheckBox
$chkPreview.Text = "Live Preview"
$chkPreview.Location = New-Object System.Drawing.Point(15, 120)
$chkPreview.AutoSize = $true
$chkPreview.Checked = $true
$chkPreview.ForeColor = $textWhite

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "REFRESH"
$btnRefresh.Location = New-Object System.Drawing.Point(15, 150)
$btnRefresh.Size = New-Object System.Drawing.Size(145, 30)
$btnRefresh.BackColor = $accentBlue
$btnRefresh.FlatStyle = "Flat"
$btnRefresh.Cursor = [System.Windows.Forms.Cursors]::Hand

$btnClearAll = New-Object System.Windows.Forms.Button
$btnClearAll.Text = "RESET ALL"
$btnClearAll.Location = New-Object System.Drawing.Point(15, 185)
$btnClearAll.Size = New-Object System.Drawing.Size(145, 30)
$btnClearAll.BackColor = $warningOrange
$btnClearAll.FlatStyle = "Flat"
$btnClearAll.Cursor = [System.Windows.Forms.Cursors]::Hand

Create-HintButton $grpSettings 155 25 "Toggle these options and click Refresh to update the file list"
$grpSettings.Controls.AddRange(@($chkRecurse, $chkHidden, $chkFolders, $chkPreview, $btnRefresh, $btnClearAll))

# Module: Text (Replace/Find) - Enhanced
$grpText = Create-Group "FIND & REPLACE" 260 230
$lblF = New-Object System.Windows.Forms.Label
$lblF.Text = "Find:"
$lblF.Location = New-Object System.Drawing.Point(15, 30)
$lblF.Width = 40
$lblF.ForeColor = $textWhite

$txtF = New-Object System.Windows.Forms.TextBox
$txtF.Location = New-Object System.Drawing.Point(60, 27)
$txtF.Width = 120
$txtF.BackColor = $bgDark
$txtF.ForeColor = $textWhite
$txtF.BorderStyle = "FixedSingle"

$chkRegex = New-Object System.Windows.Forms.CheckBox
$chkRegex.Text = "Regex"
$chkRegex.Location = New-Object System.Drawing.Point(185, 29)
$chkRegex.AutoSize = $true
$chkRegex.ForeColor = $hintYellow

$lblR = New-Object System.Windows.Forms.Label
$lblR.Text = "Replace:"
$lblR.Location = New-Object System.Drawing.Point(15, 60)
$lblR.Width = 40
$lblR.ForeColor = $textWhite

$txtR = New-Object System.Windows.Forms.TextBox
$txtR.Location = New-Object System.Drawing.Point(60, 57)
$txtR.Width = 120
$txtR.BackColor = $bgDark
$txtR.ForeColor = $textWhite
$txtR.BorderStyle = "FixedSingle"

$btnSwap = New-Object System.Windows.Forms.Button
$btnSwap.Text = "SWAP"
$btnSwap.Location = New-Object System.Drawing.Point(185, 27)
$btnSwap.Size = New-Object System.Drawing.Size(60, 52)
$btnSwap.BackColor = $warningOrange
$btnSwap.FlatStyle = "Flat"
$btnSwap.Cursor = [System.Windows.Forms.Cursors]::Hand

$btnRemoveSpaces = New-Object System.Windows.Forms.Button
$btnRemoveSpaces.Text = "Remove Spaces"
$btnRemoveSpaces.Location = New-Object System.Drawing.Point(15, 90)
$btnRemoveSpaces.Size = New-Object System.Drawing.Size(110, 25)
$btnRemoveSpaces.BackColor = $btnGray
$btnRemoveSpaces.FlatStyle = "Flat"
$btnRemoveSpaces.Cursor = [System.Windows.Forms.Cursors]::Hand

$btnAddPrefix = New-Object System.Windows.Forms.Button
$btnAddPrefix.Text = "Add Prefix"
$btnAddPrefix.Location = New-Object System.Drawing.Point(130, 90)
$btnAddPrefix.Size = New-Object System.Drawing.Size(110, 25)
$btnAddPrefix.BackColor = $btnGray
$btnAddPrefix.FlatStyle = "Flat"
$btnAddPrefix.Cursor = [System.Windows.Forms.Cursors]::Hand

$btnAddSuffix = New-Object System.Windows.Forms.Button
$btnAddSuffix.Text = "Add Suffix"
$btnAddSuffix.Location = New-Object System.Drawing.Point(15, 120)
$btnAddSuffix.Size = New-Object System.Drawing.Size(110, 25)
$btnAddSuffix.BackColor = $btnGray
$btnAddSuffix.FlatStyle = "Flat"
$btnAddSuffix.Cursor = [System.Windows.Forms.Cursors]::Hand

$btnNumbering = New-Object System.Windows.Forms.Button
$btnNumbering.Text = "Add Numbering"
$btnNumbering.Location = New-Object System.Drawing.Point(130, 120)
$btnNumbering.Size = New-Object System.Drawing.Size(110, 25)
$btnNumbering.BackColor = $btnGray
$btnNumbering.FlatStyle = "Flat"
$btnNumbering.Cursor = [System.Windows.Forms.Cursors]::Hand

$btnClean = New-Object System.Windows.Forms.Button
$btnClean.Text = "AUTO CLEAN MEDIA"
$btnClean.Location = New-Object System.Drawing.Point(15, 155)
$btnClean.Size = New-Object System.Drawing.Size(225, 30)
$btnClean.BackColor = [System.Drawing.Color]::FromArgb(104, 33, 122)
$btnClean.FlatStyle = "Flat"
$btnClean.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnClean.Cursor = [System.Windows.Forms.Cursors]::Hand

$btnRemoveDates = New-Object System.Windows.Forms.Button
$btnRemoveDates.Text = "Remove Dates"
$btnRemoveDates.Location = New-Object System.Drawing.Point(15, 190)
$btnRemoveDates.Size = New-Object System.Drawing.Size(110, 25)
$btnRemoveDates.BackColor = $btnGray
$btnRemoveDates.FlatStyle = "Flat"
$btnRemoveDates.Cursor = [System.Windows.Forms.Cursors]::Hand

$btnRemoveSpecial = New-Object System.Windows.Forms.Button
$btnRemoveSpecial.Text = "Remove Special"
$btnRemoveSpecial.Location = New-Object System.Drawing.Point(130, 190)
$btnRemoveSpecial.Size = New-Object System.Drawing.Size(110, 25)
$btnRemoveSpecial.BackColor = $btnGray
$btnRemoveSpecial.FlatStyle = "Flat"
$btnRemoveSpecial.Cursor = [System.Windows.Forms.Cursors]::Hand

Create-HintButton $grpText 235 25 "Find & Replace text in filenames. Enable 'Regex' for pattern matching (e.g., '\d+' for numbers)"
$grpText.Controls.AddRange(@($lblF, $txtF, $chkRegex, $lblR, $txtR, $btnSwap, $btnRemoveSpaces, $btnAddPrefix, $btnAddSuffix, $btnNumbering, $btnClean, $btnRemoveDates, $btnRemoveSpecial))

# Module: Case
$grpCase = Create-Group "CASE" 140 230
$btnUpper = New-Object System.Windows.Forms.Button
$btnUpper.Text = "UPPERCASE"
$btnUpper.Size = New-Object System.Drawing.Size(110, 30)
$btnUpper.Location = New-Object System.Drawing.Point(15, 25)
$btnUpper.BackColor = $btnGray
$btnUpper.FlatStyle = "Flat"
$btnUpper.Cursor = [System.Windows.Forms.Cursors]::Hand

$btnLower = New-Object System.Windows.Forms.Button
$btnLower.Text = "lowercase"
$btnLower.Size = New-Object System.Drawing.Size(110, 30)
$btnLower.Location = New-Object System.Drawing.Point(15, 60)
$btnLower.BackColor = $btnGray
$btnLower.FlatStyle = "Flat"
$btnLower.Cursor = [System.Windows.Forms.Cursors]::Hand

$btnTitle = New-Object System.Windows.Forms.Button
$btnTitle.Text = "Title Case"
$btnTitle.Size = New-Object System.Drawing.Size(110, 30)
$btnTitle.Location = New-Object System.Drawing.Point(15, 95)
$btnTitle.BackColor = $btnGray
$btnTitle.FlatStyle = "Flat"
$btnTitle.Cursor = [System.Windows.Forms.Cursors]::Hand

$btnSentence = New-Object System.Windows.Forms.Button
$btnSentence.Text = "Sentence case"
$btnSentence.Size = New-Object System.Drawing.Size(110, 30)
$btnSentence.Location = New-Object System.Drawing.Point(15, 130)
$btnSentence.BackColor = $btnGray
$btnSentence.FlatStyle = "Flat"
$btnSentence.Cursor = [System.Windows.Forms.Cursors]::Hand

$btnCamelCase = New-Object System.Windows.Forms.Button
$btnCamelCase.Text = "camelCase"
$btnCamelCase.Size = New-Object System.Drawing.Size(110, 30)
$btnCamelCase.Location = New-Object System.Drawing.Point(15, 165)
$btnCamelCase.BackColor = $btnGray
$btnCamelCase.FlatStyle = "Flat"
$btnCamelCase.Cursor = [System.Windows.Forms.Cursors]::Hand

Create-HintButton $grpCase 115 25 "Change the case of all filenames. Extension case is preserved by default."
$grpCase.Controls.AddRange(@($btnUpper, $btnLower, $btnTitle, $btnSentence, $btnCamelCase))

# Module: Extension Operations
$grpExt = Create-Group "EXTENSION" 180 230
$lblNewExt = New-Object System.Windows.Forms.Label
$lblNewExt.Text = "New Extension:"
$lblNewExt.Location = New-Object System.Drawing.Point(15, 30)
$lblNewExt.AutoSize = $true
$lblNewExt.ForeColor = $textWhite

$txtNewExt = New-Object System.Windows.Forms.TextBox
$txtNewExt.Location = New-Object System.Drawing.Point(15, 50)
$txtNewExt.Width = 145
$txtNewExt.BackColor = $bgDark
$txtNewExt.ForeColor = $textWhite
$txtNewExt.BorderStyle = "FixedSingle"
$txtNewExt.Text = ".txt"

$btnChangeExt = New-Object System.Windows.Forms.Button
$btnChangeExt.Text = "Change Extension"
$btnChangeExt.Location = New-Object System.Drawing.Point(15, 80)
$btnChangeExt.Size = New-Object System.Drawing.Size(145, 30)
$btnChangeExt.BackColor = $btnGray
$btnChangeExt.FlatStyle = "Flat"
$btnChangeExt.Cursor = [System.Windows.Forms.Cursors]::Hand

$btnRemoveExt = New-Object System.Windows.Forms.Button
$btnRemoveExt.Text = "Remove Extension"
$btnRemoveExt.Location = New-Object System.Drawing.Point(15, 115)
$btnRemoveExt.Size = New-Object System.Drawing.Size(145, 30)
$btnRemoveExt.BackColor = $btnGray
$btnRemoveExt.FlatStyle = "Flat"
$btnRemoveExt.Cursor = [System.Windows.Forms.Cursors]::Hand

$btnUpperExt = New-Object System.Windows.Forms.Button
$btnUpperExt.Text = "Extension → UPPER"
$btnUpperExt.Location = New-Object System.Drawing.Point(15, 150)
$btnUpperExt.Size = New-Object System.Drawing.Size(145, 25)
$btnUpperExt.BackColor = $btnGray
$btnUpperExt.FlatStyle = "Flat"
$btnUpperExt.Cursor = [System.Windows.Forms.Cursors]::Hand

$btnLowerExt = New-Object System.Windows.Forms.Button
$btnLowerExt.Text = "Extension → lower"
$btnLowerExt.Location = New-Object System.Drawing.Point(15, 180)
$btnLowerExt.Size = New-Object System.Drawing.Size(145, 25)
$btnLowerExt.BackColor = $btnGray
$btnLowerExt.FlatStyle = "Flat"
$btnLowerExt.Cursor = [System.Windows.Forms.Cursors]::Hand

Create-HintButton $grpExt 155 25 "Modify file extensions. Include the dot (.) when typing new extensions."
$grpExt.Controls.AddRange(@($lblNewExt, $txtNewExt, $btnChangeExt, $btnRemoveExt, $btnUpperExt, $btnLowerExt))

# Module: Final Actions
$grpAct = Create-Group "EXECUTE" 240 230
$lblStats = New-Object System.Windows.Forms.Label
$lblStats.Text = "Changes: 0 files affected"
$lblStats.Location = New-Object System.Drawing.Point(15, 25)
$lblStats.Size = New-Object System.Drawing.Size(210, 20)
$lblStats.ForeColor = $hintYellow
$lblStats.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$btnPreview = New-Object System.Windows.Forms.Button
$btnPreview.Text = "PREVIEW CHANGES"
$btnPreview.Size = New-Object System.Drawing.Size(210, 35)
$btnPreview.Location = New-Object System.Drawing.Point(15, 50)
$btnPreview.BackColor = $accentBlue
$btnPreview.FlatStyle = "Flat"
$btnPreview.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnPreview.Cursor = [System.Windows.Forms.Cursors]::Hand

$btnUndo = New-Object System.Windows.Forms.Button
$btnUndo.Text = "UNDO"
$btnUndo.Size = New-Object System.Drawing.Size(210, 35)
$btnUndo.Location = New-Object System.Drawing.Point(15, 90)
$btnUndo.BackColor = [System.Drawing.Color]::FromArgb(209, 52, 56)
$btnUndo.FlatStyle = "Flat"
$btnUndo.Enabled = $false
$btnUndo.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnUndo.Cursor = [System.Windows.Forms.Cursors]::Hand

$btnApply = New-Object System.Windows.Forms.Button
$btnApply.Text = "APPLY CHANGES"
$btnApply.Size = New-Object System.Drawing.Size(210, 50)
$btnApply.Location = New-Object System.Drawing.Point(15, 130)
$btnApply.BackColor = $successGreen
$btnApply.FlatStyle = "Flat"
$btnApply.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$btnApply.Cursor = [System.Windows.Forms.Cursors]::Hand

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(15, 190)
$progressBar.Size = New-Object System.Drawing.Size(210, 25)
$progressBar.Style = "Continuous"
$progressBar.Visible = $false

Create-HintButton $grpAct 215 25 "Preview shows what will change. Apply performs the actual rename. Undo reverts the last applied change."
$grpAct.Controls.AddRange(@($lblStats, $btnPreview, $btnUndo, $btnApply, $progressBar))

$dashPanel.Controls.AddRange(@($grpSettings, $grpText, $grpCase, $grpExt, $grpAct))

# Scrollable Grid with Search
$topToolbar = New-Object System.Windows.Forms.Panel
$topToolbar.Dock = "Top"
$topToolbar.Height = 40
$topToolbar.BackColor = $panelDark
$topToolbar.Padding = 5

$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text = "Filter:"
$lblFilter.Location = New-Object System.Drawing.Point(10, 12)
$lblFilter.AutoSize = $true

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Location = New-Object System.Drawing.Point(70, 10)
$txtFilter.Width = 200
$txtFilter.BackColor = $bgDark
$txtFilter.ForeColor = $textWhite
$txtFilter.BorderStyle = "FixedSingle"

$lblGridInfo = New-Object System.Windows.Forms.Label
$lblGridInfo.Text = "Showing: All files"
$lblGridInfo.Location = New-Object System.Drawing.Point(290, 12)
$lblGridInfo.AutoSize = $true
$lblGridInfo.ForeColor = $hintYellow

$topToolbar.Controls.AddRange(@($lblFilter, $txtFilter, $lblGridInfo))

$scrollPanel = New-Object System.Windows.Forms.Panel
$scrollPanel.Dock = "Fill"
$scrollPanel.AutoScroll = $true
$scrollPanel.BackColor = $bgDark

$grid = New-Object System.Windows.Forms.TableLayoutPanel
$grid.ColumnCount = 3
$grid.Dock = "Top"
$grid.AutoSize = $true
$grid.Padding = 15
$grid.BackColor = $bgDark
$grid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 40)))
$grid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$grid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 60)))

$scrollPanel.Controls.Add($grid)
$workspacePanel.Controls.AddRange(@($scrollPanel, $topToolbar, $dashPanel))
$splitMain.Panel2.Controls.Add($workspacePanel)

$form.Controls.Add($splitMain)

# --- 4. NAVIGATION LOGIC ---
function Get-Drives {
    $treeView.Nodes.Clear()
    foreach ($drive in [System.IO.DriveInfo]::GetDrives() | Where-Object { $_.IsReady }) {
        $label = "$($drive.Name) [$($drive.DriveType)]"
        $node = $treeView.Nodes.Add($drive.Name, $label)
        $node.Tag = $drive.RootDirectory.FullName
        $node.Nodes.Add("TEMP")
    }
}

$treeView.Add_BeforeExpand({
    $node = $_.Node
    if ($node.Nodes.Count -eq 1 -and $node.Nodes[0].Text -eq "TEMP") {
        $node.Nodes.Clear()
        try {
            $dirs = Get-ChildItem -Path $node.Tag -Directory -Force:$chkHidden.Checked -ErrorAction SilentlyContinue
            foreach ($dir in $dirs) {
                $subNode = $node.Nodes.Add($dir.Name)
                $subNode.Tag = $dir.FullName
                $subNode.Nodes.Add("TEMP")
            }
        } catch {
            $statusLabel.Text = "Error accessing: $($node.Tag)"
        }
    }
})

function Load-Files($path) {
    $script:CurrentPath = $path
    $txtPath.Text = $path
    $grid.SuspendLayout()
    $grid.Controls.Clear()
    $script:fileObjects = @()
    
    $statusLabel.Text = "Loading files from: $path"
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
    try {
        $items = @()
        
        if ($chkFolders.Checked) {
            $items += Get-ChildItem -Path $path -Directory -Recurse:($chkRecurse.Checked) -Force:($chkHidden.Checked) -ErrorAction SilentlyContinue
        }
        
        $items += Get-ChildItem -Path $path -File -Recurse:($chkRecurse.Checked) -Force:($chkHidden.Checked) -ErrorAction SilentlyContinue
        
        $count = 0
        foreach ($item in $items) {
            $count++
            
            # Original name label
            $lbl = New-Object System.Windows.Forms.Label
            $lbl.Text = $item.Name
            $lbl.Dock = "Fill"
            $lbl.Height = 35
            $lbl.TextAlign = "MiddleLeft"
            $lbl.ForeColor = $textWhite
            $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $lbl.Padding = New-Object System.Windows.Forms.Padding(5)
            
            # Arrow label
            $arrow = New-Object System.Windows.Forms.Label
            $arrow.Text = "→"
            $arrow.Dock = "Fill"
            $arrow.Height = 35
            $arrow.TextAlign = "MiddleCenter"
            $arrow.ForeColor = $accentBlue
            $arrow.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
            
            # New name textbox
            $txt = New-Object System.Windows.Forms.TextBox
            $txt.Text = $item.Name
            $txt.Dock = "Fill"
            $txt.BackColor = $bgDark
            $txt.ForeColor = $textWhite
            $txt.BorderStyle = "FixedSingle"
            $txt.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            
            # Change detection
            $txt.Add_TextChanged({
                $sender = $args[0]
                $fileObj = $script:fileObjects | Where-Object { $_.Input -eq $sender } | Select-Object -First 1
                if ($fileObj) {
                    if ($sender.Text -ne $fileObj.Info.Name) {
                        $sender.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 0)
                    } else {
                        $sender.BackColor = $bgDark
                    }
                }
                Update-Statistics
            })
            
            $grid.Controls.AddRange(@($lbl, $arrow, $txt))
            $script:fileObjects += [PSCustomObject]@{ 
                Info = $item
                Input = $txt
                Label = $lbl
                Arrow = $arrow
            }
        }
        
        $statusLabel.Text = "Ready | Loaded $count items from: $path"
        $statusCount.Text = "Files: $count"
        Update-Statistics
        
    } catch {
        $statusLabel.Text = "Error loading files: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error loading files: $($_.Exception.Message)", "Error", "OK", "Error")
    }
    
    $grid.ResumeLayout()
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
}

$treeView.Add_AfterSelect({
    if ($_.Node.Tag) { 
        Load-Files $_.Node.Tag 
    }
})

# --- 5. HELPER FUNCTIONS ---
function Save-State {
    $snap = @{}
    foreach($o in $script:fileObjects) {
        $snap[$o.Info.FullName] = $o.Input.Text
    }
    $script:History.Push($snap)
    $btnUndo.Enabled = $true
}

function Update-Statistics {
    $changed = ($script:fileObjects | Where-Object { $_.Input.Text -ne $_.Info.Name }).Count
    $lblStats.Text = "Changes: $changed file(s) will be renamed"
    if ($changed -gt 0) {
        $lblStats.ForeColor = $hintYellow
    } else {
        $lblStats.ForeColor = $textWhite
    }
}

function Apply-Transform($transformFunc) {
    Save-State
    foreach($o in $script:fileObjects) {
        $newName = & $transformFunc $o.Input.Text $o.Info
        $o.Input.Text = $newName
    }
}

function Get-TitleCase($text) {
    $textInfo = (Get-Culture).TextInfo
    return $textInfo.ToTitleCase($text.ToLower())
}

function Get-SentenceCase($text) {
    if ($text.Length -gt 0) {
        return $text.Substring(0,1).ToUpper() + $text.Substring(1).ToLower()
    }
    return $text
}

function Get-CamelCase($text) {
    $words = $text -split '\s+'
    $result = $words[0].ToLower()
    for ($i = 1; $i -lt $words.Length; $i++) {
        if ($words[$i].Length -gt 0) {
            $result += $words[$i].Substring(0,1).ToUpper() + $words[$i].Substring(1).ToLower()
        }
    }
    return $result
}

# --- 6. WORKSPACE LOGIC ---

# Clean Media Files
$btnClean.Add_Click({
    if ($script:fileObjects.Count -eq 0) { return }
    Save-State
    
    $junk = @(
        '\.10bit', '\.2160p', '\.1080p', '\.720p', '\.480p', '\.WEBRip', '\.BluRay', 
        '\.x265', '\.x264', '\.HEVC', '\.AAC', '\.6CH', '\.5\.1', '-PSA', '-RARBG',
        '\.WEB-DL', '\.HDR', '\.DDP5\.1', '\[YTS\.MX\]', '\[YTS\.AG\]'
    )
    
    foreach($o in $script:fileObjects) {
        $ext = [System.IO.Path]::GetExtension($o.Input.Text)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($o.Input.Text)
        $cleaned = $base.Replace(".", " ")
        
        foreach ($pattern in $junk) {
            $cleaned = $cleaned -replace $pattern, ""
        }
        
        $cleaned = $cleaned -replace "\s+", " "
        $cleaned = $cleaned.Trim()
        $o.Input.Text = $cleaned + $ext
    }
    
    $statusLabel.Text = "Media cleanup applied"
})

# Find & Replace
$btnSwap.Add_Click({
    if ($script:fileObjects.Count -eq 0 -or $txtF.Text -eq "") { return }
    Save-State
    
    foreach($o in $script:fileObjects) {
        if ($chkRegex.Checked) {
            try {
                $o.Input.Text = $o.Input.Text -replace $txtF.Text, $txtR.Text
            } catch {
                $statusLabel.Text = "Regex error: $($_.Exception.Message)"
            }
        } else {
            $o.Input.Text = $o.Input.Text.Replace($txtF.Text, $txtR.Text)
        }
    }
    
    $statusLabel.Text = "Find & Replace applied"
})

# Case transformations
$btnUpper.Add_Click({
    if ($script:fileObjects.Count -eq 0) { return }
    Apply-Transform {
        param($name, $info)
        $ext = [System.IO.Path]::GetExtension($name)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($name)
        return $base.ToUpper() + $ext
    }
    $statusLabel.Text = "UPPERCASE applied"
})

$btnLower.Add_Click({
    if ($script:fileObjects.Count -eq 0) { return }
    Apply-Transform {
        param($name, $info)
        $ext = [System.IO.Path]::GetExtension($name)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($name)
        return $base.ToLower() + $ext
    }
    $statusLabel.Text = "lowercase applied"
})

$btnTitle.Add_Click({
    if ($script:fileObjects.Count -eq 0) { return }
    Apply-Transform {
        param($name, $info)
        $ext = [System.IO.Path]::GetExtension($name)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($name)
        return (Get-TitleCase $base) + $ext
    }
    $statusLabel.Text = "Title Case applied"
})

$btnSentence.Add_Click({
    if ($script:fileObjects.Count -eq 0) { return }
    Apply-Transform {
        param($name, $info)
        $ext = [System.IO.Path]::GetExtension($name)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($name)
        return (Get-SentenceCase $base) + $ext
    }
    $statusLabel.Text = "Sentence case applied"
})

$btnCamelCase.Add_Click({
    if ($script:fileObjects.Count -eq 0) { return }
    Apply-Transform {
        param($name, $info)
        $ext = [System.IO.Path]::GetExtension($name)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($name)
        return (Get-CamelCase $base) + $ext
    }
    $statusLabel.Text = "camelCase applied"
})

# Remove Spaces
$btnRemoveSpaces.Add_Click({
    if ($script:fileObjects.Count -eq 0) { return }
    Apply-Transform {
        param($name, $info)
        $ext = [System.IO.Path]::GetExtension($name)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($name)
        return ($base -replace '\s+', '') + $ext
    }
    $statusLabel.Text = "Spaces removed"
})

# Add Prefix
$btnAddPrefix.Add_Click({
    if ($script:fileObjects.Count -eq 0) { return }
    $prefix = [Microsoft.VisualBasic.Interaction]::InputBox("Enter prefix:", "Add Prefix", "PREFIX_")
    if ($prefix -ne "") {
        Apply-Transform {
            param($name, $info)
            return $prefix + $name
        }
        $statusLabel.Text = "Prefix added"
    }
})

# Add Suffix
$btnAddSuffix.Add_Click({
    if ($script:fileObjects.Count -eq 0) { return }
    $suffix = [Microsoft.VisualBasic.Interaction]::InputBox("Enter suffix:", "Add Suffix", "_SUFFIX")
    if ($suffix -ne "") {
        Apply-Transform {
            param($name, $info)
            $ext = [System.IO.Path]::GetExtension($name)
            $base = [System.IO.Path]::GetFileNameWithoutExtension($name)
            return $base + $suffix + $ext
        }
        $statusLabel.Text = "Suffix added"
    }
})

# Add Numbering
$btnNumbering.Add_Click({
    if ($script:fileObjects.Count -eq 0) { return }
    
    $form2 = New-Object System.Windows.Forms.Form
    $form2.Text = "Add Numbering"
    $form2.Size = New-Object System.Drawing.Size(350, 250)
    $form2.StartPosition = "CenterParent"
    $form2.BackColor = $bgDark
    $form2.ForeColor = $textWhite
    $form2.FormBorderStyle = "FixedDialog"
    $form2.MaximizeBox = $false
    
    $lbl1 = New-Object System.Windows.Forms.Label
    $lbl1.Text = "Start Number:"
    $lbl1.Location = New-Object System.Drawing.Point(20, 20)
    $lbl1.AutoSize = $true
    
    $numStart = New-Object System.Windows.Forms.NumericUpDown
    $numStart.Location = New-Object System.Drawing.Point(20, 45)
    $numStart.Width = 100
    $numStart.Minimum = 0
    $numStart.Maximum = 9999
    $numStart.Value = 1
    $numStart.BackColor = $bgDark
    $numStart.ForeColor = $textWhite
    
    $lbl2 = New-Object System.Windows.Forms.Label
    $lbl2.Text = "Padding (digits):"
    $lbl2.Location = New-Object System.Drawing.Point(20, 80)
    $lbl2.AutoSize = $true
    
    $numPad = New-Object System.Windows.Forms.NumericUpDown
    $numPad.Location = New-Object System.Drawing.Point(20, 105)
    $numPad.Width = 100
    $numPad.Minimum = 1
    $numPad.Maximum = 10
    $numPad.Value = 3
    $numPad.BackColor = $bgDark
    $numPad.ForeColor = $textWhite
    
    $lbl3 = New-Object System.Windows.Forms.Label
    $lbl3.Text = "Position:"
    $lbl3.Location = New-Object System.Drawing.Point(180, 20)
    $lbl3.AutoSize = $true
    
    $radioPrefix = New-Object System.Windows.Forms.RadioButton
    $radioPrefix.Text = "Prefix"
    $radioPrefix.Location = New-Object System.Drawing.Point(180, 45)
    $radioPrefix.AutoSize = $true
    $radioPrefix.Checked = $true
    
    $radioSuffix = New-Object System.Windows.Forms.RadioButton
    $radioSuffix.Text = "Suffix"
    $radioSuffix.Location = New-Object System.Drawing.Point(180, 70)
    $radioSuffix.AutoSize = $true
    
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Location = New-Object System.Drawing.Point(100, 160)
    $btnOK.Size = New-Object System.Drawing.Size(75, 30)
    $btnOK.DialogResult = "OK"
    $btnOK.BackColor = $successGreen
    $btnOK.FlatStyle = "Flat"
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(180, 160)
    $btnCancel.Size = New-Object System.Drawing.Size(75, 30)
    $btnCancel.DialogResult = "Cancel"
    $btnCancel.BackColor = $btnGray
    $btnCancel.FlatStyle = "Flat"
    
    $form2.Controls.AddRange(@($lbl1, $numStart, $lbl2, $numPad, $lbl3, $radioPrefix, $radioSuffix, $btnOK, $btnCancel))
    $form2.AcceptButton = $btnOK
    $form2.CancelButton = $btnCancel
    
    if ($form2.ShowDialog() -eq "OK") {
        Save-State
        $counter = [int]$numStart.Value
        $padding = [int]$numPad.Value
        $isPrefix = $radioPrefix.Checked
        
        foreach($o in $script:fileObjects) {
            $ext = [System.IO.Path]::GetExtension($o.Input.Text)
            $base = [System.IO.Path]::GetFileNameWithoutExtension($o.Input.Text)
            $numStr = $counter.ToString("D$padding")
            
            if ($isPrefix) {
                $o.Input.Text = $numStr + "_" + $base + $ext
            } else {
                $o.Input.Text = $base + "_" + $numStr + $ext
            }
            $counter++
        }
        $statusLabel.Text = "Numbering applied"
    }
})

# Remove Dates
$btnRemoveDates.Add_Click({
    if ($script:fileObjects.Count -eq 0) { return }
    Apply-Transform {
        param($name, $info)
        # Remove common date patterns: YYYY-MM-DD, DD-MM-YYYY, MM-DD-YYYY, etc.
        $result = $name -replace '\d{4}[-_.]\d{2}[-_.]\d{2}', ''
        $result = $result -replace '\d{2}[-_.]\d{2}[-_.]\d{4}', ''
        $result = $result -replace '\d{8}', ''
        $result = $result -replace '\s+', ' '
        return $result.Trim()
    }
    $statusLabel.Text = "Dates removed"
})

# Remove Special Characters
$btnRemoveSpecial.Add_Click({
    if ($script:fileObjects.Count -eq 0) { return }
    Apply-Transform {
        param($name, $info)
        $ext = [System.IO.Path]::GetExtension($name)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($name)
        $cleaned = $base -replace '[^\w\s\-.]', ''
        $cleaned = $cleaned -replace '\s+', ' '
        return $cleaned.Trim() + $ext
    }
    $statusLabel.Text = "Special characters removed"
})

# Extension Operations
$btnChangeExt.Add_Click({
    if ($script:fileObjects.Count -eq 0) { return }
    $newExt = $txtNewExt.Text
    if (-not $newExt.StartsWith(".")) {
        $newExt = "." + $newExt
    }
    
    Apply-Transform {
        param($name, $info)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($name)
        return $base + $newExt
    }
    $statusLabel.Text = "Extension changed to $newExt"
})

$btnRemoveExt.Add_Click({
    if ($script:fileObjects.Count -eq 0) { return }
    Apply-Transform {
        param($name, $info)
        return [System.IO.Path]::GetFileNameWithoutExtension($name)
    }
    $statusLabel.Text = "Extensions removed"
})

$btnUpperExt.Add_Click({
    if ($script:fileObjects.Count -eq 0) { return }
    Apply-Transform {
        param($name, $info)
        $ext = [System.IO.Path]::GetExtension($name)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($name)
        return $base + $ext.ToUpper()
    }
    $statusLabel.Text = "Extensions converted to UPPERCASE"
})

$btnLowerExt.Add_Click({
    if ($script:fileObjects.Count -eq 0) { return }
    Apply-Transform {
        param($name, $info)
        $ext = [System.IO.Path]::GetExtension($name)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($name)
        return $base + $ext.ToLower()
    }
    $statusLabel.Text = "Extensions converted to lowercase"
})

# Preview Changes
$btnPreview.Add_Click({
    $toRename = $script:fileObjects | Where-Object { $_.Input.Text -ne $_.Info.Name }
    if ($toRename.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No changes to preview.", "Preview", "OK", "Information")
        return
    }
    
    $preview = "The following $($toRename.Count) file(s) will be renamed:`n`n"
    $count = 0
    foreach($o in $toRename) {
        $count++
        $preview += "[$count] $($o.Info.Name) → $($o.Input.Text)`n"
        if ($count -ge 50) {
            $preview += "`n... and $($toRename.Count - 50) more files"
            break
        }
    }
    
    $form2 = New-Object System.Windows.Forms.Form
    $form2.Text = "Preview Changes"
    $form2.Size = New-Object System.Drawing.Size(700, 500)
    $form2.StartPosition = "CenterParent"
    $form2.BackColor = $bgDark
    $form2.ForeColor = $textWhite
    
    $txtPreview = New-Object System.Windows.Forms.TextBox
    $txtPreview.Multiline = $true
    $txtPreview.ScrollBars = "Vertical"
    $txtPreview.Dock = "Fill"
    $txtPreview.Text = $preview
    $txtPreview.ReadOnly = $true
    $txtPreview.BackColor = $bgDark
    $txtPreview.ForeColor = $textWhite
    $txtPreview.Font = New-Object System.Drawing.Font("Consolas", 9)
    
    $form2.Controls.Add($txtPreview)
    $form2.ShowDialog()
})

# Apply Changes
$btnApply.Add_Click({
    $toRename = $script:fileObjects | Where-Object { $_.Input.Text -ne $_.Info.Name }
    
    if($toRename.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No changes to apply.", "Info", "OK", "Information")
        return
    }
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to rename $($toRename.Count) file(s)?`n`nThis action cannot be automatically undone if files are moved or deleted after renaming.",
        "Confirm Rename",
        "YesNo",
        "Warning"
    )
    
    if($result -eq "Yes") {
        # Save state before renaming
        $preRenameState = @{}
        foreach($o in $toRename) {
            $preRenameState[$o.Info.FullName] = @{
                OldName = $o.Info.Name
                NewName = $o.Input.Text
                Path = $o.Info.DirectoryName
            }
        }
        $script:History.Push($preRenameState)
        $btnUndo.Enabled = $true
        
        $progressBar.Visible = $true
        $progressBar.Maximum = $toRename.Count
        $progressBar.Value = 0
        
        $success = 0
        $failed = 0
        
        foreach($o in $toRename) { 
            try {
                Rename-Item -Path $o.Info.FullName -NewName $o.Input.Text -ErrorAction Stop
                $o.Input.BackColor = $bgDark
                $success++
            } catch {
                $o.Input.BackColor = [System.Drawing.Color]::DarkRed
                $o.Input.ToolTipText = "Error: $($_.Exception.Message)"
                $failed++
            }
            $progressBar.Value++
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        $progressBar.Visible = $false
        
        $msg = "Rename Complete!`n`nSuccess: $success`nFailed: $failed"
        [System.Windows.Forms.MessageBox]::Show($msg, "Complete", "OK", "Information")
        
        $statusLabel.Text = "Rename complete: $success succeeded, $failed failed"
        Load-Files $script:CurrentPath
    }
})

# Undo
$btnUndo.Add_Click({
    if($script:History.Count -eq 0) { return }
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Undo the last rename operation?",
        "Undo",
        "YesNo",
        "Question"
    )
    
    if ($result -eq "Yes") {
        $prev = $script:History.Pop()
        
        $progressBar.Visible = $true
        $progressBar.Maximum = $prev.Count
        $progressBar.Value = 0
        
        foreach($item in $prev.GetEnumerator()) {
            $newPath = Join-Path $item.Value.Path $item.Value.NewName
            $oldPath = Join-Path $item.Value.Path $item.Value.OldName
            
            if (Test-Path $newPath) {
                try {
                    Rename-Item -Path $newPath -NewName $item.Value.OldName -ErrorAction Stop
                } catch {
                    # Silently continue if undo fails
                }
            }
            $progressBar.Value++
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        $progressBar.Visible = $false
        $btnUndo.Enabled = ($script:History.Count -gt 0)
        
        [System.Windows.Forms.MessageBox]::Show("Undo complete!", "Undo", "OK", "Information")
        Load-Files $script:CurrentPath
    }
})

# Refresh
$btnRefresh.Add_Click({
    if ($script:CurrentPath -ne "") {
        Load-Files $script:CurrentPath
    }
})

# Clear All
$btnClearAll.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Reset all changes to original filenames?",
        "Reset",
        "YesNo",
        "Question"
    )
    
    if ($result -eq "Yes") {
        foreach($o in $script:fileObjects) {
            $o.Input.Text = $o.Info.Name
            $o.Input.BackColor = $bgDark
        }
        $statusLabel.Text = "All changes cleared"
    }
})

# Filter textbox
$txtFilter.Add_TextChanged({
    $filter = $txtFilter.Text.ToLower()
    $visible = 0
    
    foreach($o in $script:fileObjects) {
        if ($filter -eq "" -or $o.Info.Name.ToLower().Contains($filter)) {
            $o.Label.Visible = $true
            $o.Arrow.Visible = $true
            $o.Input.Visible = $true
            $visible++
        } else {
            $o.Label.Visible = $false
            $o.Arrow.Visible = $false
            $o.Input.Visible = $false
        }
    }
    
    if ($filter -eq "") {
        $lblGridInfo.Text = "Showing: All files"
    } else {
        $lblGridInfo.Text = "Showing: $visible of $($script:fileObjects.Count) files"
    }
})

# Initialize
Add-Type -AssemblyName Microsoft.VisualBasic
Get-Drives
$form.ShowDialog()
