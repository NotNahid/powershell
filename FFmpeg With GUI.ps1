# ========================================
# FFmpeg Smart Studio v9 - FULL FEATURED
# ========================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# For async/background work
Add-Type -AssemblyName WindowsBase

# ---------------------------
# Config Path
# ---------------------------
$script:configPath = Join-Path $env:APPDATA "FFmpegStudio"
$script:configFile = Join-Path $script:configPath "settings.json"
$script:presetsFile = Join-Path $script:configPath "custom_presets.json"
$script:thumbCachePath = Join-Path $script:configPath "thumbs"

if (-not (Test-Path $script:configPath)) { New-Item -ItemType Directory -Path $script:configPath -Force | Out-Null }
if (-not (Test-Path $script:thumbCachePath)) { New-Item -ItemType Directory -Path $script:thumbCachePath -Force | Out-Null }

# ---------------------------
# Color Theme
# ---------------------------
$theme = @{
    Background    = [System.Drawing.Color]::FromArgb(30, 30, 30)
    Panel         = [System.Drawing.Color]::FromArgb(45, 45, 45)
    Control       = [System.Drawing.Color]::FromArgb(60, 60, 60)
    Text          = [System.Drawing.Color]::FromArgb(220, 220, 220)
    Accent        = [System.Drawing.Color]::FromArgb(0, 122, 204)
    Success       = [System.Drawing.Color]::FromArgb(76, 175, 80)
    Error         = [System.Drawing.Color]::FromArgb(244, 67, 54)
    Warning       = [System.Drawing.Color]::FromArgb(255, 193, 7)
    ButtonBg      = [System.Drawing.Color]::FromArgb(70, 70, 70)
    ButtonHover   = [System.Drawing.Color]::FromArgb(90, 90, 90)
    GridBg        = [System.Drawing.Color]::FromArgb(35, 35, 35)
    GridAlt       = [System.Drawing.Color]::FromArgb(42, 42, 42)
    LogBg         = [System.Drawing.Color]::FromArgb(20, 20, 20)
    PresetBg      = [System.Drawing.Color]::FromArgb(50, 50, 55)
    PresetHover   = [System.Drawing.Color]::FromArgb(60, 65, 75)
    ThumbBg       = [System.Drawing.Color]::FromArgb(25, 25, 25)
}

$fontNormal = New-Object System.Drawing.Font("Segoe UI", 9)
$fontBold = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$fontTitle = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$fontMono = New-Object System.Drawing.Font("Cascadia Code,Consolas", 9)
$fontSmall = New-Object System.Drawing.Font("Segoe UI", 8)
$fontLarge = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)

# ---------------------------
# Global State
# ---------------------------
$script:currentProcess = $null
$script:cancelRequested = $false
$script:isProcessing = $false
$script:fileDataStore = [System.Collections.ArrayList]::new()
$script:recentOutputs = [System.Collections.ArrayList]::new()

# ---------------------------
# Config Save / Load
# ---------------------------
function Save-Config {
    $config = @{
        OutputFolder  = $txtOut.Text
        Format        = $cbFmt.SelectedIndex
        Resolution    = $cbRes.SelectedIndex
        CustomResW    = $txtCustomW.Text
        CustomResH    = $txtCustomH.Text
        UseCustomRes  = $chkCustomRes.Checked
        CRF           = $trackCRF.Value
        Volume        = $trackVol.Value
        Speed         = $txtSpeed.Text
        Prefix        = $txtPrefix.Text
        Suffix        = $txtSuffix.Text
        Mute          = $chkMute.Checked
        Grayscale     = $chkGray.Checked
        Invert        = $chkInvert.Checked
        Blur          = $chkBlur.Checked
        AudioCodec    = $cbACodec.SelectedIndex
        VideoCodec    = $cbVCodec.SelectedIndex
        HWAccel       = $cbHWAccel.SelectedIndex
        Bitrate       = $txtBitrate.Text
        FPS           = $txtFPS.Text
        TrimStart     = $txtStart.Text
        TrimEnd       = $txtEnd.Text
        AutoOpen      = $chkAutoOpen.Checked
        Overwrite     = $chkOverwrite.Checked
        ShutdownAfter = $chkShutdown.Checked
        WindowW       = $form.Width
        WindowH       = $form.Height
        RecentOutputs = @($script:recentOutputs)
    }
    $config | ConvertTo-Json -Depth 3 | Set-Content $script:configFile -Force
}

function Load-Config {
    if (-not (Test-Path $script:configFile)) { return }
    try {
        $config = Get-Content $script:configFile -Raw | ConvertFrom-Json
        if ($config.OutputFolder) { $txtOut.Text = $config.OutputFolder }
        if ($null -ne $config.Format) { $cbFmt.SelectedIndex = [Math]::Min($config.Format, $cbFmt.Items.Count - 1) }
        if ($null -ne $config.Resolution) { $cbRes.SelectedIndex = [Math]::Min($config.Resolution, $cbRes.Items.Count - 1) }
        if ($config.CustomResW) { $txtCustomW.Text = $config.CustomResW }
        if ($config.CustomResH) { $txtCustomH.Text = $config.CustomResH }
        if ($null -ne $config.UseCustomRes) { $chkCustomRes.Checked = $config.UseCustomRes }
        if ($null -ne $config.CRF) { $trackCRF.Value = [Math]::Max(0, [Math]::Min(51, $config.CRF)) }
        if ($null -ne $config.Volume) { $trackVol.Value = [Math]::Max(0, [Math]::Min(300, $config.Volume)) }
        if ($config.Speed) { $txtSpeed.Text = $config.Speed }
        if ($null -ne $config.Prefix) { $txtPrefix.Text = $config.Prefix }
        if ($null -ne $config.Suffix) { $txtSuffix.Text = $config.Suffix }
        if ($null -ne $config.Mute) { $chkMute.Checked = $config.Mute }
        if ($null -ne $config.Grayscale) { $chkGray.Checked = $config.Grayscale }
        if ($null -ne $config.Invert) { $chkInvert.Checked = $config.Invert }
        if ($null -ne $config.Blur) { $chkBlur.Checked = $config.Blur }
        if ($null -ne $config.AudioCodec) { $cbACodec.SelectedIndex = [Math]::Min($config.AudioCodec, $cbACodec.Items.Count - 1) }
        if ($null -ne $config.VideoCodec) { $cbVCodec.SelectedIndex = [Math]::Min($config.VideoCodec, $cbVCodec.Items.Count - 1) }
        if ($null -ne $config.HWAccel) { $cbHWAccel.SelectedIndex = [Math]::Min($config.HWAccel, $cbHWAccel.Items.Count - 1) }
        if ($config.Bitrate) { $txtBitrate.Text = $config.Bitrate }
        if ($config.FPS) { $txtFPS.Text = $config.FPS }
        if ($config.TrimStart) { $txtStart.Text = $config.TrimStart }
        if ($config.TrimEnd) { $txtEnd.Text = $config.TrimEnd }
        if ($null -ne $config.AutoOpen) { $chkAutoOpen.Checked = $config.AutoOpen }
        if ($null -ne $config.Overwrite) { $chkOverwrite.Checked = $config.Overwrite }
        if ($null -ne $config.ShutdownAfter) { $chkShutdown.Checked = $config.ShutdownAfter }
        if ($config.WindowW -gt 0 -and $config.WindowH -gt 0) {
            $form.Width = $config.WindowW
            $form.Height = $config.WindowH
        }
        if ($config.RecentOutputs) {
            foreach ($r in $config.RecentOutputs) { $script:recentOutputs.Add($r) | Out-Null }
        }
    } catch {
        # Silently ignore corrupt config
    }
}

# ---------------------------
# Thumbnail Generator
# ---------------------------
function Get-Thumbnail {
    param([string]$FilePath, [int]$Width = 160, [int]$Height = 90)

    $hash = [System.IO.Path]::GetFileName($FilePath).GetHashCode().ToString("X8")
    $thumbFile = Join-Path $script:thumbCachePath "$hash.jpg"

    if (Test-Path $thumbFile) {
        try { return [System.Drawing.Image]::FromFile($thumbFile) } catch { }
    }

    try {
        $result = & ffmpeg -y -i "$FilePath" -ss 00:00:01 -vframes 1 -s "${Width}x${Height}" -q:v 5 "$thumbFile" 2>&1
        if (Test-Path $thumbFile) {
            return [System.Drawing.Image]::FromFile($thumbFile)
        }
    } catch { }

    # Return placeholder
    $bmp = New-Object System.Drawing.Bitmap($Width, $Height)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.Clear($theme.ThumbBg)
    $g.DrawString("No Preview", $fontSmall, [System.Drawing.Brushes]::Gray,
        [System.Drawing.RectangleF]::new(0, 0, $Width, $Height),
        (New-Object System.Drawing.StringFormat -Property @{ Alignment = 'Center'; LineAlignment = 'Center' }))
    $g.Dispose()
    return $bmp
}

# ---------------------------
# Helpers
# ---------------------------
function New-StyledButton {
    param(
        [string]$Text,
        [int]$Width = 130,
        [int]$Height = 36,
        [System.Drawing.Color]$BgColor = $theme.ButtonBg,
        [System.Drawing.Color]$FgColor = $theme.Text
    )
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Size = New-Object System.Drawing.Size($Width, $Height)
    $btn.FlatStyle = 'Flat'
    $btn.FlatAppearance.BorderColor = $theme.Accent
    $btn.FlatAppearance.BorderSize = 1
    $btn.FlatAppearance.MouseOverBackColor = $theme.ButtonHover
    $btn.BackColor = $BgColor
    $btn.ForeColor = $FgColor
    $btn.Font = $fontBold
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    return $btn
}

function New-StyledLabel {
    param([string]$Text, [int]$X, [int]$Y, [switch]$Title)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Text
    $lbl.Location = New-Object System.Drawing.Point($X, $Y)
    $lbl.AutoSize = $true
    $lbl.ForeColor = $theme.Text
    $lbl.Font = if ($Title) { $fontTitle } else { $fontNormal }
    $lbl.BackColor = [System.Drawing.Color]::Transparent
    return $lbl
}

function New-StyledTextBox {
    param([int]$X, [int]$Y, [int]$W = 200, [string]$Default = "")
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point($X, $Y)
    $tb.Width = $W
    $tb.BackColor = $theme.Control
    $tb.ForeColor = $theme.Text
    $tb.Font = $fontNormal
    $tb.BorderStyle = 'FixedSingle'
    $tb.Text = $Default
    return $tb
}

function New-StyledCombo {
    param([int]$X, [int]$Y, [int]$W = 120, [string[]]$Items, [int]$Default = 0)
    $cb = New-Object System.Windows.Forms.ComboBox
    $cb.Location = New-Object System.Drawing.Point($X, $Y)
    $cb.Width = $W
    $cb.BackColor = $theme.Control
    $cb.ForeColor = $theme.Text
    $cb.Font = $fontNormal
    $cb.FlatStyle = 'Flat'
    $cb.DropDownStyle = 'DropDownList'
    if ($Items) { $cb.Items.AddRange($Items) }
    if ($cb.Items.Count -gt $Default) { $cb.SelectedIndex = $Default }
    return $cb
}

function New-StyledCheckBox {
    param([string]$Text, [int]$X, [int]$Y)
    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Text = $Text
    $chk.Location = New-Object System.Drawing.Point($X, $Y)
    $chk.ForeColor = $theme.Text
    $chk.Font = $fontNormal
    $chk.AutoSize = $true
    $chk.BackColor = [System.Drawing.Color]::Transparent
    return $chk
}

# ---------------------------
# Main Form
# ---------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "FFmpeg Smart Studio v9"
$form.Size = New-Object System.Drawing.Size(1400, 950)
$form.MinimumSize = New-Object System.Drawing.Size(1000, 700)
$form.StartPosition = "CenterScreen"
$form.BackColor = $theme.Background
$form.Font = $fontNormal
$form.FormBorderStyle = 'Sizable'
$form.MaximizeBox = $true
$form.Icon = [System.Drawing.SystemIcons]::Application

# Status Bar
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusBar.BackColor = $theme.Panel
$statusBar.SizingGrip = $true

$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready"
$statusLabel.ForeColor = $theme.Text
$statusLabel.Spring = $true
$statusLabel.TextAlign = 'MiddleLeft'

$statusFFmpeg = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusFFmpeg.Text = "Checking FFmpeg..."
$statusFFmpeg.ForeColor = $theme.Warning

$statusBar.Items.Add($statusLabel) | Out-Null
$statusBar.Items.Add($statusFFmpeg) | Out-Null
$form.Controls.Add($statusBar)

# Check FFmpeg availability
$ffmpegOk = $false
try {
    $ffVer = & ffmpeg -version 2>&1 | Select-Object -First 1
    if ($ffVer -match "ffmpeg version") {
        $statusFFmpeg.Text = "FFmpeg: OK"
        $statusFFmpeg.ForeColor = $theme.Success
        $ffmpegOk = $true
    }
} catch { }
if (-not $ffmpegOk) {
    $statusFFmpeg.Text = "FFmpeg: NOT FOUND"
    $statusFFmpeg.ForeColor = $theme.Error
}

# ---------------------------
# Tab Control
# ---------------------------
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabs.Font = $fontBold
$tabs.Padding = New-Object System.Drawing.Point(14, 6)

$tabConvert = New-Object System.Windows.Forms.TabPage
$tabConvert.Text = "  Convert / Edit  "
$tabConvert.BackColor = $theme.Background
$tabConvert.Padding = New-Object System.Windows.Forms.Padding(6)

$tabPresets = New-Object System.Windows.Forms.TabPage
$tabPresets.Text = "  Quick Presets  "
$tabPresets.BackColor = $theme.Background
$tabPresets.Padding = New-Object System.Windows.Forms.Padding(6)

$tabPreview = New-Object System.Windows.Forms.TabPage
$tabPreview.Text = "  Preview / Info  "
$tabPreview.BackColor = $theme.Background
$tabPreview.Padding = New-Object System.Windows.Forms.Padding(6)

$tabSettings = New-Object System.Windows.Forms.TabPage
$tabSettings.Text = "  Settings  "
$tabSettings.BackColor = $theme.Background
$tabSettings.Padding = New-Object System.Windows.Forms.Padding(6)

$tabs.TabPages.AddRange(@($tabConvert, $tabPresets, $tabPreview, $tabSettings))
$form.Controls.Add($tabs)

# ==========================================
# TAB 1: CONVERT / EDIT
# ==========================================
$mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
$mainLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$mainLayout.ColumnCount = 1
$mainLayout.RowCount = 6
$mainLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100))) | Out-Null
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 28))) | Out-Null
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 55))) | Out-Null
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 165))) | Out-Null
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 78))) | Out-Null
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 72))) | Out-Null
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 50))) | Out-Null
$tabConvert.Controls.Add($mainLayout)

# ---------------------------
# Row 0: File Queue with Thumbnails
# ---------------------------
$queuePanel = New-Object System.Windows.Forms.Panel
$queuePanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$queuePanel.BackColor = $theme.Panel
$queuePanel.Padding = New-Object System.Windows.Forms.Padding(5)

$lblQueueTitle = New-StyledLabel "File Queue" 8 4 -Title
$lblQueueTitle.ForeColor = $theme.Accent
$queuePanel.Controls.Add($lblQueueTitle)

$lblQueueCount = New-StyledLabel "(0 files)" 110 7
$lblQueueCount.ForeColor = [System.Drawing.Color]::Gray
$queuePanel.Controls.Add($lblQueueCount)

# DataGridView with image column
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(5, 26)
$grid.AllowUserToAddRows = $false
$grid.RowHeadersVisible = $false
$grid.AutoSizeColumnsMode = "Fill"
$grid.SelectionMode = 'FullRowSelect'
$grid.MultiSelect = $true
$grid.BackgroundColor = $theme.GridBg
$grid.GridColor = $theme.Control
$grid.DefaultCellStyle.BackColor = $theme.GridBg
$grid.DefaultCellStyle.ForeColor = $theme.Text
$grid.DefaultCellStyle.SelectionBackColor = $theme.Accent
$grid.DefaultCellStyle.Font = $fontNormal
$grid.AlternatingRowsDefaultCellStyle.BackColor = $theme.GridAlt
$grid.ColumnHeadersDefaultCellStyle.BackColor = $theme.Panel
$grid.ColumnHeadersDefaultCellStyle.ForeColor = $theme.Accent
$grid.ColumnHeadersDefaultCellStyle.Font = $fontBold
$grid.EnableHeadersVisualStyles = $false
$grid.BorderStyle = 'None'
$grid.CellBorderStyle = 'SingleHorizontal'
$grid.RowTemplate.Height = 50

# Thumbnail column
$imgCol = New-Object System.Windows.Forms.DataGridViewImageColumn
$imgCol.Name = "Thumb"
$imgCol.HeaderText = "Preview"
$imgCol.ImageLayout = 'Zoom'
$imgCol.FillWeight = 10
$grid.Columns.Add($imgCol) | Out-Null

$grid.Columns.Add("FileName", "File Path") | Out-Null
$grid.Columns.Add("Duration", "Duration") | Out-Null
$grid.Columns.Add("Resolution", "Resolution") | Out-Null
$grid.Columns.Add("Codec", "Codec") | Out-Null
$grid.Columns.Add("Size", "Size") | Out-Null
$grid.Columns.Add("Status", "Status") | Out-Null

$grid.Columns["FileName"].FillWeight = 35
$grid.Columns["Duration"].FillWeight = 10
$grid.Columns["Resolution"].FillWeight = 10
$grid.Columns["Codec"].FillWeight = 10
$grid.Columns["Size"].FillWeight = 10
$grid.Columns["Status"].FillWeight = 10

$queuePanel.Controls.Add($grid)

# Queue side buttons
$queueBtnPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$queueBtnPanel.Size = New-Object System.Drawing.Size(90, 200)
$queueBtnPanel.FlowDirection = 'TopDown'
$queueBtnPanel.BackColor = $theme.Panel
$queueBtnPanel.WrapContents = $false

$btnMoveUp = New-StyledButton "Move Up" 86 30
$btnMoveDown = New-StyledButton "Move Down" 86 30
$btnRemove = New-StyledButton "Remove" 86 30 $theme.Error ([System.Drawing.Color]::White)
$btnPreviewFile = New-StyledButton "Preview" 86 30
$queueBtnPanel.Controls.AddRange(@($btnMoveUp, $btnMoveDown, $btnRemove, $btnPreviewFile))
$queuePanel.Controls.Add($queueBtnPanel)

$queuePanel.Add_Resize({
    $grid.Size = New-Object System.Drawing.Size(($queuePanel.Width - 105), ($queuePanel.Height - 32))
    $queueBtnPanel.Location = New-Object System.Drawing.Point(($queuePanel.Width - 96), 26)
    $queueBtnPanel.Height = $queuePanel.Height - 32
})

$mainLayout.Controls.Add($queuePanel, 0, 0)

# Drag & Drop
$grid.AllowDrop = $true
$grid.Add_DragEnter({ param($s, $e)
    if ($e.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) { $e.Effect = "Copy" }
})

# ---------------------------
# FFprobe / Media Info
# ---------------------------
function Get-MediaInfo {
    param([string]$FilePath)
    $info = @{ Duration = "N/A"; Resolution = "N/A"; Size = "N/A"; Codec = "N/A";
               Width = 0; Height = 0; DurationSec = 0; Bitrate = "N/A"; AudioCodec = "N/A";
               FrameRate = "N/A" }
    try {
        $sizeBytes = (Get-Item $FilePath).Length
        if ($sizeBytes -gt 1GB) { $info.Size = "{0:N2} GB" -f ($sizeBytes / 1GB) }
        elseif ($sizeBytes -gt 1MB) { $info.Size = "{0:N1} MB" -f ($sizeBytes / 1MB) }
        else { $info.Size = "{0:N0} KB" -f ($sizeBytes / 1KB) }
    } catch {}
    try {
        $probe = & ffprobe -v quiet -print_format json -show_format -show_streams "$FilePath" 2>&1 | ConvertFrom-Json
        if ($probe.format.duration) {
            $info.DurationSec = [double]$probe.format.duration
            $ts = [TimeSpan]::FromSeconds($info.DurationSec)
            $info.Duration = "{0:D2}:{1:D2}:{2:D2}" -f $ts.Hours, $ts.Minutes, $ts.Seconds
        }
        if ($probe.format.bit_rate) {
            $br = [double]$probe.format.bit_rate
            $info.Bitrate = if ($br -gt 1000000) { "{0:N1} Mbps" -f ($br / 1000000) } else { "{0:N0} kbps" -f ($br / 1000) }
        }
        $vidStream = $probe.streams | Where-Object { $_.codec_type -eq "video" } | Select-Object -First 1
        if ($vidStream) {
            $info.Resolution = "$($vidStream.width)x$($vidStream.height)"
            $info.Width = [int]$vidStream.width
            $info.Height = [int]$vidStream.height
            $info.Codec = $vidStream.codec_name
            if ($vidStream.r_frame_rate -match "(\d+)/(\d+)") {
                $fps = [math]::Round([double]$Matches[1] / [double]$Matches[2], 2)
                $info.FrameRate = "$fps fps"
            }
        }
        $audStream = $probe.streams | Where-Object { $_.codec_type -eq "audio" } | Select-Object -First 1
        if ($audStream) { $info.AudioCodec = $audStream.codec_name }
    } catch {}
    return $info
}

function Add-FileToQueue {
    param([string]$FilePath)
    foreach ($row in $grid.Rows) {
        if ($row.Cells["FileName"].Value -eq $FilePath) { return }
    }
    $info = Get-MediaInfo $FilePath
    $script:fileDataStore.Add(@{ Path = $FilePath; Info = $info }) | Out-Null

    $rowIdx = $grid.Rows.Add()
    $grid.Rows[$rowIdx].Cells["FileName"].Value = $FilePath
    $grid.Rows[$rowIdx].Cells["Duration"].Value = $info.Duration
    $grid.Rows[$rowIdx].Cells["Resolution"].Value = $info.Resolution
    $grid.Rows[$rowIdx].Cells["Codec"].Value = $info.Codec
    $grid.Rows[$rowIdx].Cells["Size"].Value = $info.Size
    $grid.Rows[$rowIdx].Cells["Status"].Value = "Pending"

    # Generate thumbnail in background
    $thumb = Get-Thumbnail $FilePath 80 45
    if ($thumb) { $grid.Rows[$rowIdx].Cells["Thumb"].Value = $thumb }

    $lblQueueCount.Text = "($($grid.Rows.Count) files)"
}

$grid.Add_DragDrop({ param($s, $e)
    $files = $e.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
    foreach ($f in $files) {
        if (Test-Path $f -PathType Container) {
            Get-ChildItem $f -Recurse -File -Include *.mp4,*.mkv,*.avi,*.mov,*.wmv,*.flv,*.webm,*.mp3,*.wav,*.flac | ForEach-Object {
                Add-FileToQueue $_.FullName
            }
        } else {
            Add-FileToQueue $f
        }
    }
    Log "Added files via drag and drop" "success"
})

# Click on row to show preview
$grid.Add_SelectionChanged({
    if ($grid.SelectedRows.Count -eq 1) {
        $filePath = $grid.SelectedRows[0].Cells["FileName"].Value
        if ($filePath) { Update-PreviewPanel $filePath }
    }
})

# ---------------------------
# Row 1: Output Settings (responsive)
# ---------------------------
$outputRow = New-Object System.Windows.Forms.Panel
$outputRow.Dock = [System.Windows.Forms.DockStyle]::Fill
$outputRow.BackColor = $theme.Panel

$lblOut = New-StyledLabel "Output:" 8 8
$lblOut.ForeColor = $theme.Text
$outputRow.Controls.Add($lblOut)

$txtOut = New-StyledTextBox 65 5 300
$txtOut.Anchor = 'Top,Left,Right'
$outputRow.Controls.Add($txtOut)

$btnOut = New-StyledButton "Browse" 70 26
$btnOut.Anchor = 'Top,Right'
$btnOut.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($fbd.ShowDialog() -eq 'OK') { $txtOut.Text = $fbd.SelectedPath }
})
$outputRow.Controls.Add($btnOut)

$lblFmt = New-StyledLabel "Format:" 0 8
$lblFmt.ForeColor = $theme.Text
$lblFmt.Anchor = 'Top,Right'
$outputRow.Controls.Add($lblFmt)

$cbFmt = New-StyledCombo 0 5 70 @("mp4","mkv","avi","mov","webm","mp3","wav","flac","gif","ts")
$cbFmt.Anchor = 'Top,Right'
$outputRow.Controls.Add($cbFmt)

$lblPrefix = New-StyledLabel "Prefix:" 0 8
$lblPrefix.ForeColor = $theme.Text
$lblPrefix.Anchor = 'Top,Right'
$outputRow.Controls.Add($lblPrefix)
$txtPrefix = New-StyledTextBox 0 5 70
$txtPrefix.Anchor = 'Top,Right'
$outputRow.Controls.Add($txtPrefix)

$lblSuffix = New-StyledLabel "Suffix:" 0 8
$lblSuffix.ForeColor = $theme.Text
$lblSuffix.Anchor = 'Top,Right'
$outputRow.Controls.Add($lblSuffix)
$txtSuffix = New-StyledTextBox 0 5 85 "_out"
$txtSuffix.Anchor = 'Top,Right'
$outputRow.Controls.Add($txtSuffix)

$outputRow.Add_Resize({
    $w = $outputRow.Width
    $txtOut.Width = [Math]::Max(120, $w - 640)
    $btnOut.Location = New-Object System.Drawing.Point(($txtOut.Right + 5), 3)
    $lblFmt.Location = New-Object System.Drawing.Point(($btnOut.Right + 12), 8)
    $cbFmt.Location = New-Object System.Drawing.Point(($lblFmt.Right + 3), 5)
    $lblPrefix.Location = New-Object System.Drawing.Point(($cbFmt.Right + 12), 8)
    $txtPrefix.Location = New-Object System.Drawing.Point(($lblPrefix.Right + 3), 5)
    $lblSuffix.Location = New-Object System.Drawing.Point(($txtPrefix.Right + 8), 8)
    $txtSuffix.Location = New-Object System.Drawing.Point(($lblSuffix.Right + 3), 5)
})

$mainLayout.Controls.Add($outputRow, 0, 1)

# ---------------------------
# Row 2: Settings (Video + Audio side by side)
# ---------------------------
$settingsRow = New-Object System.Windows.Forms.TableLayoutPanel
$settingsRow.Dock = [System.Windows.Forms.DockStyle]::Fill
$settingsRow.ColumnCount = 2
$settingsRow.RowCount = 1
$settingsRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 55))) | Out-Null
$settingsRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 45))) | Out-Null

# --- Video Settings ---
$pnlVideo = New-Object System.Windows.Forms.Panel
$pnlVideo.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlVideo.BackColor = $theme.Panel
$pnlVideo.Margin = New-Object System.Windows.Forms.Padding(0, 0, 3, 0)

$lblVT = New-StyledLabel "Video" 8 3 -Title
$lblVT.ForeColor = $theme.Accent
$pnlVideo.Controls.Add($lblVT)

# Resolution preset
$lblRes = New-StyledLabel "Resolution:" 10 30
$lblRes.ForeColor = $theme.Text
$pnlVideo.Controls.Add($lblRes)

$cbRes = New-StyledCombo 90 27 155 @("Original","3840x2160 (4K)","2560x1440 (1440p)","1920x1080 (1080p)","1280x720 (720p)","854x480 (480p)","640x360 (360p)")
$pnlVideo.Controls.Add($cbRes)

# Custom resolution
$chkCustomRes = New-StyledCheckBox "Custom:" 260 30
$pnlVideo.Controls.Add($chkCustomRes)

$txtCustomW = New-StyledTextBox 340 27 55 ""
$txtCustomW.Enabled = $false
$pnlVideo.Controls.Add($txtCustomW)

$lblResX = New-StyledLabel "x" 400 30
$lblResX.ForeColor = $theme.Text
$pnlVideo.Controls.Add($lblResX)

$txtCustomH = New-StyledTextBox 415 27 55 ""
$txtCustomH.Enabled = $false
$pnlVideo.Controls.Add($txtCustomH)

$lblResPx = New-StyledLabel "px" 475 30
$lblResPx.ForeColor = [System.Drawing.Color]::Gray
$pnlVideo.Controls.Add($lblResPx)

$chkCustomRes.Add_CheckedChanged({
    $txtCustomW.Enabled = $chkCustomRes.Checked
    $txtCustomH.Enabled = $chkCustomRes.Checked
    $cbRes.Enabled = -not $chkCustomRes.Checked
})

# Quality CRF
$lblCRF = New-StyledLabel "Quality (CRF):" 10 58
$lblCRF.ForeColor = $theme.Text
$pnlVideo.Controls.Add($lblCRF)

$trackCRF = New-Object System.Windows.Forms.TrackBar
$trackCRF.Location = New-Object System.Drawing.Point(115, 52)
$trackCRF.Size = New-Object System.Drawing.Size(180, 28)
$trackCRF.Minimum = 0
$trackCRF.Maximum = 51
$trackCRF.Value = 23
$trackCRF.TickFrequency = 5
$trackCRF.BackColor = $theme.Panel
$pnlVideo.Controls.Add($trackCRF)

$lblCRFVal = New-StyledLabel "23 (Good)" 300 58
$lblCRFVal.ForeColor = $theme.Warning
$pnlVideo.Controls.Add($lblCRFVal)

$trackCRF.Add_ValueChanged({
    $val = $trackCRF.Value
    $q = switch ($true) {
        ($val -le 12) { "Lossless-like"; break }
        ($val -le 18) { "Very High"; break }
        ($val -le 23) { "Good"; break }
        ($val -le 28) { "Medium"; break }
        ($val -le 35) { "Low"; break }
        default { "Very Low" }
    }
    $lblCRFVal.Text = "$val ($q)"
})

# Video Codec
$lblVCodec = New-StyledLabel "V.Codec:" 400 58
$lblVCodec.ForeColor = $theme.Text
$pnlVideo.Controls.Add($lblVCodec)

$cbVCodec = New-StyledCombo 470 55 110 @("Auto (libx264)","libx265 (HEVC)","libvpx-vp9","copy","mpeg4")
$pnlVideo.Controls.Add($cbVCodec)

# FPS
$lblFPSLbl = New-StyledLabel "FPS:" 10 87
$lblFPSLbl.ForeColor = $theme.Text
$pnlVideo.Controls.Add($lblFPSLbl)

$txtFPS = New-StyledTextBox 45 84 45 ""
$pnlVideo.Controls.Add($txtFPS)

$lblFPSHint = New-StyledLabel "(blank=original)" 95 87
$lblFPSHint.ForeColor = [System.Drawing.Color]::Gray
$lblFPSHint.Font = $fontSmall
$pnlVideo.Controls.Add($lblFPSHint)

# Bitrate
$lblBitrate = New-StyledLabel "Bitrate:" 210 87
$lblBitrate.ForeColor = $theme.Text
$pnlVideo.Controls.Add($lblBitrate)

$txtBitrate = New-StyledTextBox 270 84 60 ""
$pnlVideo.Controls.Add($txtBitrate)

$lblBitrateHint = New-StyledLabel "(e.g. 5M, blank=auto)" 335 87
$lblBitrateHint.ForeColor = [System.Drawing.Color]::Gray
$lblBitrateHint.Font = $fontSmall
$pnlVideo.Controls.Add($lblBitrateHint)

# HW Accel
$lblHW = New-StyledLabel "HW Accel:" 10 115
$lblHW.ForeColor = $theme.Text
$pnlVideo.Controls.Add($lblHW)

$cbHWAccel = New-StyledCombo 85 112 120 @("None","NVIDIA (NVENC)","AMD (AMF)","Intel (QSV)","CUDA Decode")
$pnlVideo.Controls.Add($cbHWAccel)

# Filters
$chkGray = New-StyledCheckBox "Grayscale" 230 115
$pnlVideo.Controls.Add($chkGray)

$chkInvert = New-StyledCheckBox "Invert" 330 115
$pnlVideo.Controls.Add($chkInvert)

$chkBlur = New-StyledCheckBox "Blur" 410 115
$pnlVideo.Controls.Add($chkBlur)

# Subtitle
$lblSub = New-StyledLabel "Sub:" 10 143
$lblSub.ForeColor = $theme.Text
$pnlVideo.Controls.Add($lblSub)

$txtSub = New-StyledTextBox 45 140 380
$txtSub.Anchor = 'Top,Left,Right'
$pnlVideo.Controls.Add($txtSub)

$btnSub = New-StyledButton "..." 30 24
$btnSub.Anchor = 'Top,Right'
$btnSub.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "Subtitle Files|*.srt;*.ass;*.ssa;*.sub;*.vtt"
    if ($ofd.ShowDialog() -eq 'OK') { $txtSub.Text = $ofd.FileName }
})
$pnlVideo.Controls.Add($btnSub)

$pnlVideo.Add_Resize({
    $txtSub.Width = [Math]::Max(100, $pnlVideo.Width - 95)
    $btnSub.Location = New-Object System.Drawing.Point(($pnlVideo.Width - 40), 138)
})

$settingsRow.Controls.Add($pnlVideo, 0, 0)

# --- Audio & Trim ---
$pnlAudio = New-Object System.Windows.Forms.Panel
$pnlAudio.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlAudio.BackColor = $theme.Panel
$pnlAudio.Margin = New-Object System.Windows.Forms.Padding(3, 0, 0, 0)

$lblAT = New-StyledLabel "Audio & Trim" 8 3 -Title
$lblAT.ForeColor = $theme.Accent
$pnlAudio.Controls.Add($lblAT)

# Volume
$lblVol = New-StyledLabel "Volume:" 10 30
$lblVol.ForeColor = $theme.Text
$pnlAudio.Controls.Add($lblVol)

$trackVol = New-Object System.Windows.Forms.TrackBar
$trackVol.Location = New-Object System.Drawing.Point(70, 24)
$trackVol.Size = New-Object System.Drawing.Size(150, 28)
$trackVol.Minimum = 0
$trackVol.Maximum = 300
$trackVol.Value = 100
$trackVol.TickFrequency = 50
$trackVol.BackColor = $theme.Panel
$pnlAudio.Controls.Add($trackVol)

$lblVolVal = New-StyledLabel "100%" 225 30
$lblVolVal.ForeColor = $theme.Warning
$pnlAudio.Controls.Add($lblVolVal)

$trackVol.Add_ValueChanged({ $lblVolVal.Text = "$($trackVol.Value)%" })

# Audio Codec
$lblACodec = New-StyledLabel "A.Codec:" 280 30
$lblACodec.ForeColor = $theme.Text
$pnlAudio.Controls.Add($lblACodec)

$cbACodec = New-StyledCombo 345 27 90 @("Auto","aac","mp3","opus","flac","copy")
$pnlAudio.Controls.Add($cbACodec)

# Speed
$lblSpeed = New-StyledLabel "Speed:" 10 60
$lblSpeed.ForeColor = $theme.Text
$pnlAudio.Controls.Add($lblSpeed)

$txtSpeed = New-StyledTextBox 65 57 50 "1.0"
$pnlAudio.Controls.Add($txtSpeed)

$lblSpeedHint = New-StyledLabel "x (0.25 - 4.0)" 120 60
$lblSpeedHint.ForeColor = [System.Drawing.Color]::Gray
$lblSpeedHint.Font = $fontSmall
$pnlAudio.Controls.Add($lblSpeedHint)

# Mute
$chkMute = New-StyledCheckBox "Strip Audio" 260 60
$pnlAudio.Controls.Add($chkMute)

# Trim
$lblStart = New-StyledLabel "Trim Start:" 10 90
$lblStart.ForeColor = $theme.Text
$pnlAudio.Controls.Add($lblStart)
$txtStart = New-StyledTextBox 85 87 90 "00:00:00"
$pnlAudio.Controls.Add($txtStart)

$lblEnd = New-StyledLabel "Trim End:" 190 90
$lblEnd.ForeColor = $theme.Text
$pnlAudio.Controls.Add($lblEnd)
$txtEnd = New-StyledTextBox 260 87 90 "00:00:00"
$pnlAudio.Controls.Add($txtEnd)

# Quick trim buttons
$btnTrimReset = New-StyledButton "Reset Trim" 86 24
$btnTrimReset.Location = New-Object System.Drawing.Point(365, 87)
$btnTrimReset.Font = $fontSmall
$btnTrimReset.Add_Click({ $txtStart.Text = "00:00:00"; $txtEnd.Text = "00:00:00" })
$pnlAudio.Controls.Add($btnTrimReset)

# Extra options
$chkFade = New-StyledCheckBox "Fade In/Out" 10 120
$pnlAudio.Controls.Add($chkFade)

$chkNormalize = New-StyledCheckBox "Normalize Audio" 130 120
$pnlAudio.Controls.Add($chkNormalize)

$settingsRow.Controls.Add($pnlAudio, 1, 0)
$mainLayout.Controls.Add($settingsRow, 0, 2)

# ---------------------------
# Row 3: Progress
# ---------------------------
$progressPanel = New-Object System.Windows.Forms.Panel
$progressPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$progressPanel.BackColor = $theme.Panel

$lblCur = New-StyledLabel "Ready" 10 5
$lblCur.ForeColor = $theme.Text
$lblCur.MaximumSize = New-Object System.Drawing.Size(500, 0)
$progressPanel.Controls.Add($lblCur)

$lblETA = New-StyledLabel "" 520 5
$lblETA.ForeColor = $theme.Warning
$lblETA.Anchor = 'Top,Right'
$progressPanel.Controls.Add($lblETA)

$lblFileProgress = New-StyledLabel "File: 0%" 10 32
$lblFileProgress.ForeColor = $theme.Accent
$progressPanel.Controls.Add($lblFileProgress)

$progressFile = New-Object System.Windows.Forms.ProgressBar
$progressFile.Location = New-Object System.Drawing.Point(80, 30)
$progressFile.Size = New-Object System.Drawing.Size(350, 16)
$progressFile.Style = 'Continuous'
$progressFile.Anchor = 'Top,Left,Right'
$progressPanel.Controls.Add($progressFile)

$lblTotalProgress = New-StyledLabel "Total: 0%" 10 55
$lblTotalProgress.ForeColor = $theme.Accent
$progressPanel.Controls.Add($lblTotalProgress)

$progressTotal = New-Object System.Windows.Forms.ProgressBar
$progressTotal.Location = New-Object System.Drawing.Point(80, 53)
$progressTotal.Size = New-Object System.Drawing.Size(350, 16)
$progressTotal.Style = 'Continuous'
$progressTotal.Anchor = 'Top,Left,Right'
$progressPanel.Controls.Add($progressTotal)

$progressPanel.Add_Resize({
    $half = [int](($progressPanel.Width - 100) / 2)
    $progressFile.Width = $half
    $progressTotal.Location = New-Object System.Drawing.Point(($progressFile.Right + 80), 53)
    $progressTotal.Width = [Math]::Max(80, $progressPanel.Width - $progressTotal.Left - 10)
    $lblTotalProgress.Location = New-Object System.Drawing.Point(($progressFile.Right + 15), 55)
    $lblETA.Location = New-Object System.Drawing.Point(($progressFile.Right + 15), 32)
})

$mainLayout.Controls.Add($progressPanel, 0, 3)

# ---------------------------
# Row 4: Log
# ---------------------------
$logPanel = New-Object System.Windows.Forms.Panel
$logPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$logPanel.BackColor = $theme.Panel
$logPanel.Padding = New-Object System.Windows.Forms.Padding(5)

$lblLogT = New-StyledLabel "Log" 8 3
$lblLogT.ForeColor = $theme.Accent
$lblLogT.Font = $fontBold
$logPanel.Controls.Add($lblLogT)

$btnClearLog = New-StyledButton "Clear Log" 70 22
$btnClearLog.Font = $fontSmall
$btnClearLog.Anchor = 'Top,Right'
$btnClearLog.Add_Click({ $txtLog.Clear() })
$logPanel.Controls.Add($btnClearLog)

$btnSaveLog = New-StyledButton "Save Log" 70 22
$btnSaveLog.Font = $fontSmall
$btnSaveLog.Anchor = 'Top,Right'
$btnSaveLog.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "Text Files|*.txt|All|*.*"
    $sfd.FileName = "ffmpeg_log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    if ($sfd.ShowDialog() -eq 'OK') {
        $txtLog.Text | Set-Content $sfd.FileName
        Log "Log saved to: $($sfd.FileName)" "success"
    }
})
$logPanel.Controls.Add($btnSaveLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(5, 28)
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Both'
$txtLog.ReadOnly = $true
$txtLog.BackColor = $theme.LogBg
$txtLog.ForeColor = [System.Drawing.Color]::FromArgb(0, 255, 0)
$txtLog.Font = $fontMono
$txtLog.BorderStyle = 'None'
$txtLog.WordWrap = $false
$txtLog.Anchor = 'Top,Bottom,Left,Right'
$logPanel.Controls.Add($txtLog)

$logPanel.Add_Resize({
    $txtLog.Size = New-Object System.Drawing.Size(($logPanel.Width - 12), ($logPanel.Height - 34))
    $btnClearLog.Location = New-Object System.Drawing.Point(($logPanel.Width - 155), 2)
    $btnSaveLog.Location = New-Object System.Drawing.Point(($logPanel.Width - 80), 2)
})

$mainLayout.Controls.Add($logPanel, 0, 4)

# ---------------------------
# Row 5: Action Buttons
# ---------------------------
$actionPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$actionPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$actionPanel.BackColor = $theme.Background
$actionPanel.Padding = New-Object System.Windows.Forms.Padding(0, 5, 0, 0)
$actionPanel.WrapContents = $true

$btnAdd = New-StyledButton "Add Files" 110 38
$btnAddFolder = New-StyledButton "Add Folder" 110 38
$btnClear = New-StyledButton "Clear Queue" 110 38
$btnStart = New-StyledButton "START" 160 38 $theme.Success ([System.Drawing.Color]::White)
$btnCancel = New-StyledButton "CANCEL" 110 38 $theme.Error ([System.Drawing.Color]::White)
$btnCancel.Enabled = $false
$btnOpenFolder = New-StyledButton "Open Output" 110 38

$actionPanel.Controls.AddRange(@($btnAdd, $btnAddFolder, $btnClear, $btnStart, $btnCancel, $btnOpenFolder))
$mainLayout.Controls.Add($actionPanel, 0, 5)

# ==========================================
# TAB 2: QUICK PRESETS
# ==========================================
$presetScroll = New-Object System.Windows.Forms.FlowLayoutPanel
$presetScroll.Dock = [System.Windows.Forms.DockStyle]::Fill
$presetScroll.FlowDirection = 'LeftToRight'
$presetScroll.WrapContents = $true
$presetScroll.AutoScroll = $true
$presetScroll.BackColor = $theme.Background
$presetScroll.Padding = New-Object System.Windows.Forms.Padding(10)
$tabPresets.Controls.Add($presetScroll)

$presets = @(
    @{ Name="YouTube 1080p";    Desc="H.264 AAC optimized for YouTube";    Cat="Social";  Fmt="mp4"; Res="1920x1080 (1080p)"; CRF=20; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0 },
    @{ Name="YouTube 4K";       Desc="4K upload, high quality";            Cat="Social";  Fmt="mp4"; Res="3840x2160 (4K)";    CRF=18; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0 },
    @{ Name="Discord <25MB";    Desc="Compressed 720p for Discord";        Cat="Social";  Fmt="mp4"; Res="1280x720 (720p)";   CRF=30; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0 },
    @{ Name="Discord <8MB";     Desc="Max compress for free Discord";      Cat="Social";  Fmt="mp4"; Res="854x480 (480p)";    CRF=36; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0 },
    @{ Name="Twitter/X";        Desc="720p fast encode";                   Cat="Social";  Fmt="mp4"; Res="1280x720 (720p)";   CRF=25; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0 },
    @{ Name="Instagram Reel";   Desc="1080x1920 vertical";                Cat="Social";  Fmt="mp4"; Res="Original";           CRF=22; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0; CustomW="1080"; CustomH="1920" },
    @{ Name="TikTok";           Desc="1080x1920 vertical short";          Cat="Social";  Fmt="mp4"; Res="Original";           CRF=23; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0; CustomW="1080"; CustomH="1920" },
    @{ Name="WhatsApp";         Desc="480p small file";                    Cat="Social";  Fmt="mp4"; Res="854x480 (480p)";    CRF=32; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0 },
    @{ Name="Telegram";         Desc="720p good quality";                  Cat="Social";  Fmt="mp4"; Res="1280x720 (720p)";   CRF=24; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0 },

    @{ Name="Mute Video";       Desc="Strip all audio, keep video";        Cat="Audio";   Fmt="mp4"; Res="Original";  CRF=18; Mute=$true;  Vol=100; Speed="1.0"; VCodec=0; ACodec=0 },
    @{ Name="Extract MP3";      Desc="Audio only, MP3 format";            Cat="Audio";   Fmt="mp3"; Res="Original";  CRF=23; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=2 },
    @{ Name="Extract WAV";      Desc="Lossless WAV audio";                Cat="Audio";   Fmt="wav"; Res="Original";  CRF=23; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0 },
    @{ Name="Extract FLAC";     Desc="Lossless FLAC audio";               Cat="Audio";   Fmt="flac"; Res="Original"; CRF=23; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=4 },
    @{ Name="Boost Vol 200%";   Desc="Double audio volume";               Cat="Audio";   Fmt="mp4"; Res="Original";  CRF=18; Mute=$false; Vol=200; Speed="1.0"; VCodec=0; ACodec=0 },
    @{ Name="Reduce Vol 50%";   Desc="Halve audio volume";                Cat="Audio";   Fmt="mp4"; Res="Original";  CRF=18; Mute=$false; Vol=50;  Speed="1.0"; VCodec=0; ACodec=0 },
    @{ Name="Boost Vol 300%";   Desc="Triple audio volume";               Cat="Audio";   Fmt="mp4"; Res="Original";  CRF=18; Mute=$false; Vol=300; Speed="1.0"; VCodec=0; ACodec=0 },

    @{ Name="GIF (480p)";       Desc="Animated GIF 480p";                 Cat="Convert"; Fmt="gif"; Res="854x480 (480p)";  CRF=23; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0 },
    @{ Name="GIF (360p)";       Desc="Animated GIF 360p smaller";         Cat="Convert"; Fmt="gif"; Res="640x360 (360p)";  CRF=23; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0 },
    @{ Name="To MKV";           Desc="Remux to MKV container";            Cat="Convert"; Fmt="mkv"; Res="Original";        CRF=18; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0 },
    @{ Name="To WebM";          Desc="WebM for web";                      Cat="Convert"; Fmt="webm"; Res="Original";       CRF=23; Mute=$false; Vol=100; Speed="1.0"; VCodec=2; ACodec=3 },
    @{ Name="To AVI";           Desc="AVI legacy format";                 Cat="Convert"; Fmt="avi"; Res="Original";        CRF=18; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0 },
    @{ Name="To MOV";           Desc="MOV for Apple";                     Cat="Convert"; Fmt="mov"; Res="Original";        CRF=18; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0 },
    @{ Name="HEVC/H.265";       Desc="Modern codec, smaller files";       Cat="Convert"; Fmt="mp4"; Res="Original";        CRF=23; Mute=$false; Vol=100; Speed="1.0"; VCodec=1; ACodec=0 },

    @{ Name="720p";             Desc="Downscale to 720p";                 Cat="Resize";  Fmt="mp4"; Res="1280x720 (720p)";   CRF=20; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0 },
    @{ Name="480p";             Desc="Downscale to 480p";                 Cat="Resize";  Fmt="mp4"; Res="854x480 (480p)";    CRF=23; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0 },
    @{ Name="360p";             Desc="Downscale to 360p";                 Cat="Resize";  Fmt="mp4"; Res="640x360 (360p)";    CRF=25; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0 },
    @{ Name="1080p";            Desc="Scale to 1080p";                    Cat="Resize";  Fmt="mp4"; Res="1920x1080 (1080p)"; CRF=18; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0 },
    @{ Name="1440p";            Desc="Scale to 1440p";                    Cat="Resize";  Fmt="mp4"; Res="2560x1440 (1440p)"; CRF=18; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0 },
    @{ Name="Square 1080";      Desc="1080x1080 square crop";             Cat="Resize";  Fmt="mp4"; Res="Original"; CRF=20; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0; CustomW="1080"; CustomH="1080" },

    @{ Name="2x Speed";         Desc="Double playback speed";             Cat="Speed";   Fmt="mp4"; Res="Original"; CRF=20; Mute=$false; Vol=100; Speed="2.0"; VCodec=0; ACodec=0 },
    @{ Name="1.5x Speed";       Desc="50% faster playback";               Cat="Speed";   Fmt="mp4"; Res="Original"; CRF=20; Mute=$false; Vol=100; Speed="1.5"; VCodec=0; ACodec=0 },
    @{ Name="0.5x Slow";        Desc="Half speed slow-mo";                Cat="Speed";   Fmt="mp4"; Res="Original"; CRF=20; Mute=$false; Vol=100; Speed="0.5"; VCodec=0; ACodec=0 },
    @{ Name="0.25x Ultra Slow"; Desc="Quarter speed ultra slo-mo";        Cat="Speed";   Fmt="mp4"; Res="Original"; CRF=18; Mute=$false; Vol=100; Speed="0.25"; VCodec=0; ACodec=0 },
    @{ Name="3x Fast";          Desc="Triple speed timelapse";            Cat="Speed";   Fmt="mp4"; Res="Original"; CRF=20; Mute=$false; Vol=100; Speed="3.0"; VCodec=0; ACodec=0 },

    @{ Name="Archive Lossless";  Desc="MKV near-lossless CRF 8";          Cat="Quality"; Fmt="mkv"; Res="Original";        CRF=8;  Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=4 },
    @{ Name="4K Preserve";       Desc="4K high quality";                   Cat="Quality"; Fmt="mp4"; Res="3840x2160 (4K)"; CRF=16; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0 },
    @{ Name="Max Compress";      Desc="Smallest possible file";            Cat="Quality"; Fmt="mp4"; Res="640x360 (360p)"; CRF=42; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0 },
    @{ Name="Balanced";          Desc="Good quality, reasonable size";     Cat="Quality"; Fmt="mp4"; Res="Original";        CRF=23; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0 },
    @{ Name="High Quality";      Desc="Very high quality, larger";        Cat="Quality"; Fmt="mp4"; Res="Original";        CRF=15; Mute=$false; Vol=100; Speed="1.0"; VCodec=0; ACodec=0 }
)

$catColors = @{
    "Social"  = [System.Drawing.Color]::FromArgb(0, 122, 204)
    "Audio"   = [System.Drawing.Color]::FromArgb(156, 39, 176)
    "Convert" = [System.Drawing.Color]::FromArgb(255, 152, 0)
    "Resize"  = [System.Drawing.Color]::FromArgb(0, 150, 136)
    "Speed"   = [System.Drawing.Color]::FromArgb(233, 30, 99)
    "Quality" = [System.Drawing.Color]::FromArgb(76, 175, 80)
}

$currentCat = ""
foreach ($preset in ($presets | Sort-Object { $_.Cat })) {
    if ($preset.Cat -ne $currentCat) {
        $currentCat = $preset.Cat

        $catHeader = New-Object System.Windows.Forms.Label
        $catHeader.Text = "  $currentCat"
        $catHeader.Size = New-Object System.Drawing.Size(1300, 28)
        $catHeader.Font = $fontTitle
        $catHeader.ForeColor = if ($catColors.ContainsKey($currentCat)) { $catColors[$currentCat] } else { $theme.Accent }
        $catHeader.BackColor = [System.Drawing.Color]::Transparent
        $catHeader.TextAlign = 'MiddleLeft'
        $catHeader.Margin = New-Object System.Windows.Forms.Padding(0, 10, 0, 2)
        $presetScroll.Controls.Add($catHeader)

        $sep = New-Object System.Windows.Forms.Label
        $sep.Size = New-Object System.Drawing.Size(1300, 1)
        $sep.BackColor = $theme.Control
        $sep.Margin = New-Object System.Windows.Forms.Padding(5, 0, 5, 5)
        $presetScroll.Controls.Add($sep)
    }

    $presetBox = New-Object System.Windows.Forms.Panel
    $presetBox.Size = New-Object System.Drawing.Size(230, 75)
    $presetBox.BackColor = $theme.PresetBg
    $presetBox.Margin = New-Object System.Windows.Forms.Padding(6, 3, 6, 3)
    $presetBox.Cursor = [System.Windows.Forms.Cursors]::Hand

    $strip = New-Object System.Windows.Forms.Panel
    $strip.Size = New-Object System.Drawing.Size(4, 75)
    $strip.Location = New-Object System.Drawing.Point(0, 0)
    $strip.BackColor = if ($catColors.ContainsKey($preset.Cat)) { $catColors[$preset.Cat] } else { $theme.Accent }
    $presetBox.Controls.Add($strip)

    $pName = New-StyledLabel $preset.Name 12 5
    $pName.ForeColor = [System.Drawing.Color]::White
    $pName.Font = $fontBold
    $presetBox.Controls.Add($pName)

    $pDesc = New-StyledLabel $preset.Desc 12 25
    $pDesc.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170)
    $pDesc.Font = $fontSmall
    $pDesc.MaximumSize = New-Object System.Drawing.Size(210, 0)
    $presetBox.Controls.Add($pDesc)

    $tagText = "$($preset.Fmt.ToUpper())"
    if ($preset.CustomW) { $tagText += " | $($preset.CustomW)x$($preset.CustomH)" }
    elseif ($preset.Res -ne "Original") { $tagText += " | $($preset.Res -replace ' \(.*\)','')" }
    if ($preset.Mute) { $tagText += " | MUTE" }
    if ($preset.Speed -ne "1.0") { $tagText += " | $($preset.Speed)x" }
    $pInfo = New-StyledLabel $tagText 12 55
    $pInfo.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 110)
    $pInfo.Font = $fontSmall
    $presetBox.Controls.Add($pInfo)

    $cap = $preset
    foreach ($ctrl in @($presetBox, $strip, $pName, $pDesc, $pInfo)) {
        $ctrl.Add_MouseEnter({ $presetBox.BackColor = $theme.PresetHover }.GetNewClosure())
        $ctrl.Add_MouseLeave({ $presetBox.BackColor = $theme.PresetBg }.GetNewClosure())
        $ctrl.Add_Click({
            $tabs.SelectedTab = $tabConvert
            # Format
            for ($i = 0; $i -lt $cbFmt.Items.Count; $i++) {
                if ($cbFmt.Items[$i] -eq $cap.Fmt) { $cbFmt.SelectedIndex = $i; break }
            }
            # Resolution
            if ($cap.CustomW) {
                $chkCustomRes.Checked = $true
                $txtCustomW.Text = $cap.CustomW
                $txtCustomH.Text = $cap.CustomH
            } else {
                $chkCustomRes.Checked = $false
                for ($i = 0; $i -lt $cbRes.Items.Count; $i++) {
                    $ri = $cbRes.Items[$i].ToString()
                    if ($cap.Res -eq "Original" -and $i -eq 0) { $cbRes.SelectedIndex = $i; break }
                    if ($ri -eq $cap.Res -or $ri -match ($cap.Res -replace ' \(.*\)','')) { $cbRes.SelectedIndex = $i; break }
                }
            }
            $trackCRF.Value = $cap.CRF
            $chkMute.Checked = $cap.Mute
            $trackVol.Value = $cap.Vol
            $txtSpeed.Text = $cap.Speed
            if ($null -ne $cap.VCodec) { $cbVCodec.SelectedIndex = [Math]::Min($cap.VCodec, $cbVCodec.Items.Count - 1) }
            if ($null -ne $cap.ACodec) { $cbACodec.SelectedIndex = [Math]::Min($cap.ACodec, $cbACodec.Items.Count - 1) }
            $chkGray.Checked = $false
            $chkInvert.Checked = $false
            $chkBlur.Checked = $false
            Log "Preset: $($cap.Name) - $($cap.Desc)" "success"
        }.GetNewClosure())
    }
    $presetScroll.Controls.Add($presetBox)
}

# ==========================================
# TAB 3: PREVIEW / INFO
# ==========================================
$previewLayout = New-Object System.Windows.Forms.TableLayoutPanel
$previewLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$previewLayout.ColumnCount = 2
$previewLayout.RowCount = 1
$previewLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 45))) | Out-Null
$previewLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 55))) | Out-Null
$tabPreview.Controls.Add($previewLayout)

# Left: Thumbnail
$previewImgPanel = New-Object System.Windows.Forms.Panel
$previewImgPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$previewImgPanel.BackColor = $theme.ThumbBg
$previewImgPanel.Padding = New-Object System.Windows.Forms.Padding(10)

$picPreview = New-Object System.Windows.Forms.PictureBox
$picPreview.Dock = [System.Windows.Forms.DockStyle]::Fill
$picPreview.SizeMode = 'Zoom'
$picPreview.BackColor = $theme.ThumbBg
$previewImgPanel.Controls.Add($picPreview)

$btnPlayPreview = New-StyledButton "Play in Default Player" 200 36
$btnPlayPreview.Dock = [System.Windows.Forms.DockStyle]::Bottom
$btnPlayPreview.Add_Click({
    if ($grid.SelectedRows.Count -eq 1) {
        $f = $grid.SelectedRows[0].Cells["FileName"].Value
        if ($f -and (Test-Path $f)) { Start-Process $f }
    }
})
$previewImgPanel.Controls.Add($btnPlayPreview)

$previewLayout.Controls.Add($previewImgPanel, 0, 0)

# Right: Media Info
$previewInfoPanel = New-Object System.Windows.Forms.Panel
$previewInfoPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$previewInfoPanel.BackColor = $theme.Panel
$previewInfoPanel.Padding = New-Object System.Windows.Forms.Padding(15)

$lblPreviewTitle = New-StyledLabel "Media Information" 10 10 -Title
$lblPreviewTitle.ForeColor = $theme.Accent
$previewInfoPanel.Controls.Add($lblPreviewTitle)

$txtMediaInfo = New-Object System.Windows.Forms.TextBox
$txtMediaInfo.Location = New-Object System.Drawing.Point(10, 40)
$txtMediaInfo.Multiline = $true
$txtMediaInfo.ReadOnly = $true
$txtMediaInfo.ScrollBars = 'Vertical'
$txtMediaInfo.BackColor = $theme.LogBg
$txtMediaInfo.ForeColor = $theme.Text
$txtMediaInfo.Font = $fontMono
$txtMediaInfo.BorderStyle = 'None'
$txtMediaInfo.WordWrap = $true
$txtMediaInfo.Anchor = 'Top,Bottom,Left,Right'
$txtMediaInfo.Text = "Select a file from the queue to see details..."
$previewInfoPanel.Controls.Add($txtMediaInfo)

$previewInfoPanel.Add_Resize({
    $txtMediaInfo.Size = New-Object System.Drawing.Size(($previewInfoPanel.Width - 25), ($previewInfoPanel.Height - 55))
})

$previewLayout.Controls.Add($previewInfoPanel, 1, 0)

function Update-PreviewPanel {
    param([string]$FilePath)
    if (-not $FilePath -or -not (Test-Path $FilePath)) { return }

    # Thumbnail
    try {
        $hash = [System.IO.Path]::GetFileName($FilePath).GetHashCode().ToString("X8")
        $thumbFile = Join-Path $script:thumbCachePath "${hash}_large.jpg"
        if (-not (Test-Path $thumbFile)) {
            & ffmpeg -y -i "$FilePath" -ss 00:00:02 -vframes 1 -s "640x360" -q:v 3 "$thumbFile" 2>&1 | Out-Null
        }
        if (Test-Path $thumbFile) {
            $picPreview.Image = [System.Drawing.Image]::FromFile($thumbFile)
        }
    } catch {}

    # Media info text
    $info = Get-MediaInfo $FilePath
    $infoText = @"
File:       $(Split-Path $FilePath -Leaf)
Path:       $FilePath
Size:       $($info.Size)
Duration:   $($info.Duration)
Resolution: $($info.Resolution)
V.Codec:    $($info.Codec)
A.Codec:    $($info.AudioCodec)
Bitrate:    $($info.Bitrate)
Frame Rate: $($info.FrameRate)

--- Full FFprobe Output ---
"@
    try {
        $fullProbe = & ffprobe -v quiet -print_format json -show_format -show_streams "$FilePath" 2>&1
        $infoText += "`r`n$fullProbe"
    } catch {
        $infoText += "`r`n(FFprobe not available)"
    }

    $txtMediaInfo.Text = $infoText
}

# ==========================================
# TAB 4: SETTINGS
# ==========================================
$settingsScroll = New-Object System.Windows.Forms.Panel
$settingsScroll.Dock = [System.Windows.Forms.DockStyle]::Fill
$settingsScroll.AutoScroll = $true
$settingsScroll.BackColor = $theme.Background
$settingsScroll.Padding = New-Object System.Windows.Forms.Padding(20)
$tabSettings.Controls.Add($settingsScroll)

$lblSettingsTitle = New-StyledLabel "Settings & Preferences" 10 10 -Title
$lblSettingsTitle.ForeColor = $theme.Accent
$settingsScroll.Controls.Add($lblSettingsTitle)

# General Options
$gbGeneral = New-Object System.Windows.Forms.GroupBox
$gbGeneral.Text = "General"
$gbGeneral.ForeColor = $theme.Accent
$gbGeneral.Font = $fontBold
$gbGeneral.BackColor = $theme.Panel
$gbGeneral.Location = New-Object System.Drawing.Point(10, 45)
$gbGeneral.Size = New-Object System.Drawing.Size(500, 160)
$settingsScroll.Controls.Add($gbGeneral)

$chkAutoOpen = New-StyledCheckBox "Auto-open output folder when done" 15 25
$gbGeneral.Controls.Add($chkAutoOpen)

$chkOverwrite = New-StyledCheckBox "Overwrite files without asking" 15 50
$gbGeneral.Controls.Add($chkOverwrite)

$chkShutdown = New-StyledCheckBox "Shutdown PC after batch completes" 15 75
$chkShutdown.ForeColor = $theme.Error
$gbGeneral.Controls.Add($chkShutdown)

$chkSaveConfig = New-StyledCheckBox "Auto-save settings on exit" 15 100
$chkSaveConfig.Checked = $true
$gbGeneral.Controls.Add($chkSaveConfig)

$chkGenThumb = New-StyledCheckBox "Generate thumbnail previews (slower add)" 15 125
$chkGenThumb.Checked = $true
$gbGeneral.Controls.Add($chkGenThumb)

# Config buttons
$gbConfigBtns = New-Object System.Windows.Forms.GroupBox
$gbConfigBtns.Text = "Configuration"
$gbConfigBtns.ForeColor = $theme.Accent
$gbConfigBtns.Font = $fontBold
$gbConfigBtns.BackColor = $theme.Panel
$gbConfigBtns.Location = New-Object System.Drawing.Point(10, 215)
$gbConfigBtns.Size = New-Object System.Drawing.Size(500, 80)
$settingsScroll.Controls.Add($gbConfigBtns)

$btnSaveSettings = New-StyledButton "Save Settings" 130 34 $theme.Success ([System.Drawing.Color]::White)
$btnSaveSettings.Location = New-Object System.Drawing.Point(15, 30)
$btnSaveSettings.Add_Click({
    Save-Config
    Log "Settings saved to: $($script:configFile)" "success"
    [System.Windows.Forms.MessageBox]::Show("Settings saved!", "Saved", 'OK', 'Information')
})
$gbConfigBtns.Controls.Add($btnSaveSettings)

$btnLoadSettings = New-StyledButton "Load Settings" 130 34
$btnLoadSettings.Location = New-Object System.Drawing.Point(155, 30)
$btnLoadSettings.Add_Click({
    Load-Config
    Log "Settings loaded" "success"
})
$gbConfigBtns.Controls.Add($btnLoadSettings)

$btnResetSettings = New-StyledButton "Reset Defaults" 130 34 $theme.Error ([System.Drawing.Color]::White)
$btnResetSettings.Location = New-Object System.Drawing.Point(295, 30)
$btnResetSettings.Add_Click({
    $confirm = [System.Windows.Forms.MessageBox]::Show("Reset all settings to defaults?", "Reset", 'YesNo', 'Warning')
    if ($confirm -eq 'Yes') {
        if (Test-Path $script:configFile) { Remove-Item $script:configFile -Force }
        $cbFmt.SelectedIndex = 0; $cbRes.SelectedIndex = 0; $trackCRF.Value = 23; $trackVol.Value = 100
        $txtSpeed.Text = "1.0"; $txtPrefix.Text = ""; $txtSuffix.Text = "_out"
        $chkMute.Checked = $false; $chkGray.Checked = $false; $chkInvert.Checked = $false; $chkBlur.Checked = $false
        $cbVCodec.SelectedIndex = 0; $cbACodec.SelectedIndex = 0; $cbHWAccel.SelectedIndex = 0
        $txtBitrate.Text = ""; $txtFPS.Text = ""; $txtStart.Text = "00:00:00"; $txtEnd.Text = "00:00:00"
        $chkCustomRes.Checked = $false; $txtCustomW.Text = ""; $txtCustomH.Text = ""
        $chkAutoOpen.Checked = $false; $chkOverwrite.Checked = $false; $chkShutdown.Checked = $false
        Log "Settings reset to defaults" "warn"
    }
})
$gbConfigBtns.Controls.Add($btnResetSettings)

# FFmpeg Path
$gbFFmpeg = New-Object System.Windows.Forms.GroupBox
$gbFFmpeg.Text = "FFmpeg"
$gbFFmpeg.ForeColor = $theme.Accent
$gbFFmpeg.Font = $fontBold
$gbFFmpeg.BackColor = $theme.Panel
$gbFFmpeg.Location = New-Object System.Drawing.Point(10, 305)
$gbFFmpeg.Size = New-Object System.Drawing.Size(500, 70)
$settingsScroll.Controls.Add($gbFFmpeg)

$lblFFPath = New-StyledLabel "FFmpeg Path:" 15 30
$lblFFPath.ForeColor = $theme.Text
$gbFFmpeg.Controls.Add($lblFFPath)

$txtFFPath = New-StyledTextBox 110 27 250 "ffmpeg.exe"
$gbFFmpeg.Controls.Add($txtFFPath)

$btnFFBrowse = New-StyledButton "Browse" 70 26
$btnFFBrowse.Location = New-Object System.Drawing.Point(370, 26)
$btnFFBrowse.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "FFmpeg|ffmpeg.exe|All|*.*"
    if ($ofd.ShowDialog() -eq 'OK') { $txtFFPath.Text = $ofd.FileName }
})
$gbFFmpeg.Controls.Add($btnFFBrowse)

# Cache management
$gbCache = New-Object System.Windows.Forms.GroupBox
$gbCache.Text = "Cache"
$gbCache.ForeColor = $theme.Accent
$gbCache.Font = $fontBold
$gbCache.BackColor = $theme.Panel
$gbCache.Location = New-Object System.Drawing.Point(10, 385)
$gbCache.Size = New-Object System.Drawing.Size(500, 70)
$settingsScroll.Controls.Add($gbCache)

$btnClearThumbCache = New-StyledButton "Clear Thumbnail Cache" 180 32
$btnClearThumbCache.Location = New-Object System.Drawing.Point(15, 28)
$btnClearThumbCache.Add_Click({
    Get-ChildItem $script:thumbCachePath -File | Remove-Item -Force
    Log "Thumbnail cache cleared" "success"
    [System.Windows.Forms.MessageBox]::Show("Cache cleared!", "Done", 'OK', 'Information')
})
$gbCache.Controls.Add($btnClearThumbCache)

$lblCacheSize = New-StyledLabel "" 210 34
$lblCacheSize.ForeColor = [System.Drawing.Color]::Gray
$gbCache.Controls.Add($lblCacheSize)

try {
    $cacheSize = (Get-ChildItem $script:thumbCachePath -File -Recurse | Measure-Object -Property Length -Sum).Sum
    $lblCacheSize.Text = "Cache: {0:N1} MB" -f ($cacheSize / 1MB)
} catch { $lblCacheSize.Text = "Cache: 0 MB" }

# About
$gbAbout = New-Object System.Windows.Forms.GroupBox
$gbAbout.Text = "About"
$gbAbout.ForeColor = $theme.Accent
$gbAbout.Font = $fontBold
$gbAbout.BackColor = $theme.Panel
$gbAbout.Location = New-Object System.Drawing.Point(10, 465)
$gbAbout.Size = New-Object System.Drawing.Size(500, 80)
$settingsScroll.Controls.Add($gbAbout)

$lblAbout = New-StyledLabel "FFmpeg Smart Studio v9`nPowered by FFmpeg. Built with PowerShell." 15 25
$lblAbout.ForeColor = $theme.Text
$lblAbout.MaximumSize = New-Object System.Drawing.Size(470, 0)
$gbAbout.Controls.Add($lblAbout)

# ---------------------------
# Core Logic
# ---------------------------
function Log {
    param([string]$msg, [string]$type = "info")
    $stamp = Get-Date -Format "HH:mm:ss"
    $prefix = switch ($type) {
        "error"   { "[ERROR]" }
        "success" { "[OK]" }
        "warn"    { "[WARN]" }
        default   { "[INFO]" }
    }
    $txtLog.AppendText("[$stamp] $prefix $msg`r`n")
    $txtLog.SelectionStart = $txtLog.Text.Length
    $txtLog.ScrollToCaret()
    $statusLabel.Text = $msg
}

function Build-FFmpegArgs {
    param([string]$file, [string]$outfile)

    $vf = @()
    $af = @()
    $extraArgs = @()

    # Video filters
    if ($chkGray.Checked) { $vf += "hue=s=0" }
    if ($chkInvert.Checked) { $vf += "negate" }
    if ($chkBlur.Checked) { $vf += "boxblur=2:1" }

    # Resolution (custom or preset)
    if ($chkCustomRes.Checked -and $txtCustomW.Text -match '^\d+$' -and $txtCustomH.Text -match '^\d+$') {
        $vf += "scale=$($txtCustomW.Text):$($txtCustomH.Text)"
    } else {
        $resText = $cbRes.SelectedItem.ToString()
        if ($resText -ne "Original") {
            $resMatch = [regex]::Match($resText, "(\d+)x(\d+)")
            if ($resMatch.Success) { $vf += "scale=$($resMatch.Groups[1].Value):$($resMatch.Groups[2].Value)" }
        }
    }

    # FPS
    if ($txtFPS.Text -match '^\d+(\.\d+)?$') {
        $vf += "fps=$($txtFPS.Text)"
    }

    # Subtitles
    if ($txtSub.Text -ne "" -and (Test-Path $txtSub.Text)) {
        $cleanSub = $txtSub.Text -replace "\\", "/" -replace "'", "'\''" -replace ":", "\\:"
        $vf += "subtitles='$cleanSub'"
    }

    # Volume
    if ($trackVol.Value -ne 100) {
        $volFactor = [math]::Round($trackVol.Value / 100, 2)
        $af += "volume=$volFactor"
    }

    # Normalize audio
    if ($chkNormalize.Checked) { $af += "loudnorm" }

    # Fade
    if ($chkFade.Checked) {
        $vf += "fade=t=in:st=0:d=1,fade=t=out:st=0:d=1"
        $af += "afade=t=in:st=0:d=1,afade=t=out:st=0:d=1"
    }

    $vfArg = if ($vf.Count -gt 0) { "-vf `"$($vf -join ',')`"" } else { "" }
    $afArg = if ($af.Count -gt 0) { "-af `"$($af -join ',')`"" } else { "" }

    # Trim
    $trimArg = ""
    if ($txtStart.Text -ne "00:00:00") { $trimArg += " -ss $($txtStart.Text)" }
    if ($txtEnd.Text -ne "00:00:00") { $trimArg += " -to $($txtEnd.Text)" }

    # CRF
    $crfArg = "-crf $($trackCRF.Value)"

    # Video codec
    $vCodecArg = ""
    $vCodecText = $cbVCodec.SelectedItem.ToString()
    if ($vCodecText -match "libx265") { $vCodecArg = "-c:v libx265 -tag:v hvc1" }
    elseif ($vCodecText -match "libvpx") { $vCodecArg = "-c:v libvpx-vp9" }
    elseif ($vCodecText -match "copy") { $vCodecArg = "-c:v copy"; $crfArg = "" }
    elseif ($vCodecText -match "mpeg4") { $vCodecArg = "-c:v mpeg4" }

    # HW Accel
    $hwArg = ""
    $hwText = $cbHWAccel.SelectedItem.ToString()
    if ($hwText -match "NVIDIA") { $hwArg = "-hwaccel cuda"; if ($vCodecArg -eq "") { $vCodecArg = "-c:v h264_nvenc"; $crfArg = "-cq $($trackCRF.Value)" } }
    elseif ($hwText -match "AMD") { $hwArg = "-hwaccel auto"; if ($vCodecArg -eq "") { $vCodecArg = "-c:v h264_amf" } }
    elseif ($hwText -match "Intel") { $hwArg = "-hwaccel qsv"; if ($vCodecArg -eq "") { $vCodecArg = "-c:v h264_qsv" } }
    elseif ($hwText -match "CUDA") { $hwArg = "-hwaccel cuda" }

    # Audio codec
    $aCodecArg = ""
    if ($cbACodec.SelectedItem -ne "Auto" -and $null -ne $cbACodec.SelectedItem) {
        if ($cbACodec.SelectedItem -eq "copy") { $aCodecArg = "-c:a copy" }
        else { $aCodecArg = "-c:a $($cbACodec.SelectedItem)" }
    }

    # Bitrate
    $bitrateArg = ""
    if ($txtBitrate.Text -match '\S') { $bitrateArg = "-b:v $($txtBitrate.Text)" }

    # Mute
    $muteArg = if ($chkMute.Checked) { "-an" } else { "" }

    # Speed
    $speed = 1.0
    [double]::TryParse($txtSpeed.Text, [ref]$speed) | Out-Null

    $speedArg = ""
    if ($speed -ne 1.0 -and $speed -gt 0) {
        $vSpeed = [math]::Round(1 / $speed, 4)
        $allVideoFilters = @("setpts=$vSpeed*PTS")
        $allVideoFilters += $vf
        $filterComplex = "[0:v]$($allVideoFilters -join ',')[v]"

        if (-not $chkMute.Checked -and $cbFmt.SelectedItem -notin @("mp3","wav","flac","gif")) {
            $atempoChain = @()
            $remaining = $speed
            while ($remaining -gt 2.0) { $atempoChain += "atempo=2.0"; $remaining /= 2.0 }
            while ($remaining -lt 0.5) { $atempoChain += "atempo=0.5"; $remaining /= 0.5 }
            $atempoChain += "atempo=$([math]::Round($remaining, 4))"
            if ($af.Count -gt 0) { $atempoChain += $af }
            $filterComplex += ";[0:a]$($atempoChain -join ',')[a]"
            $speedArg = "-filter_complex `"$filterComplex`" -map `"[v]`" -map `"[a]`""
        } else {
            $speedArg = "-filter_complex `"$filterComplex`" -map `"[v]`""
        }
        $vfArg = ""
        $afArg = ""
    }

    # Audio-only
    if ($cbFmt.SelectedItem -in @("mp3","wav","flac")) {
        return "-y$trimArg -i `"$file`" $afArg -vn `"$outfile`""
    }

    # Build
    $parts = @("-y")
    if ($hwArg -ne "") { $parts += $hwArg }
    if ($trimArg.Trim() -ne "") { $parts += $trimArg.Trim() }
    $parts += "-i `"$file`""
    if ($speedArg -ne "") { $parts += $speedArg }
    elseif ($vfArg -ne "") { $parts += $vfArg }
    if ($afArg -ne "" -and $speedArg -eq "") { $parts += $afArg }
    if ($vCodecArg -ne "") { $parts += $vCodecArg }
    if ($muteArg -ne "") { $parts += $muteArg }
    if ($crfArg -ne "") { $parts += $crfArg }
    if ($bitrateArg -ne "") { $parts += $bitrateArg }
    if ($aCodecArg -ne "" -and -not $chkMute.Checked) { $parts += $aCodecArg }
    $parts += "`"$outfile`""

    return ($parts | Where-Object { $_ -ne "" }) -join " "
}

function Get-OutputFileName {
    param([string]$inputFile)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($inputFile)
    return "$($txtPrefix.Text)${baseName}$($txtSuffix.Text).$($cbFmt.SelectedItem)"
}

function Update-RowStatus {
    param($row, [string]$status)
    $row.Cells["Status"].Value = $status
    switch -Wildcard ($status) {
        "*Done*"       { $row.Cells["Status"].Style.ForeColor = $theme.Success }
        "*Error*"      { $row.Cells["Status"].Style.ForeColor = $theme.Error }
        "*Processing*" { $row.Cells["Status"].Style.ForeColor = $theme.Warning }
        "*Pending*"    { $row.Cells["Status"].Style.ForeColor = $theme.Text }
        "*Cancelled*"  { $row.Cells["Status"].Style.ForeColor = $theme.Error }
        "*Skipped*"    { $row.Cells["Status"].Style.ForeColor = $theme.Warning }
    }
}

function Parse-FFmpegProgress {
    param([string]$line, [double]$totalSeconds)
    if ($totalSeconds -le 0) { return -1 }
    if ($line -match "time=(\d{2}):(\d{2}):(\d{2})\.(\d{2})") {
        $cur = [int]$Matches[1] * 3600 + [int]$Matches[2] * 60 + [int]$Matches[3]
        return [math]::Min(100, [int](($cur / $totalSeconds) * 100))
    }
    return -1
}

function Get-DurationSeconds {
    param([string]$s)
    if ($s -match "(\d{2}):(\d{2}):(\d{2})") {
        return [int]$Matches[1] * 3600 + [int]$Matches[2] * 60 + [int]$Matches[3]
    }
    return 0
}

# ---------------------------
# Button Events
# ---------------------------
$btnAdd.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Multiselect = $true
    $ofd.Filter = "Media|*.mp4;*.mkv;*.avi;*.mov;*.wmv;*.flv;*.webm;*.ts;*.m2ts;*.mp3;*.wav;*.flac;*.aac;*.ogg;*.m4a;*.wma|All|*.*"
    if ($ofd.ShowDialog() -eq "OK") {
        foreach ($f in $ofd.FileNames) { Add-FileToQueue $f }
        Log "Added $($ofd.FileNames.Count) file(s)" "success"
    }
})

$btnAddFolder.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = "Select folder to add all media files"
    if ($fbd.ShowDialog() -eq 'OK') {
        $count = 0
        Get-ChildItem $fbd.SelectedPath -File -Include *.mp4,*.mkv,*.avi,*.mov,*.wmv,*.flv,*.webm,*.mp3,*.wav,*.flac,*.ts -Recurse | ForEach-Object {
            Add-FileToQueue $_.FullName
            $count++
        }
        Log "Added $count file(s) from folder" "success"
    }
})

$btnClear.Add_Click({
    $grid.Rows.Clear()
    $script:fileDataStore.Clear()
    $progressTotal.Value = 0; $progressFile.Value = 0
    $lblCur.Text = "Ready"; $lblETA.Text = ""
    $lblFileProgress.Text = "File: 0%"; $lblTotalProgress.Text = "Total: 0%"
    $lblQueueCount.Text = "(0 files)"
    Log "Queue cleared" "info"
})

$btnMoveUp.Add_Click({
    if ($grid.SelectedRows.Count -eq 1) {
        $idx = $grid.SelectedRows[0].Index
        if ($idx -gt 0) {
            $v = @(); foreach ($c in $grid.Rows[$idx].Cells) { $v += $c.Value }
            $grid.Rows.RemoveAt($idx); $grid.Rows.Insert($idx - 1, $v)
            $grid.Rows[$idx - 1].Selected = $true
        }
    }
})

$btnMoveDown.Add_Click({
    if ($grid.SelectedRows.Count -eq 1) {
        $idx = $grid.SelectedRows[0].Index
        if ($idx -lt $grid.Rows.Count - 1) {
            $v = @(); foreach ($c in $grid.Rows[$idx].Cells) { $v += $c.Value }
            $grid.Rows.RemoveAt($idx); $grid.Rows.Insert($idx + 1, $v)
            $grid.Rows[$idx + 1].Selected = $true
        }
    }
})

$btnRemove.Add_Click({
    $toRemove = @(); foreach ($r in $grid.SelectedRows) { $toRemove += $r.Index }
    $toRemove | Sort-Object -Descending | ForEach-Object { $grid.Rows.RemoveAt($_) }
    $lblQueueCount.Text = "($($grid.Rows.Count) files)"
    if ($toRemove.Count -gt 0) { Log "Removed $($toRemove.Count) file(s)" "info" }
})

$btnPreviewFile.Add_Click({
    if ($grid.SelectedRows.Count -eq 1) {
        $f = $grid.SelectedRows[0].Cells["FileName"].Value
        if ($f) {
            Update-PreviewPanel $f
            $tabs.SelectedTab = $tabPreview
        }
    }
})

$btnOpenFolder.Add_Click({
    if ($txtOut.Text -ne "" -and (Test-Path $txtOut.Text)) { Start-Process explorer.exe $txtOut.Text }
    else { [System.Windows.Forms.MessageBox]::Show("Output folder not set.", "Info", 'OK', 'Information') }
})

$btnCancel.Add_Click({
    $script:cancelRequested = $true
    if ($script:currentProcess -and -not $script:currentProcess.HasExited) {
        $script:currentProcess.Kill()
        Log "Cancel requested, killing process..." "warn"
    }
})

# ---------------------------
# START (Async via Timer-based approach)
# ---------------------------
$script:processQueue = [System.Collections.Queue]::new()
$script:batchStartTime = $null
$script:batchTotal = 0
$script:batchDone = 0
$script:batchSuccess = 0
$script:batchFail = 0

$processTimer = New-Object System.Windows.Forms.Timer
$processTimer.Interval = 100

$processTimer.Add_Tick({
    # Check if current process is still running
    if ($script:currentProcess -and -not $script:currentProcess.HasExited) {
        # Read available stderr
        try {
            while (-not $script:currentProcess.StandardError.EndOfStream) {
                $line = $script:currentProcess.StandardError.ReadLine()
                if ($null -eq $line) { break }

                $totalSec = Get-DurationSeconds $script:currentRow.Cells["Duration"].Value
                $pct = Parse-FFmpegProgress $line $totalSec
                if ($pct -ge 0) {
                    $progressFile.Value = $pct
                    $lblFileProgress.Text = "File: $pct%"
                }
                if ($line -notmatch "^frame=|^size=|^\s*$") { Log $line }
                break  # Process one line per tick to keep UI responsive
            }
        } catch {}
        return
    }

    # Process finished or no process - handle completion
    if ($script:currentProcess) {
        $row = $script:currentRow
        $inFile = $row.Cells["FileName"].Value
        $fileName = [System.IO.Path]::GetFileName($inFile)

        if ($script:cancelRequested) {
            Update-RowStatus $row "Cancelled"
            Log "Cancelled: $fileName" "warn"
        } elseif ($script:currentProcess.ExitCode -eq 0) {
            Update-RowStatus $row "Done"
            $script:batchSuccess++
            $outFile = $script:currentOutFile
            if (Test-Path $outFile) {
                $sz = (Get-Item $outFile).Length
                $szStr = if ($sz -gt 1MB) { "{0:N1} MB" -f ($sz / 1MB) } else { "{0:N0} KB" -f ($sz / 1KB) }
                Log "Done: $fileName -> $szStr" "success"
            }
        } else {
            Update-RowStatus $row "Error"
            $script:batchFail++
            Log "Failed (exit $($script:currentProcess.ExitCode)): $fileName" "error"
        }

        $script:batchDone++
        $totalPct = [math]::Min(100, [int](($script:batchDone / $script:batchTotal) * 100))
        $progressTotal.Value = $totalPct
        $lblTotalProgress.Text = "Total: $totalPct% ($($script:batchDone)/$($script:batchTotal))"
        $progressFile.Value = 0
        $lblFileProgress.Text = "File: 0%"

        # ETA
        $elapsed = (Get-Date) - $script:batchStartTime
        if ($script:batchDone -lt $script:batchTotal -and $script:batchDone -gt 0) {
            $avg = $elapsed.TotalSeconds / $script:batchDone
            $rem = ($script:batchTotal - $script:batchDone) * $avg
            $eta = [TimeSpan]::FromSeconds($rem)
            $lblETA.Text = "ETA: $("{0:D2}:{1:D2}:{2:D2}" -f $eta.Hours, $eta.Minutes, $eta.Seconds)"
        }

        $script:currentProcess = $null
    }

    # Check cancellation
    if ($script:cancelRequested) {
        # Mark remaining as cancelled
        while ($script:processQueue.Count -gt 0) {
            $item = $script:processQueue.Dequeue()
            Update-RowStatus $item.Row "Cancelled"
            $script:batchDone++
        }
    }

    # Start next or finish
    if ($script:processQueue.Count -eq 0 -or $script:cancelRequested) {
        $processTimer.Stop()
        $script:isProcessing = $false

        $totalElapsed = (Get-Date) - $script:batchStartTime
        $timeStr = "{0:D2}:{1:D2}:{2:D2}" -f $totalElapsed.Hours, $totalElapsed.Minutes, $totalElapsed.Seconds
        $lblCur.Text = "Done! $($script:batchSuccess) ok, $($script:batchFail) failed"
        $lblETA.Text = "Time: $timeStr"

        $btnStart.Enabled = $true; $btnCancel.Enabled = $false
        $btnAdd.Enabled = $true; $btnAddFolder.Enabled = $true; $btnClear.Enabled = $true

        Log "Batch complete: $($script:batchSuccess) success, $($script:batchFail) failed, $timeStr" "success"

        if ($chkAutoOpen.Checked -and $txtOut.Text -ne "" -and (Test-Path $txtOut.Text)) {
            Start-Process explorer.exe $txtOut.Text
        }

        if ($script:cancelRequested) {
            [System.Windows.Forms.MessageBox]::Show("Cancelled.", "Cancelled", 'OK', 'Warning')
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "Done!`n$($script:batchSuccess) success, $($script:batchFail) failed`nTime: $timeStr",
                "Complete", 'OK', 'Information')
            if ($chkShutdown.Checked) {
                $confirmShut = [System.Windows.Forms.MessageBox]::Show(
                    "Shutdown PC now?", "Shutdown", 'YesNo', 'Warning')
                if ($confirmShut -eq 'Yes') { Stop-Computer -Force }
            }
        }
        return
    }

    # Start next file
    $item = $script:processQueue.Dequeue()
    $row = $item.Row
    $inFile = $row.Cells["FileName"].Value
    $fileName = [System.IO.Path]::GetFileName($inFile)

    Update-RowStatus $row "Processing..."
    $lblCur.Text = "($($script:batchDone + 1)/$($script:batchTotal)) $fileName"

    $outName = Get-OutputFileName $inFile
    $outFile = Join-Path $txtOut.Text $outName
    $script:currentOutFile = $outFile
    $script:currentRow = $row

    # Overwrite check
    if ((Test-Path $outFile) -and -not $chkOverwrite.Checked) {
        $ow = [System.Windows.Forms.MessageBox]::Show("'$outName' exists. Overwrite?", "Exists", 'YesNo', 'Question')
        if ($ow -eq 'No') {
            Update-RowStatus $row "Skipped"
            $script:batchDone++
            return
        }
    }

    $ffArgs = Build-FFmpegArgs $inFile $outFile
    Log "ffmpeg $ffArgs"

    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.FileName = if ($txtFFPath.Text -ne "") { $txtFFPath.Text } else { "ffmpeg.exe" }
    $pinfo.Arguments = $ffArgs
    $pinfo.CreateNoWindow = $true
    $pinfo.UseShellExecute = $false
    $pinfo.RedirectStandardOutput = $true
    $pinfo.RedirectStandardError = $true

    try {
        $script:currentProcess = [System.Diagnostics.Process]::Start($pinfo)
    } catch {
        Update-RowStatus $row "Error"
        $script:batchFail++
        $script:batchDone++
        Log "Failed to start: $($_.Exception.Message)" "error"
        $script:currentProcess = $null
    }
})

$btnStart.Add_Click({
    if ($grid.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Add files first!", "Empty", 'OK', 'Warning'); return
    }
    if ($txtOut.Text -eq "") {
        [System.Windows.Forms.MessageBox]::Show("Select output folder!", "No Output", 'OK', 'Warning'); return
    }
    if (-not (Test-Path $txtOut.Text)) {
        New-Item -ItemType Directory -Path $txtOut.Text -Force | Out-Null
        Log "Created: $($txtOut.Text)" "info"
    }

    $script:cancelRequested = $false
    $script:isProcessing = $true
    $script:processQueue.Clear()
    $script:batchStartTime = Get-Date
    $script:batchTotal = $grid.Rows.Count
    $script:batchDone = 0
    $script:batchSuccess = 0
    $script:batchFail = 0
    $script:currentProcess = $null

    $btnStart.Enabled = $false; $btnCancel.Enabled = $true
    $btnAdd.Enabled = $false; $btnAddFolder.Enabled = $false; $btnClear.Enabled = $false
    $progressTotal.Value = 0; $progressFile.Value = 0

    foreach ($row in $grid.Rows) {
        $script:processQueue.Enqueue(@{ Row = $row }) | Out-Null
    }

    Log "Starting batch: $($script:batchTotal) files" "info"
    $processTimer.Start()
})

# ---------------------------
# Form Events
# ---------------------------
$form.Add_Shown({ Load-Config })

$form.Add_FormClosing({
    if ($chkSaveConfig.Checked) { Save-Config }
    # Cleanup thumbnails from PictureBox
    if ($picPreview.Image) { $picPreview.Image.Dispose() }
    $processTimer.Stop()
    if ($script:currentProcess -and -not $script:currentProcess.HasExited) {
        $script:currentProcess.Kill()
    }
})

# ---------------------------
# Keyboard Shortcuts
# ---------------------------
$form.KeyPreview = $true
$form.Add_KeyDown({
    param($s, $e)
    if ($e.Control) {
        switch ($e.KeyCode) {
            'O' { $btnAdd.PerformClick(); $e.Handled = $true }
            'S' { Save-Config; Log "Settings saved (Ctrl+S)" "success"; $e.Handled = $true }
            'R' { $btnStart.PerformClick(); $e.Handled = $true }
        }
    }
    if ($e.KeyCode -eq 'Delete' -and -not $script:isProcessing) {
        $btnRemove.PerformClick(); $e.Handled = $true
    }
    if ($e.KeyCode -eq 'Escape' -and $script:isProcessing) {
        $btnCancel.PerformClick(); $e.Handled = $true
    }
})

# ---------------------------
# Show
# ---------------------------
[void]$form.ShowDialog()
