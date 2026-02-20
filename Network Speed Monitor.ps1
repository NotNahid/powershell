Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================
# WIN32 HELPERS
# ============================================
Add-Type @"
using System;
using System.Runtime.InteropServices;

public class Win32UI {
    [DllImport("Gdi32.dll")]
    public static extern IntPtr CreateRoundRectRgn(int l, int t, int r, int b, int w, int h);

    [DllImport("user32.dll")]
    public static extern int SetWindowRgn(IntPtr hWnd, IntPtr hRgn, bool bRedraw);

    [DllImport("dwmapi.dll")]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int val, int size);

    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();

    [DllImport("user32.dll")]
    public static extern bool DestroyIcon(IntPtr handle);
}
"@

# ============================================
# DOUBLE BUFFERED CONTROLS (separate Add-Type)
# ============================================
Add-Type -ReferencedAssemblies @(
    'System.Windows.Forms',
    'System.Drawing'
) @"
using System.Windows.Forms;

public class DBPanel : Panel {
    public DBPanel() {
        this.SetStyle(
            ControlStyles.AllPaintingInWmPaint |
            ControlStyles.UserPaint |
            ControlStyles.OptimizedDoubleBuffer |
            ControlStyles.ResizeRedraw, true);
        this.UpdateStyles();
    }
}

public class DBForm : Form {
    public DBForm() {
        this.SetStyle(
            ControlStyles.AllPaintingInWmPaint |
            ControlStyles.UserPaint |
            ControlStyles.OptimizedDoubleBuffer |
            ControlStyles.ResizeRedraw, true);
        this.UpdateStyles();
    }
}
"@

[Win32UI]::SetProcessDPIAware()

# ============================================
# DESIGN SYSTEM
# ============================================
$theme = @{
    BgPrimary      = [System.Drawing.Color]::FromArgb(18, 18, 24)
    BgSecondary    = [System.Drawing.Color]::FromArgb(26, 26, 36)
    BgCard         = [System.Drawing.Color]::FromArgb(32, 34, 45)
    BgCardHover    = [System.Drawing.Color]::FromArgb(40, 42, 55)
    BgInput        = [System.Drawing.Color]::FromArgb(22, 22, 32)
    BgMini         = [System.Drawing.Color]::FromArgb(16, 16, 22)
    AccentGreen    = [System.Drawing.Color]::FromArgb(46, 213, 115)
    AccentGreenDim = [System.Drawing.Color]::FromArgb(30, 140, 75)
    AccentBlue     = [System.Drawing.Color]::FromArgb(72, 152, 255)
    AccentBlueDim  = [System.Drawing.Color]::FromArgb(50, 110, 200)
    AccentPurple   = [System.Drawing.Color]::FromArgb(168, 130, 255)
    AccentOrange   = [System.Drawing.Color]::FromArgb(255, 165, 50)
    AccentRed      = [System.Drawing.Color]::FromArgb(255, 71, 87)
    TextPrimary    = [System.Drawing.Color]::FromArgb(240, 240, 245)
    TextSecondary  = [System.Drawing.Color]::FromArgb(160, 165, 180)
    TextMuted      = [System.Drawing.Color]::FromArgb(90, 95, 110)
    TextDisabled   = [System.Drawing.Color]::FromArgb(60, 62, 75)
    Border         = [System.Drawing.Color]::FromArgb(45, 48, 62)
    BorderLight    = [System.Drawing.Color]::FromArgb(55, 58, 72)
    GraphBg        = [System.Drawing.Color]::FromArgb(22, 24, 32)
    GraphGrid      = [System.Drawing.Color]::FromArgb(38, 40, 52)
}

$fontFamily = "Segoe UI"
$monoFont = "Cascadia Code"
try {
    $testFont = New-Object System.Drawing.Font($monoFont, 10)
    $testFont.Dispose()
} catch {
    $monoFont = "Consolas"
}

# ============================================
# PRE-CACHED GDI RESOURCES
# ============================================
$script:fonts = @{
    TitleLarge   = New-Object System.Drawing.Font($fontFamily, 16, [System.Drawing.FontStyle]::Bold)
    Subtitle     = New-Object System.Drawing.Font($fontFamily, 8.5)
    LabelBold    = New-Object System.Drawing.Font($fontFamily, 8, [System.Drawing.FontStyle]::Bold)
    SpeedLarge   = New-Object System.Drawing.Font($monoFont, 28, [System.Drawing.FontStyle]::Bold)
    UnitMedium   = New-Object System.Drawing.Font($fontFamily, 10, [System.Drawing.FontStyle]::Bold)
    PeakSmall    = New-Object System.Drawing.Font($fontFamily, 7.5)
    ScaleTiny    = New-Object System.Drawing.Font($fontFamily, 7)
    LegendTiny   = New-Object System.Drawing.Font($fontFamily, 7)
    StatLabel    = New-Object System.Drawing.Font($fontFamily, 6.5, [System.Drawing.FontStyle]::Bold)
    StatValue    = New-Object System.Drawing.Font($monoFont, 9, [System.Drawing.FontStyle]::Bold)
    StatIcon     = New-Object System.Drawing.Font($fontFamily, 9)
    IconEmoji    = New-Object System.Drawing.Font($fontFamily, 11)
    MiniLabel    = New-Object System.Drawing.Font($fontFamily, 6)
    MiniSpeed    = New-Object System.Drawing.Font($monoFont, 11, [System.Drawing.FontStyle]::Bold)
    MiniClose    = New-Object System.Drawing.Font($fontFamily, 8, [System.Drawing.FontStyle]::Bold)
    BtnFont      = New-Object System.Drawing.Font($fontFamily, 8.5)
    HintFont     = New-Object System.Drawing.Font($fontFamily, 7.5)
    DropdownFont = New-Object System.Drawing.Font($fontFamily, 9)
}

$script:brushes = @{
    TextPrimary   = New-Object System.Drawing.SolidBrush($theme.TextPrimary)
    TextSecondary = New-Object System.Drawing.SolidBrush($theme.TextSecondary)
    TextMuted     = New-Object System.Drawing.SolidBrush($theme.TextMuted)
    AccentGreen   = New-Object System.Drawing.SolidBrush($theme.AccentGreen)
    AccentBlue    = New-Object System.Drawing.SolidBrush($theme.AccentBlue)
    AccentPurple  = New-Object System.Drawing.SolidBrush($theme.AccentPurple)
    AccentRed     = New-Object System.Drawing.SolidBrush($theme.AccentRed)
    GraphBg       = New-Object System.Drawing.SolidBrush($theme.GraphBg)
    BgCard        = New-Object System.Drawing.SolidBrush($theme.BgCard)
}

$script:pens = @{
    Border  = New-Object System.Drawing.Pen($theme.Border, 1)
    GridDot = New-Object System.Drawing.Pen($theme.GraphGrid, 1)
}
$script:pens.GridDot.DashStyle = 'Dot'

# ============================================
# CONFIGURATION
# ============================================
$config = @{
    UpdateIntervalMs = 1000
    GraphHistoryLen  = 60
    SmoothingAlpha   = 0.35
    MiniOpacity      = 0.95
}

# ============================================
# STATE
# ============================================
$state = @{
    PrevBytesRx   = [long]0
    PrevBytesTx   = [long]0
    TotalRx       = [double]0
    TotalTx       = [double]0
    LastTime      = [DateTime]::Now
    SmoothedDl    = [double]0
    SmoothedUl    = [double]0
    PeakDl        = [double]0
    PeakUl        = [double]0
    DlHistory     = [System.Collections.Generic.List[double]]::new()
    UlHistory     = [System.Collections.Generic.List[double]]::new()
    SelectedNicId = $null
    SessionStart  = [DateTime]::Now
    IsConnected   = $true
}

for ($i = 0; $i -lt $config.GraphHistoryLen; $i++) {
    $state.DlHistory.Add(0)
    $state.UlHistory.Add(0)
}

# Pre-allocated graph buffers
$script:dlPointBuffer = [System.Drawing.PointF[]]::new($config.GraphHistoryLen + 2)
$script:ulPointBuffer = [System.Drawing.PointF[]]::new($config.GraphHistoryLen + 2)

# Previous display values for dirty checking
$script:prevDlDisplay = ""
$script:prevUlDisplay = ""
$script:prevStatsRx = ""
$script:prevStatsTx = ""
$script:prevUptime = ""
$script:prevStatus = ""

# ============================================
# HELPER FUNCTIONS
# ============================================
function Format-Speed([double]$bps) {
    if ($bps -ge 1GB) { "{0:N2} GB/s" -f ($bps / 1GB) }
    elseif ($bps -ge 1MB) { "{0:N2} MB/s" -f ($bps / 1MB) }
    elseif ($bps -ge 1KB) { "{0:N2} KB/s" -f ($bps / 1KB) }
    else { "{0:N0} B/s" -f $bps }
}

function Format-SpeedShort([double]$bps) {
    if ($bps -ge 1GB) { "{0:N1}G" -f ($bps / 1GB) }
    elseif ($bps -ge 1MB) { "{0:N1}M" -f ($bps / 1MB) }
    elseif ($bps -ge 1KB) { "{0:N1}K" -f ($bps / 1KB) }
    else { "{0:N0}B" -f $bps }
}

function Format-Size([double]$b) {
    if ($b -ge 1GB) { "{0:N2} GB" -f ($b / 1GB) }
    elseif ($b -ge 1MB) { "{0:N2} MB" -f ($b / 1MB) }
    elseif ($b -ge 1KB) { "{0:N2} KB" -f ($b / 1KB) }
    else { "{0:N0} B" -f $b }
}

function Split-Speed([double]$bps) {
    if ($bps -ge 1GB) { @{ Value = "{0:N2}" -f ($bps / 1GB); Unit = "GB/s" } }
    elseif ($bps -ge 1MB) { @{ Value = "{0:N2}" -f ($bps / 1MB); Unit = "MB/s" } }
    elseif ($bps -ge 1KB) { @{ Value = "{0:N2}" -f ($bps / 1KB); Unit = "KB/s" } }
    else { @{ Value = "{0:N0}" -f $bps; Unit = "B/s" } }
}

function Get-SpeedTierColor([double]$bps, [System.Drawing.Color]$base) {
    if ($bps -ge 100MB) { return $theme.AccentRed }
    if ($bps -ge 50MB) { return $theme.AccentOrange }
    if ($bps -ge 10MB) { return $theme.AccentPurple }
    return $base
}

function Get-ActiveNics {
    [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() |
        Where-Object {
            $_.OperationalStatus -eq 'Up' -and
            $_.NetworkInterfaceType -ne 'Loopback' -and
            $_.NetworkInterfaceType -ne 'Tunnel' -and
            $_.Description -notmatch 'Virtual|Pseudo|WAN Miniport|Teredo|isatap' -and
            $_.GetIPv4Statistics().BytesReceived -gt 0
        } | Sort-Object { $_.GetIPv4Statistics().BytesReceived } -Descending
}

function Init-Nic($nic) {
    if (-not $nic) { return }
    $s = $nic.GetIPv4Statistics()
    $state.PrevBytesRx = $s.BytesReceived
    $state.PrevBytesTx = $s.BytesSent
    $state.LastTime = [DateTime]::Now
    $state.SelectedNicId = $nic.Id
    $state.SmoothedDl = 0
    $state.SmoothedUl = 0
}

function Draw-RoundedRect($g, $rect, $radius, $color) {
    $path = $null
    $brush = $null
    try {
        $d = [Math]::Max($radius * 2, 1)
        if ($rect.Width -lt $d -or $rect.Height -lt $d) {
            $brush = New-Object System.Drawing.SolidBrush($color)
            $g.FillRectangle($brush, $rect)
            return
        }
        $path = New-Object System.Drawing.Drawing2D.GraphicsPath
        $path.AddArc($rect.X, $rect.Y, $d, $d, 180, 90)
        $path.AddArc(($rect.Right - $d), $rect.Y, $d, $d, 270, 90)
        $path.AddArc(($rect.Right - $d), ($rect.Bottom - $d), $d, $d, 0, 90)
        $path.AddArc($rect.X, ($rect.Bottom - $d), $d, $d, 90, 90)
        $path.CloseFigure()
        $brush = New-Object System.Drawing.SolidBrush($color)
        $g.FillPath($brush, $path)
    }
    finally {
        if ($brush) { $brush.Dispose() }
        if ($path) { $path.Dispose() }
    }
}

function Draw-RoundedBorder($g, $rect, $radius, $color, $width) {
    $path = $null
    $pen = $null
    try {
        $d = [Math]::Max($radius * 2, 1)
        if ($rect.Width -lt $d -or $rect.Height -lt $d) {
            $pen = New-Object System.Drawing.Pen($color, $width)
            $g.DrawRectangle($pen, $rect)
            return
        }
        $path = New-Object System.Drawing.Drawing2D.GraphicsPath
        $path.AddArc($rect.X, $rect.Y, $d, $d, 180, 90)
        $path.AddArc(($rect.Right - $d), $rect.Y, $d, $d, 270, 90)
        $path.AddArc(($rect.Right - $d), ($rect.Bottom - $d), $d, $d, 0, 90)
        $path.AddArc($rect.X, ($rect.Bottom - $d), $d, $d, 90, 90)
        $path.CloseFigure()
        $pen = New-Object System.Drawing.Pen($color, $width)
        $g.DrawPath($pen, $path)
    }
    finally {
        if ($pen) { $pen.Dispose() }
        if ($path) { $path.Dispose() }
    }
}

function Draw-Icon-Download($g, $x, $y, $size, $color) {
    $pen = $null
    try {
        $pen = New-Object System.Drawing.Pen($color, 2.5)
        $pen.StartCap = 'Round'; $pen.EndCap = 'Round'
        $cx = $x + $size / 2
        $g.DrawLine($pen, $cx, ($y + 3), $cx, ($y + $size - 6))
        $g.DrawLine($pen, ($cx - 5), ($y + $size - 10), $cx, ($y + $size - 5))
        $g.DrawLine($pen, ($cx + 5), ($y + $size - 10), $cx, ($y + $size - 5))
        $g.DrawLine($pen, ($x + 4), ($y + $size - 2), ($x + $size - 4), ($y + $size - 2))
    }
    finally {
        if ($pen) { $pen.Dispose() }
    }
}

function Draw-Icon-Upload($g, $x, $y, $size, $color) {
    $pen = $null
    try {
        $pen = New-Object System.Drawing.Pen($color, 2.5)
        $pen.StartCap = 'Round'; $pen.EndCap = 'Round'
        $cx = $x + $size / 2
        $g.DrawLine($pen, $cx, ($y + $size - 5), $cx, ($y + 5))
        $g.DrawLine($pen, ($cx - 5), ($y + 9), $cx, ($y + 4))
        $g.DrawLine($pen, ($cx + 5), ($y + 9), $cx, ($y + 4))
        $g.DrawLine($pen, ($x + 4), ($y + $size - 2), ($x + $size - 4), ($y + $size - 2))
    }
    finally {
        if ($pen) { $pen.Dispose() }
    }
}

function Draw-SpeedBar($g, $rect, $value, $maxValue, $color) {
    $bgColor = [System.Drawing.Color]::FromArgb(25, $color.R, $color.G, $color.B)
    Draw-RoundedRect $g $rect 4 $bgColor
    if ($maxValue -gt 0 -and $value -gt 0) {
        $ratio = [Math]::Min(1.0, $value / $maxValue)
        $fillWidth = [int]($rect.Width * $ratio)
        if ($fillWidth -gt 8) {
            $fillRect = New-Object System.Drawing.Rectangle($rect.X, $rect.Y, $fillWidth, $rect.Height)
            $fillColor = [System.Drawing.Color]::FromArgb(120, $color.R, $color.G, $color.B)
            Draw-RoundedRect $g $fillRect 4 $fillColor
        }
    }
}

function Paint-SpeedCard($g, $w, $h, $label, $speed, $unit, $peak, $color, $barRatio, $isDownload) {
    try {
        Draw-RoundedRect $g (New-Object System.Drawing.Rectangle(0, 0, $w, $h)) 12 $theme.BgCard

        $glowColor = [System.Drawing.Color]::FromArgb(8, $color.R, $color.G, $color.B)
        $glowBrush = New-Object System.Drawing.SolidBrush($glowColor)
        try { $g.FillRectangle($glowBrush, 0, 0, $w, 3) }
        finally { $glowBrush.Dispose() }

        Draw-RoundedRect $g (New-Object System.Drawing.Rectangle(12, 0, ($w - 24), 3)) 1 $color

        if ($isDownload) { Draw-Icon-Download $g 14 18 28 $color }
        else { Draw-Icon-Upload $g 14 18 28 $color }

        $g.DrawString($label, $script:fonts.LabelBold, $script:brushes.TextMuted, 48, 18)

        $speedBrush = New-Object System.Drawing.SolidBrush($color)
        try { $g.DrawString($speed, $script:fonts.SpeedLarge, $speedBrush, 10, 48) }
        finally { $speedBrush.Dispose() }

        $unitColor = [System.Drawing.Color]::FromArgb(120, $color.R, $color.G, $color.B)
        $unitBrush = New-Object System.Drawing.SolidBrush($unitColor)
        try { $g.DrawString($unit, $script:fonts.UnitMedium, $unitBrush, 165, 70) }
        finally { $unitBrush.Dispose() }

        $barRect = New-Object System.Drawing.Rectangle(12, 100, ($w - 24), 6)
        Draw-SpeedBar $g $barRect $barRatio 1.0 $color

        $peakText = [char]0x25B2 + " Peak: $peak"
        $g.DrawString($peakText, $script:fonts.PeakSmall, $script:brushes.TextMuted, 12, 115)

        Draw-RoundedBorder $g (New-Object System.Drawing.Rectangle(0, 0, ($w - 1), ($h - 1))) 12 $theme.Border 1
    }
    catch {
        Write-Debug "SpeedCard paint error: $_"
    }
}

# ============================================
# TRAY ICON
# ============================================
function New-TrayIcon {
    $bmp = $null
    $g = $null
    try {
        $bmp = New-Object System.Drawing.Bitmap(32, 32)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.SmoothingMode = 'HighQuality'
        $g.InterpolationMode = 'HighQualityBicubic'
        $g.Clear([System.Drawing.Color]::Transparent)

        $bgBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(220, 26, 26, 36))
        try { $g.FillEllipse($bgBrush, 1, 1, 30, 30) }
        finally { $bgBrush.Dispose() }

        $dlPen = New-Object System.Drawing.Pen($theme.AccentGreen, 2.5)
        $dlPen.StartCap = 'Round'; $dlPen.EndCap = 'Round'
        try {
            $g.DrawLine($dlPen, 10, 7, 10, 20)
            $g.DrawLine($dlPen, 6, 16, 10, 21)
            $g.DrawLine($dlPen, 14, 16, 10, 21)
        }
        finally { $dlPen.Dispose() }

        $ulPen = New-Object System.Drawing.Pen($theme.AccentBlue, 2.5)
        $ulPen.StartCap = 'Round'; $ulPen.EndCap = 'Round'
        try {
            $g.DrawLine($ulPen, 22, 25, 22, 12)
            $g.DrawLine($ulPen, 18, 16, 22, 11)
            $g.DrawLine($ulPen, 26, 16, 22, 11)
        }
        finally { $ulPen.Dispose() }

        $hIcon = $bmp.GetHicon()
        $icon = [System.Drawing.Icon]::FromHandle($hIcon).Clone()
        [Win32UI]::DestroyIcon($hIcon) | Out-Null
        return $icon
    }
    finally {
        if ($g) { $g.Dispose() }
        if ($bmp) { $bmp.Dispose() }
    }
}

# ============================================
# MAIN FORM
# ============================================
$form = New-Object DBForm
$form.Text = "NetFlow Monitor"
$form.ClientSize = New-Object System.Drawing.Size(500, 620)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.BackColor = $theme.BgPrimary
$form.ForeColor = $theme.TextPrimary
$form.KeyPreview = $true
$form.Font = New-Object System.Drawing.Font($fontFamily, 9)

try {
    $val = 1
    [Win32UI]::DwmSetWindowAttribute($form.Handle, 20, [ref]$val, 4) | Out-Null
} catch {}

# ============================================
# HEADER PANEL
# ============================================
$headerPanel = New-Object DBPanel
$headerPanel.Dock = "Top"
$headerPanel.Height = 70
$headerPanel.BackColor = $theme.BgSecondary

$headerPanel.Add_Paint({
    param($s, $e)
    $g = $e.Graphics
    $g.SmoothingMode = 'HighQuality'
    $g.TextRenderingHint = 'ClearTypeGridFit'

    try {
        $gradBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            (New-Object System.Drawing.Point(0, ($s.Height - 2))),
            (New-Object System.Drawing.Point($s.Width, ($s.Height - 2))),
            $theme.AccentGreen, $theme.AccentBlue
        )
        try { $g.FillRectangle($gradBrush, 0, ($s.Height - 2), $s.Width, 2) }
        finally { $gradBrush.Dispose() }

        $iconX = 20; $iconY = 18
        $barColors = @($theme.AccentGreen, $theme.AccentGreen, $theme.AccentBlue, $theme.AccentBlue)
        $barHeights = @(10, 18, 26, 34)
        for ($i = 0; $i -lt 4; $i++) {
            $bx = $iconX + ($i * 8)
            $by = $iconY + (34 - $barHeights[$i])
            Draw-RoundedRect $g (New-Object System.Drawing.Rectangle($bx, $by, 5, $barHeights[$i])) 2 $barColors[$i]
        }

        $g.DrawString("NetFlow Monitor", $script:fonts.TitleLarge, $script:brushes.TextPrimary, 60, 12)
        $g.DrawString("Real-time network speed monitoring", $script:fonts.Subtitle, $script:brushes.TextMuted, 62, 40)
    }
    catch {
        Write-Debug "Header paint error: $_"
    }
})

$form.Controls.Add($headerPanel)

# ============================================
# ADAPTER SELECTOR
# ============================================
$adapterPanel = New-Object DBPanel
$adapterPanel.Location = New-Object System.Drawing.Point(16, 82)
$adapterPanel.Size = New-Object System.Drawing.Size(468, 42)
$adapterPanel.BackColor = $theme.BgCard

$adapterPanel.Add_Paint({
    param($s, $e)
    $g = $e.Graphics
    $g.SmoothingMode = 'HighQuality'
    $g.TextRenderingHint = 'ClearTypeGridFit'
    try {
        Draw-RoundedBorder $g (New-Object System.Drawing.Rectangle(0, 0, ($s.Width - 1), ($s.Height - 1))) 8 $theme.Border 1
        $g.DrawString("ADAPTER", $script:fonts.Subtitle, $script:brushes.TextMuted, 12, 13)
    }
    catch {
        Write-Debug "Adapter paint error: $_"
    }
})

$nicDropdown = New-Object System.Windows.Forms.ComboBox
$nicDropdown.Font = $script:fonts.DropdownFont
$nicDropdown.BackColor = $theme.BgInput
$nicDropdown.ForeColor = $theme.TextPrimary
$nicDropdown.FlatStyle = "Flat"
$nicDropdown.DropDownStyle = "DropDownList"
$nicDropdown.Location = New-Object System.Drawing.Point(105, 9)
$nicDropdown.Size = New-Object System.Drawing.Size(350, 26)
$adapterPanel.Controls.Add($nicDropdown)

$script:nicMap = @{}
$allNics = Get-ActiveNics
$idx = 0
foreach ($nic in $allNics) {
    $speed = ""
    if ($nic.Speed -ge 1000000000) { $speed = " [{0:N0} Gbps]" -f ($nic.Speed / 1000000000) }
    elseif ($nic.Speed -ge 1000000) { $speed = " [{0:N0} Mbps]" -f ($nic.Speed / 1000000) }
    $nicDropdown.Items.Add("$($nic.Name)$speed") | Out-Null
    $script:nicMap[$idx] = $nic.Id
    $idx++
}
if ($nicDropdown.Items.Count -gt 0) { $nicDropdown.SelectedIndex = 0 }

$selectedNIC = $allNics | Select-Object -First 1
if (-not $selectedNIC) {
    [System.Windows.Forms.MessageBox]::Show(
        "No active network adapter detected.`nPlease check your connection.",
        "NetFlow Monitor", "OK", "Error")
    return
}
Init-Nic $selectedNIC

$nicDropdown.Add_SelectedIndexChanged({
    $idx = $nicDropdown.SelectedIndex
    if ($idx -ge 0 -and $script:nicMap.ContainsKey($idx)) {
        $targetId = $script:nicMap[$idx]
        $nic = [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() |
            Where-Object { $_.Id -eq $targetId } | Select-Object -First 1
        if ($nic) {
            $script:selectedNIC = $nic
            Init-Nic $nic
            $state.TotalRx = 0; $state.TotalTx = 0
            $state.PeakDl = 0; $state.PeakUl = 0
            $state.SessionStart = [DateTime]::Now
            for ($i = 0; $i -lt $config.GraphHistoryLen; $i++) {
                $state.DlHistory[$i] = 0
                $state.UlHistory[$i] = 0
            }
        }
    }
})

$form.Controls.Add($adapterPanel)

# ============================================
# SPEED CARDS
# ============================================
$script:dlDisplaySpeed = "0.00"
$script:dlDisplayUnit = "KB/s"
$script:dlDisplayPeak = "0.00 KB/s"
$script:dlDisplayColor = $theme.AccentGreen
$script:dlBarRatio = 0.0

$script:ulDisplaySpeed = "0.00"
$script:ulDisplayUnit = "KB/s"
$script:ulDisplayPeak = "0.00 KB/s"
$script:ulDisplayColor = $theme.AccentBlue
$script:ulBarRatio = 0.0

# --- DOWNLOAD CARD ---
$dlCard = New-Object DBPanel
$dlCard.Location = New-Object System.Drawing.Point(16, 134)
$dlCard.Size = New-Object System.Drawing.Size(226, 140)
$dlCard.BackColor = $theme.BgCard

$dlCard.Add_Paint({
    param($s, $e)
    $e.Graphics.SmoothingMode = 'HighQuality'
    $e.Graphics.TextRenderingHint = 'ClearTypeGridFit'
    Paint-SpeedCard $e.Graphics $s.Width $s.Height "DOWNLOAD" `
        $script:dlDisplaySpeed $script:dlDisplayUnit $script:dlDisplayPeak `
        $script:dlDisplayColor $script:dlBarRatio $true
})

$form.Controls.Add($dlCard)

# --- UPLOAD CARD ---
$ulCard = New-Object DBPanel
$ulCard.Location = New-Object System.Drawing.Point(258, 134)
$ulCard.Size = New-Object System.Drawing.Size(226, 140)
$ulCard.BackColor = $theme.BgCard

$ulCard.Add_Paint({
    param($s, $e)
    $e.Graphics.SmoothingMode = 'HighQuality'
    $e.Graphics.TextRenderingHint = 'ClearTypeGridFit'
    Paint-SpeedCard $e.Graphics $s.Width $s.Height "UPLOAD" `
        $script:ulDisplaySpeed $script:ulDisplayUnit $script:ulDisplayPeak `
        $script:ulDisplayColor $script:ulBarRatio $false
})

$form.Controls.Add($ulCard)

# ============================================
# GRAPH PANEL
# ============================================
$graphCard = New-Object DBPanel
$graphCard.Location = New-Object System.Drawing.Point(16, 286)
$graphCard.Size = New-Object System.Drawing.Size(468, 175)
$graphCard.BackColor = $theme.BgCard

$graphCard.Add_Paint({
    param($s, $e)
    $g = $e.Graphics
    $g.SmoothingMode = 'HighQuality'
    $g.TextRenderingHint = 'ClearTypeGridFit'
    $w = $s.Width; $h = $s.Height

    try {
        Draw-RoundedRect $g (New-Object System.Drawing.Rectangle(0, 0, $w, $h)) 12 $theme.BgCard
        Draw-RoundedBorder $g (New-Object System.Drawing.Rectangle(0, 0, ($w - 1), ($h - 1))) 12 $theme.Border 1

        $g.DrawString("SPEED HISTORY", $script:fonts.LabelBold, $script:brushes.TextMuted, 14, 10)

        # Legend
        $g.FillRectangle($script:brushes.AccentGreen, ($w - 155), 11, 8, 8)
        $g.DrawString("Download", $script:fonts.LegendTiny, $script:brushes.AccentGreen, ($w - 143), 9)
        $g.FillRectangle($script:brushes.AccentBlue, ($w - 75), 11, 8, 8)
        $g.DrawString("Upload", $script:fonts.LegendTiny, $script:brushes.AccentBlue, ($w - 63), 9)

        # Graph area
        $gx = 14; $gy = 30; $gw = $w - 28; $gh = $h - 45
        Draw-RoundedRect $g (New-Object System.Drawing.Rectangle($gx, $gy, $gw, $gh)) 6 $theme.GraphBg

        # Grid lines
        for ($gi = 1; $gi -le 3; $gi++) {
            $ly = $gy + [int]($gh * $gi / 4)
            $g.DrawLine($script:pens.GridDot, ($gx + 4), $ly, ($gx + $gw - 4), $ly)
        }

        # Calculate scale
        $maxVal = 1024.0
        for ($di = 0; $di -lt $state.DlHistory.Count; $di++) {
            if ($state.DlHistory[$di] -gt $maxVal) { $maxVal = $state.DlHistory[$di] }
            if ($state.UlHistory[$di] -gt $maxVal) { $maxVal = $state.UlHistory[$di] }
        }
        $maxVal *= 1.15

        $ptCount = $state.DlHistory.Count
        if ($ptCount -lt 2) { return }
        $stepX = $gw / [Math]::Max($ptCount - 1, 1)

        # Fill point buffers
        for ($pi = 0; $pi -lt $ptCount; $pi++) {
            $px = $gx + ($pi * $stepX)

            $dlPy = $gy + $gh - (($state.DlHistory[$pi] / $maxVal) * $gh)
            if ($dlPy -lt ($gy + 2)) { $dlPy = $gy + 2 }
            if ($dlPy -gt ($gy + $gh - 2)) { $dlPy = $gy + $gh - 2 }
            $script:dlPointBuffer[$pi] = [System.Drawing.PointF]::new($px, $dlPy)

            $ulPy = $gy + $gh - (($state.UlHistory[$pi] / $maxVal) * $gh)
            if ($ulPy -lt ($gy + 2)) { $ulPy = $gy + 2 }
            if ($ulPy -gt ($gy + $gh - 2)) { $ulPy = $gy + $gh - 2 }
            $script:ulPointBuffer[$pi] = [System.Drawing.PointF]::new($px, $ulPy)
        }

        $script:dlPointBuffer[$ptCount] = [System.Drawing.PointF]::new(($gx + $gw), ($gy + $gh))
        $script:dlPointBuffer[$ptCount + 1] = [System.Drawing.PointF]::new($gx, ($gy + $gh))
        $script:ulPointBuffer[$ptCount] = [System.Drawing.PointF]::new(($gx + $gw), ($gy + $gh))
        $script:ulPointBuffer[$ptCount + 1] = [System.Drawing.PointF]::new($gx, ($gy + $gh))

        $totalPts = $ptCount + 2

        # Download fill
        $dlFillBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            (New-Object System.Drawing.Point(0, $gy)),
            (New-Object System.Drawing.Point(0, ($gy + $gh))),
            [System.Drawing.Color]::FromArgb(40, 46, 213, 115),
            [System.Drawing.Color]::FromArgb(2, 46, 213, 115)
        )
        $fillPath = New-Object System.Drawing.Drawing2D.GraphicsPath
        try {
            $polyDl = $script:dlPointBuffer[0..($totalPts - 1)]
            $fillPath.AddPolygon($polyDl)
            $g.FillPath($dlFillBrush, $fillPath)
        }
        finally {
            $dlFillBrush.Dispose()
            $fillPath.Dispose()
        }

        # Download line
        $dlLineArr = $script:dlPointBuffer[0..($ptCount - 1)]
        if ($dlLineArr.Count -ge 2) {
            $dlPen = New-Object System.Drawing.Pen($theme.AccentGreen, 2)
            $dlPen.LineJoin = 'Round'
            try { $g.DrawLines($dlPen, $dlLineArr) }
            finally { $dlPen.Dispose() }
        }

        # Upload fill
        $ulFillBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            (New-Object System.Drawing.Point(0, $gy)),
            (New-Object System.Drawing.Point(0, ($gy + $gh))),
            [System.Drawing.Color]::FromArgb(30, 72, 152, 255),
            [System.Drawing.Color]::FromArgb(2, 72, 152, 255)
        )
        $fillPath2 = New-Object System.Drawing.Drawing2D.GraphicsPath
        try {
            $polyUl = $script:ulPointBuffer[0..($totalPts - 1)]
            $fillPath2.AddPolygon($polyUl)
            $g.FillPath($ulFillBrush, $fillPath2)
        }
        finally {
            $ulFillBrush.Dispose()
            $fillPath2.Dispose()
        }

        # Upload line
        $ulLineArr = $script:ulPointBuffer[0..($ptCount - 1)]
        if ($ulLineArr.Count -ge 2) {
            $ulPen = New-Object System.Drawing.Pen($theme.AccentBlue, 2)
            $ulPen.LineJoin = 'Round'
            try { $g.DrawLines($ulPen, $ulLineArr) }
            finally { $ulPen.Dispose() }
        }

        # Scale label
        $g.DrawString("Max: $(Format-Speed $maxVal)", $script:fonts.ScaleTiny, $script:brushes.TextMuted, ($gx + 5), ($gy + 3))
        $g.DrawString("60s ago", $script:fonts.ScaleTiny, $script:brushes.TextMuted, ($gx + 2), ($gy + $gh + 2))
        $g.DrawString("now", $script:fonts.ScaleTiny, $script:brushes.TextMuted, ($gx + $gw - 22), ($gy + $gh + 2))
    }
    catch {
        Write-Debug "Graph paint error: $_"
    }
})

$form.Controls.Add($graphCard)

# ============================================
# STATS PANEL
# ============================================
$statsPanel = New-Object DBPanel
$statsPanel.Location = New-Object System.Drawing.Point(16, 472)
$statsPanel.Size = New-Object System.Drawing.Size(468, 80)
$statsPanel.BackColor = $theme.BgCard

$script:statsTotalRx = "0 B"
$script:statsTotalTx = "0 B"
$script:statsUptime = "00:00:00"
$script:statsStatus = "Connected"

$statsPanel.Add_Paint({
    param($s, $e)
    $g = $e.Graphics
    $g.SmoothingMode = 'HighQuality'
    $g.TextRenderingHint = 'ClearTypeGridFit'
    $w = $s.Width; $h = $s.Height

    try {
        Draw-RoundedRect $g (New-Object System.Drawing.Rectangle(0, 0, $w, $h)) 12 $theme.BgCard
        Draw-RoundedBorder $g (New-Object System.Drawing.Rectangle(0, 0, ($w - 1), ($h - 1))) 12 $theme.Border 1

        $colW = [int]($w / 4)
        $statusColor = if ($state.IsConnected) { $theme.AccentGreen } else { $theme.AccentRed }

        $statItems = @(
            @{ Label = "DOWNLOADED"; Value = $script:statsTotalRx; Color = $theme.AccentGreen; Icon = "DL" },
            @{ Label = "UPLOADED";   Value = $script:statsTotalTx; Color = $theme.AccentBlue;  Icon = "UL" },
            @{ Label = "UPTIME";     Value = $script:statsUptime;  Color = $theme.AccentPurple; Icon = "UP" },
            @{ Label = "STATUS";     Value = $script:statsStatus;  Color = $statusColor;        Icon = "ST" }
        )

        for ($si = 0; $si -lt 4; $si++) {
            $item = $statItems[$si]
            $cx = ($si * $colW) + ($colW / 2)

            if ($si -gt 0) {
                $g.DrawLine($script:pens.Border, ($si * $colW), 15, ($si * $colW), ($h - 15))
            }

            # Icon text
            $iconBrush = New-Object System.Drawing.SolidBrush($item.Color)
            try {
                $iconSize = $g.MeasureString($item.Icon, $script:fonts.StatIcon)
                $g.DrawString($item.Icon, $script:fonts.StatIcon, $iconBrush, ($cx - $iconSize.Width / 2), 10)
            }
            finally { $iconBrush.Dispose() }

            # Label
            $lblSize = $g.MeasureString($item.Label, $script:fonts.StatLabel)
            $g.DrawString($item.Label, $script:fonts.StatLabel, $script:brushes.TextMuted, ($cx - $lblSize.Width / 2), 28)

            # Value
            $valSize = $g.MeasureString($item.Value, $script:fonts.StatValue)
            $g.DrawString($item.Value, $script:fonts.StatValue, $script:brushes.TextPrimary, ($cx - $valSize.Width / 2), 46)
        }
    }
    catch {
        Write-Debug "Stats paint error: $_"
    }
})

$form.Controls.Add($statsPanel)

# ============================================
# BOTTOM BAR
# ============================================
$bottomPanel = New-Object DBPanel
$bottomPanel.Location = New-Object System.Drawing.Point(16, 564)
$bottomPanel.Size = New-Object System.Drawing.Size(468, 44)
$bottomPanel.BackColor = $theme.BgPrimary

# Reset button
$resetBtn = New-Object DBPanel
$resetBtn.Location = New-Object System.Drawing.Point(0, 4)
$resetBtn.Size = New-Object System.Drawing.Size(110, 34)
$resetBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
$script:resetHover = $false

$resetBtn.Add_Paint({
    param($s, $e)
    $g = $e.Graphics
    $g.SmoothingMode = 'HighQuality'
    $g.TextRenderingHint = 'ClearTypeGridFit'
    try {
        $bgColor = if ($script:resetHover) { $theme.BgCardHover } else { $theme.BgCard }
        Draw-RoundedRect $g (New-Object System.Drawing.Rectangle(0, 0, $s.Width, $s.Height)) 8 $bgColor
        Draw-RoundedBorder $g (New-Object System.Drawing.Rectangle(0, 0, ($s.Width - 1), ($s.Height - 1))) 8 $theme.Border 1
        $g.DrawString("Reset Stats", $script:fonts.BtnFont, $script:brushes.TextSecondary, 16, 8)
    }
    catch {
        Write-Debug "Reset button paint error: $_"
    }
})

$resetBtn.Add_MouseEnter({ $script:resetHover = $true; $resetBtn.Invalidate() })
$resetBtn.Add_MouseLeave({ $script:resetHover = $false; $resetBtn.Invalidate() })
$resetBtn.Add_Click({
    $state.TotalRx = 0; $state.TotalTx = 0
    $state.PeakDl = 0; $state.PeakUl = 0
    $state.SessionStart = [DateTime]::Now
    for ($i = 0; $i -lt $config.GraphHistoryLen; $i++) {
        $state.DlHistory[$i] = 0
        $state.UlHistory[$i] = 0
    }
})
$bottomPanel.Controls.Add($resetBtn)

$hintLabel = New-Object System.Windows.Forms.Label
$hintLabel.Text = "Press Esc or minimize for floating widget"
$hintLabel.Font = $script:fonts.HintFont
$hintLabel.ForeColor = $theme.TextMuted
$hintLabel.AutoSize = $true
$hintLabel.Location = New-Object System.Drawing.Point(175, 12)
$hintLabel.BackColor = $theme.BgPrimary
$bottomPanel.Controls.Add($hintLabel)

$form.Controls.Add($bottomPanel)

# ============================================
# MINI OVERLAY
# ============================================
$miniForm = New-Object DBForm
$miniForm.FormBorderStyle = "None"
$miniForm.Size = New-Object System.Drawing.Size(200, 70)
$miniForm.TopMost = $true
$miniForm.ShowInTaskbar = $false
$miniForm.BackColor = $theme.BgMini
$miniForm.Opacity = $config.MiniOpacity
$miniForm.Visible = $false
$miniForm.StartPosition = "Manual"

$miniForm.Add_Load({
    $rgn = [Win32UI]::CreateRoundRectRgn(0, 0, $miniForm.Width, $miniForm.Height, 16, 16)
    [Win32UI]::SetWindowRgn($miniForm.Handle, $rgn, $true)
    try {
        $val = 1
        [Win32UI]::DwmSetWindowAttribute($miniForm.Handle, 20, [ref]$val, 4) | Out-Null
    } catch {}
})

$screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$miniForm.Location = New-Object System.Drawing.Point(
    ($screen.Right - $miniForm.Width - 12),
    ($screen.Bottom - $miniForm.Height - 8)
)

$script:miniDlText = "0.00 KB/s"
$script:miniUlText = "0.00 KB/s"
$script:miniDlColor = $theme.AccentGreen
$script:miniUlColor = $theme.AccentBlue
$script:miniCloseHover = $false

$miniForm.Add_Paint({
    param($s, $e)
    $g = $e.Graphics
    $g.SmoothingMode = 'HighQuality'
    $g.TextRenderingHint = 'ClearTypeGridFit'
    $w = $s.Width; $h = $s.Height

    try {
        # Gradient background
        $bgBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            (New-Object System.Drawing.Point(0, 0)),
            (New-Object System.Drawing.Point($w, $h)),
            [System.Drawing.Color]::FromArgb(22, 22, 30),
            [System.Drawing.Color]::FromArgb(16, 16, 22)
        )
        try { $g.FillRectangle($bgBrush, 0, 0, $w, $h) }
        finally { $bgBrush.Dispose() }

        # Top accent
        $accentBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            (New-Object System.Drawing.Point(0, 0)),
            (New-Object System.Drawing.Point($w, 0)),
            $theme.AccentGreen, $theme.AccentBlue
        )
        try { $g.FillRectangle($accentBrush, 0, 0, $w, 2) }
        finally { $accentBrush.Dispose() }

        # Dot
        $dotBrush = New-Object System.Drawing.SolidBrush($theme.AccentGreen)
        try { $g.FillEllipse($dotBrush, 10, 10, 6, 6) }
        finally { $dotBrush.Dispose() }

        # Label
        $g.DrawString("NetFlow", $script:fonts.MiniLabel, $script:brushes.TextMuted, 20, 8)

        # Download
        $dlBrush = New-Object System.Drawing.SolidBrush($script:miniDlColor)
        try { $g.DrawString("DL $($script:miniDlText)", $script:fonts.MiniSpeed, $dlBrush, 8, 24) }
        finally { $dlBrush.Dispose() }

        # Upload
        $ulBrush = New-Object System.Drawing.SolidBrush($script:miniUlColor)
        try { $g.DrawString("UL $($script:miniUlText)", $script:fonts.MiniSpeed, $ulBrush, 8, 46) }
        finally { $ulBrush.Dispose() }

        # Close
        $closeColor = if ($script:miniCloseHover) { $theme.AccentRed } else { $theme.TextMuted }
        $closeBrush = New-Object System.Drawing.SolidBrush($closeColor)
        try { $g.DrawString("X", $script:fonts.MiniClose, $closeBrush, ($w - 18), 5) }
        finally { $closeBrush.Dispose() }
    }
    catch {
        Write-Debug "Mini paint error: $_"
    }
})

# Mini form dragging
$script:miniDragging = $false
$script:miniDragStart = [System.Drawing.Point]::Empty
$script:miniFormStartLoc = [System.Drawing.Point]::Empty

$miniForm.Add_MouseDown({
    param($s, $e)
    if ($e.Button -eq "Left") {
        if ($e.X -ge ($miniForm.Width - 25) -and $e.Y -le 22) {
            $miniForm.Hide()
            $notifyIcon.Visible = $false
            $form.Show()
            $form.WindowState = "Normal"
            $form.Activate()
            return
        }
        $script:miniDragging = $true
        $script:miniDragStart = [System.Windows.Forms.Cursor]::Position
        $script:miniFormStartLoc = $miniForm.Location
    }
})

$miniForm.Add_MouseMove({
    param($s, $e)
    if ($script:miniDragging) {
        $cur = [System.Windows.Forms.Cursor]::Position
        $miniForm.Location = New-Object System.Drawing.Point(
            ($script:miniFormStartLoc.X + $cur.X - $script:miniDragStart.X),
            ($script:miniFormStartLoc.Y + $cur.Y - $script:miniDragStart.Y)
        )
    }
    $wasHover = $script:miniCloseHover
    $script:miniCloseHover = ($e.X -ge ($miniForm.Width - 25) -and $e.Y -le 22)
    if ($wasHover -ne $script:miniCloseHover) { $miniForm.Invalidate() }
})

$miniForm.Add_MouseUp({ $script:miniDragging = $false })

$miniForm.Add_DoubleClick({
    $miniForm.Hide()
    $notifyIcon.Visible = $false
    $form.Show()
    $form.WindowState = "Normal"
    $form.Activate()
})

# ============================================
# SYSTEM TRAY
# ============================================
$notifyIcon = New-Object System.Windows.Forms.NotifyIcon
$notifyIcon.Icon = New-TrayIcon
$notifyIcon.Text = "NetFlow Monitor"
$notifyIcon.Visible = $false

$trayMenu = New-Object System.Windows.Forms.ContextMenuStrip
$trayMenu.BackColor = [System.Drawing.Color]::FromArgb(35, 35, 45)
$trayMenu.ForeColor = $theme.TextPrimary
$trayMenu.Font = New-Object System.Drawing.Font($fontFamily, 9.5)
$trayMenu.ShowImageMargin = $false

$restoreItem = New-Object System.Windows.Forms.ToolStripMenuItem("Open Full Window")
$restoreItem.Add_Click({
    $miniForm.Hide(); $notifyIcon.Visible = $false
    $form.Show(); $form.WindowState = "Normal"; $form.Activate()
})

$exportItem = New-Object System.Windows.Forms.ToolStripMenuItem("Export Session")
$exportItem.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "CSV Files|*.csv"
    $sfd.FileName = "netflow_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    if ($sfd.ShowDialog() -eq 'OK') {
        $data = for ($i = 0; $i -lt $state.DlHistory.Count; $i++) {
            [PSCustomObject]@{
                SecondAgo   = $state.DlHistory.Count - $i
                DownloadBps = [Math]::Round($state.DlHistory[$i], 2)
                UploadBps   = [Math]::Round($state.UlHistory[$i], 2)
            }
        }
        $data | Export-Csv -Path $sfd.FileName -NoTypeInformation
    }
    $sfd.Dispose()
})

$exitItem = New-Object System.Windows.Forms.ToolStripMenuItem("Exit NetFlow")
$exitItem.Add_Click({
    $timer.Stop()
    $miniForm.Close()
    $notifyIcon.Visible = $false
    $notifyIcon.Dispose()
    $form.Close()
})

$trayMenu.Items.AddRange(@(
    $restoreItem,
    $exportItem,
    (New-Object System.Windows.Forms.ToolStripSeparator),
    $exitItem
))
$notifyIcon.ContextMenuStrip = $trayMenu

$notifyIcon.Add_DoubleClick({
    $miniForm.Hide(); $notifyIcon.Visible = $false
    $form.Show(); $form.WindowState = "Normal"; $form.Activate()
})

# ============================================
# WINDOW BEHAVIOR
# ============================================
$form.Add_Resize({
    if ($form.WindowState -eq "Minimized") {
        $form.Hide()
        $notifyIcon.Visible = $true
        $miniForm.Show()
        $miniForm.Invalidate()
    }
})

$form.Add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq 'Escape') { $form.WindowState = "Minimized" }
})

# ============================================
# UPDATE TIMER
# ============================================
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = $config.UpdateIntervalMs

$timer.Add_Tick({
    try {
        $nic = [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() |
            Where-Object { $_.Id -eq $state.SelectedNicId } | Select-Object -First 1

        if (-not $nic -or $nic.OperationalStatus -ne 'Up') {
            $state.IsConnected = $false
            $script:statsStatus = "Disconnected"

            $fb = Get-ActiveNics | Select-Object -First 1
            if ($fb) {
                $script:selectedNIC = $fb
                Init-Nic $fb
                $state.IsConnected = $true
                $script:statsStatus = "Reconnected"
            }

            if ($script:prevStatus -ne $script:statsStatus) {
                $script:prevStatus = $script:statsStatus
                $statsPanel.Invalidate()
            }
            return
        }

        $state.IsConnected = $true
        $script:statsStatus = "Connected"

        $stats = $nic.GetIPv4Statistics()
        $now = [DateTime]::Now
        $elapsed = ($now - $state.LastTime).TotalSeconds
        if ($elapsed -lt 0.1) { return }

        $currentRx = [long]$stats.BytesReceived
        $currentTx = [long]$stats.BytesSent

        # Handle 32-bit counter wraps
        $rxDiff = $currentRx - $state.PrevBytesRx
        $txDiff = $currentTx - $state.PrevBytesTx
        if ($rxDiff -lt 0) { $rxDiff += [long]([uint32]::MaxValue) + 1 }
        if ($txDiff -lt 0) { $txDiff += [long]([uint32]::MaxValue) + 1 }

        # Sanity cap
        $maxReasonable = 10GB * $elapsed
        if ($rxDiff -gt $maxReasonable) { $rxDiff = 0 }
        if ($txDiff -gt $maxReasonable) { $txDiff = 0 }

        $rawDl = $rxDiff / $elapsed
        $rawUl = $txDiff / $elapsed

        # EMA smoothing
        $a = $config.SmoothingAlpha
        $state.SmoothedDl = ($a * $rawDl) + ((1 - $a) * $state.SmoothedDl)
        $state.SmoothedUl = ($a * $rawUl) + ((1 - $a) * $state.SmoothedUl)

        $dl = $state.SmoothedDl
        $ul = $state.SmoothedUl

        # Peaks
        if ($dl -gt $state.PeakDl) { $state.PeakDl = $dl }
        if ($ul -gt $state.PeakUl) { $state.PeakUl = $ul }

        # Totals
        $state.TotalRx += $rxDiff
        $state.TotalTx += $txDiff

        # History
        $state.DlHistory.RemoveAt(0); $state.DlHistory.Add($dl)
        $state.UlHistory.RemoveAt(0); $state.UlHistory.Add($ul)

        # === Batch UI updates with suspended layout ===
        $form.SuspendLayout()
        try {
            # Download card
            $dlSplit = Split-Speed $dl
            $newDlKey = "$($dlSplit.Value)|$($dlSplit.Unit)"
            if ($newDlKey -ne $script:prevDlDisplay) {
                $script:dlDisplaySpeed = $dlSplit.Value
                $script:dlDisplayUnit = $dlSplit.Unit
                $script:dlDisplayPeak = Format-Speed $state.PeakDl
                $script:dlDisplayColor = Get-SpeedTierColor $dl $theme.AccentGreen
                $maxRef = if ($state.PeakDl -gt 0) { $state.PeakDl } else { 1KB }
                $script:dlBarRatio = [Math]::Min(1.0, $dl / $maxRef)
                $script:prevDlDisplay = $newDlKey
                $dlCard.Invalidate()
            }

            # Upload card
            $ulSplit = Split-Speed $ul
            $newUlKey = "$($ulSplit.Value)|$($ulSplit.Unit)"
            if ($newUlKey -ne $script:prevUlDisplay) {
                $script:ulDisplaySpeed = $ulSplit.Value
                $script:ulDisplayUnit = $ulSplit.Unit
                $script:ulDisplayPeak = Format-Speed $state.PeakUl
                $script:ulDisplayColor = Get-SpeedTierColor $ul $theme.AccentBlue
                $maxRef2 = if ($state.PeakUl -gt 0) { $state.PeakUl } else { 1KB }
                $script:ulBarRatio = [Math]::Min(1.0, $ul / $maxRef2)
                $script:prevUlDisplay = $newUlKey
                $ulCard.Invalidate()
            }

            # Graph always repaints (history shifts)
            $graphCard.Invalidate()

            # Stats
            $newRx = Format-Size $state.TotalRx
            $newTx = Format-Size $state.TotalTx
            $uptime = $now - $state.SessionStart
            $newUptime = "{0:hh\:mm\:ss}" -f $uptime

            if ($newRx -ne $script:prevStatsRx -or
                $newTx -ne $script:prevStatsTx -or
                $newUptime -ne $script:prevUptime -or
                $script:statsStatus -ne $script:prevStatus) {
                $script:statsTotalRx = $newRx
                $script:statsTotalTx = $newTx
                $script:statsUptime = $newUptime
                $script:prevStatsRx = $newRx
                $script:prevStatsTx = $newTx
                $script:prevUptime = $newUptime
                $script:prevStatus = $script:statsStatus
                $statsPanel.Invalidate()
            }

            # Mini overlay
            if ($miniForm.Visible) {
                $script:miniDlText = "$(Format-SpeedShort $dl)/s"
                $script:miniUlText = "$(Format-SpeedShort $ul)/s"
                $script:miniDlColor = Get-SpeedTierColor $dl $theme.AccentGreen
                $script:miniUlColor = Get-SpeedTierColor $ul $theme.AccentBlue
                $miniForm.Invalidate()
            }

            # Tray tooltip
            if ($notifyIcon.Visible) {
                $tt = "DL:$(Format-SpeedShort $dl)/s UL:$(Format-SpeedShort $ul)/s"
                if ($tt.Length -gt 63) { $tt = $tt.Substring(0, 63) }
                $notifyIcon.Text = $tt
            }
        }
        finally {
            $form.ResumeLayout($false)
        }

        # Store for next tick
        $state.PrevBytesRx = $currentRx
        $state.PrevBytesTx = $currentTx
        $state.LastTime = $now
    }
    catch {
        $script:statsStatus = "Error"
        $script:prevStatus = ""
        $statsPanel.Invalidate()
    }
})

# ============================================
# NIC AUTO-REFRESH TIMER
# ============================================
$nicRefreshTimer = New-Object System.Windows.Forms.Timer
$nicRefreshTimer.Interval = 5000

$nicRefreshTimer.Add_Tick({
    try {
        $currentNics = Get-ActiveNics
        $currentIds = @($currentNics | ForEach-Object { $_.Id })
        if ($state.SelectedNicId -and $state.SelectedNicId -notin $currentIds) {
            $fb = $currentNics | Select-Object -First 1
            if ($fb) {
                Init-Nic $fb
                $script:selectedNIC = $fb
                $state.IsConnected = $true
                $script:statsStatus = "Reconnected"
            }
            else {
                $state.IsConnected = $false
                $script:statsStatus = "No Adapter"
            }
            $statsPanel.Invalidate()
        }
    }
    catch {
        Write-Debug "NIC refresh error: $_"
    }
})

# ============================================
# LIFECYCLE
# ============================================
$form.Add_Shown({
    $timer.Start()
    $nicRefreshTimer.Start()
})

$form.Add_FormClosing({
    param($s, $e)

    $timer.Stop(); $timer.Dispose()
    $nicRefreshTimer.Stop(); $nicRefreshTimer.Dispose()
    $miniForm.Close(); $miniForm.Dispose()
    $notifyIcon.Visible = $false; $notifyIcon.Dispose()
    $trayMenu.Dispose()

    # Dispose all cached GDI resources
    foreach ($key in @($script:fonts.Keys)) {
        try { $script:fonts[$key].Dispose() } catch {}
    }
    foreach ($key in @($script:brushes.Keys)) {
        try { $script:brushes[$key].Dispose() } catch {}
    }
    foreach ($key in @($script:pens.Keys)) {
        try { $script:pens[$key].Dispose() } catch {}
    }
})

[System.Windows.Forms.Application]::Run($form)
