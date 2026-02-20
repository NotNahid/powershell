Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ========================
# DOWNLOAD ICON FROM GITHUB
# ========================

$iconBitmap = $null
try {
    $wc = New-Object System.Net.WebClient
    $bytes = $wc.DownloadData("https://avatars.githubusercontent.com/u/218765473?v=4&s=64")
    $ms = New-Object System.IO.MemoryStream(,$bytes)
    $img = [System.Drawing.Image]::FromStream($ms)
    $iconBitmap = New-Object System.Drawing.Bitmap($img, 32, 32)
    $appIcon = [System.Drawing.Icon]::FromHandle($iconBitmap.GetHicon())
    $wc.Dispose()
} catch {
    $appIcon = [System.Drawing.SystemIcons]::Application
}

# ========================
# CONFIG
# ========================

$Name = "Nahid"
$Tagline = "Web Developer | Front-End | Always Learning"
$Bio = @"
A passionate self-taught developer from Bangladesh, interested in creating beautiful and engaging user experiences. Currently diving deeper into the React ecosystem and exploring full-stack development.
"@

$Details = [ordered]@{
    "Website"  = @{ Text = "nahid.rf.gd"; Url = "https://nahid.rf.gd/" }
	"Website 2"  = @{ Text = "notnahid.rf.gd"; Url = "https://notnahid.rf.gd/" }
    "Email"    = @{ Text = "nahidul.live@gmail.com"; Url = "mailto:nahidul.live@gmail.com" }
    "Timezone" = @{ Text = "GMT+6:00 Bangladesh Standard Time"; Url = $null }
}

$Skills = [ordered]@{
    "Languages"  = @("JavaScript", "HTML5", "CSS3")
    "Frameworks" = @("React", "Node.js", "Vite")
    "Tools"      = @("Git", "GitHub", "VS Code")
    "Learning"   = @("Next.js", "Tailwind CSS", "Firebase")
}

$Projects = @(
    @{ Name = "Vintage Book App"; Url = "https://github.com/NotNahid/vintage-book-app"; Desc = "A book browsing application with vintage aesthetics" }
)

$Links = @(
    @{ Name = "GitHub";    Url = "https://github.com/NotNahid" },
    @{ Name = "Facebook";  Url = "https://www.facebook.com/notnahid" },
    @{ Name = "Twitter";   Url = "https://x.com/notnahid" },
    @{ Name = "Instagram"; Url = "https://www.instagram.com/notnahid.me/" },
    @{ Name = "Goodreads"; Url = "https://www.goodreads.com/notnahid" },
    @{ Name = "Email";     Url = "mailto:nahidul.live@gmail.com" }
)

# ========================
# THEME
# ========================

$Theme = @{
    Bg         = [System.Drawing.Color]::FromArgb(18, 18, 24)
    Surface    = [System.Drawing.Color]::FromArgb(28, 28, 38)
    SurfaceLt  = [System.Drawing.Color]::FromArgb(36, 36, 48)
    Border     = [System.Drawing.Color]::FromArgb(50, 50, 65)
    Accent     = [System.Drawing.Color]::FromArgb(100, 140, 255)
    AccentDim  = [System.Drawing.Color]::FromArgb(60, 90, 180)
    Text       = [System.Drawing.Color]::FromArgb(230, 230, 240)
    TextDim    = [System.Drawing.Color]::FromArgb(140, 140, 165)
    TextFaint  = [System.Drawing.Color]::FromArgb(80, 80, 100)
    Hover      = [System.Drawing.Color]::FromArgb(40, 40, 55)
    Active     = [System.Drawing.Color]::FromArgb(42, 45, 68)
    Success    = [System.Drawing.Color]::FromArgb(80, 200, 120)
    PillBg     = [System.Drawing.Color]::FromArgb(35, 40, 58)
}

# ========================
# FONTS
# ========================

$Fonts = @{
    Main    = New-Object System.Drawing.Font("Segoe UI", 9.5)
    Bold    = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    Title   = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
    Section = New-Object System.Drawing.Font("Segoe UI", 11.5, [System.Drawing.FontStyle]::Bold)
    Small   = New-Object System.Drawing.Font("Segoe UI", 8.5)
    Nav     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    Tiny    = New-Object System.Drawing.Font("Segoe UI", 7.5)
    Detail  = New-Object System.Drawing.Font("Segoe UI", 9)
}

# ========================
# SAFE URL OPENER
# ========================

function Open-Url([string]$url) {
    if (-not $url) { return }
    try {
        $si = New-Object System.Diagnostics.ProcessStartInfo
        $si.FileName = $url
        $si.UseShellExecute = $true
        [System.Diagnostics.Process]::Start($si) | Out-Null
    } catch { }
}

# ========================
# FORM
# ========================

$form = New-Object System.Windows.Forms.Form
$form.Text = "$Name - Portfolio"
$form.Size = New-Object System.Drawing.Size(950, 620)
$form.MinimumSize = New-Object System.Drawing.Size(650, 450)
$form.StartPosition = "CenterScreen"
$form.BackColor = $Theme.Bg
$form.Icon = $appIcon

$prop = $form.GetType().GetProperty(
    "DoubleBuffered",
    [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic
)
$prop.SetValue($form, $true)

# ========================
# STATUS BAR
# ========================

$statusBar = New-Object System.Windows.Forms.Panel
$statusBar.Dock = "Bottom"
$statusBar.Height = 28
$statusBar.BackColor = $Theme.Surface

$statusBorder = New-Object System.Windows.Forms.Panel
$statusBorder.Dock = "Top"
$statusBorder.Height = 1
$statusBorder.BackColor = $Theme.Border
$statusBar.Controls.Add($statusBorder)

$statusLeft = New-Object System.Windows.Forms.Label
$statusLeft.Text = "  Ready"
$statusLeft.Font = $Fonts.Tiny
$statusLeft.ForeColor = $Theme.TextDim
$statusLeft.BackColor = [System.Drawing.Color]::Transparent
$statusLeft.Dock = "Left"
$statusLeft.AutoSize = $true
$statusLeft.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$statusLeft.Padding = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
$statusBar.Controls.Add($statusLeft)

$statusRight = New-Object System.Windows.Forms.Label
$statusRight.Text = "PowerShell Portfolio  "
$statusRight.Font = $Fonts.Tiny
$statusRight.ForeColor = $Theme.TextFaint
$statusRight.BackColor = [System.Drawing.Color]::Transparent
$statusRight.Dock = "Right"
$statusRight.AutoSize = $true
$statusRight.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$statusRight.Padding = New-Object System.Windows.Forms.Padding(0, 0, 5, 0)
$statusBar.Controls.Add($statusRight)

$form.Controls.Add($statusBar)

# ========================
# SIDEBAR
# ========================

$sidebar = New-Object System.Windows.Forms.Panel
$sidebar.Dock = "Left"
$sidebar.Width = 210
$sidebar.BackColor = $Theme.Surface

$sBorder = New-Object System.Windows.Forms.Panel
$sBorder.Dock = "Left"
$sBorder.Width = 1
$sBorder.BackColor = $Theme.Border

# --- Profile with avatar ---

$profilePanel = New-Object System.Windows.Forms.Panel
$profilePanel.Dock = "Top"
$profilePanel.Height = 140
$profilePanel.BackColor = [System.Drawing.Color]::Transparent
$profilePanel.Padding = New-Object System.Windows.Forms.Padding(15, 12, 15, 0)

# Avatar
$avatarBox = New-Object System.Windows.Forms.PictureBox
$avatarBox.Size = New-Object System.Drawing.Size(56, 56)
$avatarBox.Location = New-Object System.Drawing.Point(77, 10)
$avatarBox.SizeMode = "Zoom"
$avatarBox.BackColor = [System.Drawing.Color]::Transparent

if ($iconBitmap) {
    # Make circular avatar
    $circularBmp = New-Object System.Drawing.Bitmap(56, 56)
    $g = [System.Drawing.Graphics]::FromImage($circularBmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.Clear([System.Drawing.Color]::Transparent)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddEllipse(0, 0, 55, 55)
    $g.SetClip($path)
    $scaledImg = New-Object System.Drawing.Bitmap($iconBitmap, 56, 56)
    $g.DrawImage($scaledImg, 0, 0, 56, 56)
    $g.Dispose()
    $scaledImg.Dispose()
    $avatarBox.Image = $circularBmp
}

$profilePanel.Controls.Add($avatarBox)

$lblName = New-Object System.Windows.Forms.Label
$lblName.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$lblName.ForeColor = $Theme.Text
$lblName.BackColor = [System.Drawing.Color]::Transparent
$lblName.Location = New-Object System.Drawing.Point(0, 70)
$lblName.Size = New-Object System.Drawing.Size(210, 25)
$lblName.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$lblName.Text = $Name
$profilePanel.Controls.Add($lblName)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Font = $Fonts.Tiny
$lblStatus.ForeColor = $Theme.Success
$lblStatus.BackColor = [System.Drawing.Color]::Transparent
$lblStatus.Location = New-Object System.Drawing.Point(0, 95)
$lblStatus.Size = New-Object System.Drawing.Size(210, 15)
$lblStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$lblStatus.Text = "$([char]0x25CF) Available for work"
$profilePanel.Controls.Add($lblStatus)

$lblTag = New-Object System.Windows.Forms.Label
$lblTag.Text = $Tagline
$lblTag.Font = $Fonts.Small
$lblTag.ForeColor = $Theme.TextDim
$lblTag.BackColor = [System.Drawing.Color]::Transparent
$lblTag.Location = New-Object System.Drawing.Point(5, 112)
$lblTag.Size = New-Object System.Drawing.Size(200, 28)
$lblTag.TextAlign = [System.Drawing.ContentAlignment]::TopCenter
$profilePanel.Controls.Add($lblTag)

$sidebar.Controls.Add($profilePanel)

# Separator
$sepLine = New-Object System.Windows.Forms.Panel
$sepLine.Dock = "Top"
$sepLine.Height = 1
$sepLine.BackColor = $Theme.Border

# Nav header
$navHeader = New-Object System.Windows.Forms.Label
$navHeader.Text = "  NAVIGATION"
$navHeader.Font = $Fonts.Tiny
$navHeader.ForeColor = $Theme.TextFaint
$navHeader.BackColor = [System.Drawing.Color]::Transparent
$navHeader.Dock = "Top"
$navHeader.Height = 30
$navHeader.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
$navHeader.Padding = New-Object System.Windows.Forms.Padding(18, 0, 0, 3)

# Footer
$lblFoot = New-Object System.Windows.Forms.Label
$lblFoot.Text = "v1.0 | Made with PowerShell"
$lblFoot.Font = $Fonts.Tiny
$lblFoot.ForeColor = $Theme.TextFaint
$lblFoot.BackColor = [System.Drawing.Color]::Transparent
$lblFoot.Dock = "Bottom"
$lblFoot.Height = 30
$lblFoot.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$sidebar.Controls.Add($lblFoot)

# ========================
# NAV
# ========================

$script:navBtns = @()
$script:activeN = $null

function Activate-Nav($btn) {
    foreach ($b in $script:navBtns) {
        $b.BackColor = [System.Drawing.Color]::Transparent
        $b.NL.ForeColor = $Theme.TextDim
        $b.NI.BackColor = [System.Drawing.Color]::Transparent
    }
    $btn.BackColor = $Theme.Active
    $btn.NL.ForeColor = $Theme.Text
    $btn.NI.BackColor = $Theme.Accent
    $script:activeN = $btn
    Switch-Page $btn.Tag
    $statusLeft.Text = "  $($btn.Tag.Substring(0,1).ToUpper() + $btn.Tag.Substring(1))"
}

function Make-Nav([string]$text, [string]$pageKey, [string]$shortcut) {
    $p = New-Object System.Windows.Forms.Panel
    $p.Dock = "Top"
    $p.Height = 38
    $p.BackColor = [System.Drawing.Color]::Transparent
    $p.Cursor = [System.Windows.Forms.Cursors]::Hand
    $p.Tag = $pageKey

    $ind = New-Object System.Windows.Forms.Panel
    $ind.Dock = "Left"
    $ind.Width = 3
    $ind.BackColor = [System.Drawing.Color]::Transparent
    $p.Controls.Add($ind)

    $l = New-Object System.Windows.Forms.Label
    $l.Text = "  $text"
    $l.Font = $Fonts.Nav
    $l.ForeColor = $Theme.TextDim
    $l.BackColor = [System.Drawing.Color]::Transparent
    $l.Dock = "Fill"
    $l.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $l.Padding = New-Object System.Windows.Forms.Padding(15, 0, 0, 0)
    $l.Cursor = [System.Windows.Forms.Cursors]::Hand
    $p.Controls.Add($l)

    if ($shortcut) {
        $hint = New-Object System.Windows.Forms.Label
        $hint.Text = $shortcut
        $hint.Font = $Fonts.Tiny
        $hint.ForeColor = $Theme.TextFaint
        $hint.BackColor = [System.Drawing.Color]::Transparent
        $hint.Dock = "Right"
        $hint.Width = 35
        $hint.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $hint.Cursor = [System.Windows.Forms.Cursors]::Hand
        $p.Controls.Add($hint)
        $hint.Add_Click({
            param($s, $e)
            $btn = $s.Parent
            if ($btn -eq $script:activeN) { return }
            Activate-Nav $btn
        }.GetNewClosure())
    }

    $p | Add-Member -NotePropertyName "NL" -NotePropertyValue $l
    $p | Add-Member -NotePropertyName "NI" -NotePropertyValue $ind

    $click = {
        param($s, $e)
        $btn = if ($s -is [System.Windows.Forms.Label]) { $s.Parent } else { $s }
        if ($btn -eq $script:activeN) { return }
        Activate-Nav $btn
    }.GetNewClosure()

    $eIn  = { if ($this -ne $script:activeN) { $this.BackColor = $Theme.Hover } }.GetNewClosure()
    $eOut = { if ($this -ne $script:activeN) { $this.BackColor = [System.Drawing.Color]::Transparent } }.GetNewClosure()
    $lIn  = { if ($this.Parent -ne $script:activeN) { $this.Parent.BackColor = $Theme.Hover } }.GetNewClosure()
    $lOut = { if ($this.Parent -ne $script:activeN) { $this.Parent.BackColor = [System.Drawing.Color]::Transparent } }.GetNewClosure()

    $p.Add_Click($click); $l.Add_Click($click)
    $p.Add_MouseEnter($eIn); $p.Add_MouseLeave($eOut)
    $l.Add_MouseEnter($lIn); $l.Add_MouseLeave($lOut)

    $script:navBtns += $p
    return $p
}

$nContact  = Make-Nav "Contact"  "contact"  "F3"
$nProjects = Make-Nav "Projects" "projects" "F2"
$nAbout    = Make-Nav "About"    "about"    "F1"

$sidebar.Controls.Add($nContact)
$sidebar.Controls.Add($nProjects)
$sidebar.Controls.Add($nAbout)
$sidebar.Controls.Add($navHeader)
$sidebar.Controls.Add($sepLine)

$form.Controls.Add($sBorder)
$form.Controls.Add($sidebar)

# Keyboard shortcuts
$form.KeyPreview = $true
$form.Add_KeyDown({
    param($s, $e)
    switch ($e.KeyCode) {
        "F1" { Activate-Nav $nAbout;    $e.Handled = $true }
        "F2" { Activate-Nav $nProjects; $e.Handled = $true }
        "F3" { Activate-Nav $nContact;  $e.Handled = $true }
    }
})

# ========================
# CONTENT HOLDER
# ========================

$content = New-Object System.Windows.Forms.Panel
$content.Dock = "Fill"
$content.BackColor = $Theme.Bg
$form.Controls.Add($content)
$form.Controls.SetChildIndex($content, 0)

# ========================
# CARD BUILDER
# ========================

function Make-Card([string]$title, [string]$desc, [string]$url) {
    $c = New-Object System.Windows.Forms.Panel
    $c.Dock = "Top"
    $c.Height = if ($desc) { 62 } else { 44 }
    $c.BackColor = $Theme.Surface
    $c.Cursor = [System.Windows.Forms.Cursors]::Hand
    $c.Padding = New-Object System.Windows.Forms.Padding(14, 0, 10, 0)

    $sp = New-Object System.Windows.Forms.Panel
    $sp.Dock = "Top"
    $sp.Height = 5
    $sp.BackColor = $Theme.Bg

    $leftBar = New-Object System.Windows.Forms.Panel
    $leftBar.Dock = "Left"
    $leftBar.Width = 3
    $leftBar.BackColor = $Theme.AccentDim
    $c.Controls.Add($leftBar)

    if ($url) {
        $arrow = New-Object System.Windows.Forms.Label
        $arrow.Text = [char]0x2192
        $arrow.Font = $Fonts.Bold
        $arrow.ForeColor = $Theme.TextFaint
        $arrow.BackColor = [System.Drawing.Color]::Transparent
        $arrow.Dock = "Right"
        $arrow.Width = 30
        $arrow.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $arrow.Cursor = [System.Windows.Forms.Cursors]::Hand
        $c.Controls.Add($arrow)
        $arrow.Add_Click({ Open-Url $url }.GetNewClosure())
    }

    $t = New-Object System.Windows.Forms.Label
    $t.Text = "  $title"
    $t.Font = $Fonts.Bold
    $t.ForeColor = $Theme.Accent
    $t.BackColor = [System.Drawing.Color]::Transparent
    $t.Cursor = [System.Windows.Forms.Cursors]::Hand

    $safeUrl = $url

    if ($desc) {
        $t.Dock = "Top"
        $t.Height = 28
        $t.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft

        $d = New-Object System.Windows.Forms.Label
        $d.Text = "  $desc"
        $d.Font = $Fonts.Small
        $d.ForeColor = $Theme.TextDim
        $d.BackColor = [System.Drawing.Color]::Transparent
        $d.Dock = "Fill"
        $d.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
        $d.Cursor = [System.Windows.Forms.Cursors]::Hand

        $c.Controls.Add($d)
        $c.Controls.Add($t)

        $d.Add_Click({ Open-Url $safeUrl }.GetNewClosure())
        $d.Add_MouseEnter({ $this.Parent.BackColor = $Theme.Hover }.GetNewClosure())
        $d.Add_MouseLeave({ $this.Parent.BackColor = $Theme.Surface }.GetNewClosure())
    } else {
        $t.Dock = "Fill"
        $t.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        $c.Controls.Add($t)
    }

    if ($safeUrl) {
        $c.Add_Click({ Open-Url $safeUrl }.GetNewClosure())
        $t.Add_Click({ Open-Url $safeUrl }.GetNewClosure())
    }

    $c.Add_MouseEnter({
        $this.BackColor = $Theme.Hover
        $bar = $this.Controls[0]
        if ($bar.Width -eq 3) { $bar.BackColor = $Theme.Accent }
    }.GetNewClosure())
    $c.Add_MouseLeave({
        $this.BackColor = $Theme.Surface
        $bar = $this.Controls[0]
        if ($bar.Width -eq 3) { $bar.BackColor = $Theme.AccentDim }
    }.GetNewClosure())
    $t.Add_MouseEnter({ $this.Parent.BackColor = $Theme.Hover }.GetNewClosure())
    $t.Add_MouseLeave({ $this.Parent.BackColor = $Theme.Surface }.GetNewClosure())

    return @($sp, $c)
}

# ========================
# PAGE: ABOUT
# ========================

$pgAbout = New-Object System.Windows.Forms.Panel
$pgAbout.Dock = "Fill"
$pgAbout.AutoScroll = $true
$pgAbout.BackColor = $Theme.Bg
$pgAbout.Visible = $false

$aboutFlow = New-Object System.Windows.Forms.FlowLayoutPanel
$aboutFlow.FlowDirection = "TopDown"
$aboutFlow.WrapContents = $false
$aboutFlow.AutoSize = $true
$aboutFlow.AutoSizeMode = "GrowOnly"
$aboutFlow.Dock = "Top"
$aboutFlow.BackColor = $Theme.Bg
$aboutFlow.Padding = New-Object System.Windows.Forms.Padding(30, 20, 30, 20)

$aTitle = New-Object System.Windows.Forms.Label
$aTitle.Text = "About Me"
$aTitle.Font = $Fonts.Title
$aTitle.ForeColor = $Theme.Text
$aTitle.AutoSize = $true
$aTitle.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 5)
$aboutFlow.Controls.Add($aTitle)

$aSubtitle = New-Object System.Windows.Forms.Label
$aSubtitle.Text = "Front-end developer from Bangladesh"
$aSubtitle.Font = $Fonts.Detail
$aSubtitle.ForeColor = $Theme.AccentDim
$aSubtitle.AutoSize = $true
$aSubtitle.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
$aboutFlow.Controls.Add($aSubtitle)

$aBio = New-Object System.Windows.Forms.Label
$aBio.Text = $Bio
$aBio.Font = $Fonts.Main
$aBio.ForeColor = $Theme.TextDim
$aBio.AutoSize = $true
$aBio.MaximumSize = New-Object System.Drawing.Size(700, 0)
$aBio.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
$aboutFlow.Controls.Add($aBio)

foreach ($key in $Details.Keys) {
    $d = $Details[$key]
    $row = New-Object System.Windows.Forms.Label
    $row.Font = $Fonts.Detail
    $row.AutoSize = $true
    $row.Margin = New-Object System.Windows.Forms.Padding(2, 2, 0, 2)

    if ($d.Url) {
        $row.Text = "$($key): $($d.Text)"
        $row.ForeColor = $Theme.Accent
        $row.Cursor = [System.Windows.Forms.Cursors]::Hand
        $detailUrl = $d.Url
        $row.Add_Click({ Open-Url $detailUrl }.GetNewClosure())
        $row.Add_MouseEnter({ $this.ForeColor = $Theme.Text }.GetNewClosure())
        $row.Add_MouseLeave({ $this.ForeColor = $Theme.Accent }.GetNewClosure())
    } else {
        $row.Text = "$($key): $($d.Text)"
        $row.ForeColor = $Theme.TextDim
    }
    $aboutFlow.Controls.Add($row)
}

$aSep = New-Object System.Windows.Forms.Panel
$aSep.Height = 1
$aSep.Width = 400
$aSep.BackColor = $Theme.Border
$aSep.Margin = New-Object System.Windows.Forms.Padding(0, 15, 0, 10)
$aboutFlow.Controls.Add($aSep)

$aSkTitle = New-Object System.Windows.Forms.Label
$aSkTitle.Text = "Skills & Technologies"
$aSkTitle.Font = $Fonts.Section
$aSkTitle.ForeColor = $Theme.Text
$aSkTitle.AutoSize = $true
$aSkTitle.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 8)
$aboutFlow.Controls.Add($aSkTitle)

foreach ($key in $Skills.Keys) {
    $cat = New-Object System.Windows.Forms.Label
    $cat.Text = $key
    $cat.Font = $Fonts.Bold
    $cat.ForeColor = $Theme.Text
    $cat.AutoSize = $true
    $cat.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 4)
    $aboutFlow.Controls.Add($cat)

    $pillFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    $pillFlow.FlowDirection = "LeftToRight"
    $pillFlow.WrapContents = $true
    $pillFlow.AutoSize = $true
    $pillFlow.AutoSizeMode = "GrowOnly"
    $pillFlow.BackColor = [System.Drawing.Color]::Transparent
    $pillFlow.Margin = New-Object System.Windows.Forms.Padding(2, 0, 0, 6)

    foreach ($skill in $Skills[$key]) {
        $pill = New-Object System.Windows.Forms.Label
        $pill.Text = " $skill "
        $pill.Font = $Fonts.Detail
        $pill.ForeColor = $Theme.Accent
        $pill.BackColor = $Theme.PillBg
        $pill.AutoSize = $true
        $pill.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)
        $pill.Margin = New-Object System.Windows.Forms.Padding(0, 2, 6, 2)
        $pillFlow.Controls.Add($pill)
    }
    $aboutFlow.Controls.Add($pillFlow)
}

$pgAbout.Controls.Add($aboutFlow)

$pgAbout.Add_Resize({
    $w = $this.ClientSize.Width - 80
    if ($w -lt 200) { $w = 200 }
    $aBio.MaximumSize = New-Object System.Drawing.Size($w, 0)
    $aSep.Width = [Math]::Min($w, 500)
}.GetNewClosure())

# ========================
# PAGE: PROJECTS
# ========================

$pgProjects = New-Object System.Windows.Forms.Panel
$pgProjects.Dock = "Fill"
$pgProjects.AutoScroll = $true
$pgProjects.BackColor = $Theme.Bg
$pgProjects.Visible = $false

$projInner = New-Object System.Windows.Forms.Panel
$projInner.Dock = "Top"
$projInner.AutoSize = $true
$projInner.BackColor = $Theme.Bg
$projInner.Padding = New-Object System.Windows.Forms.Padding(30, 20, 30, 20)

$pTitle = New-Object System.Windows.Forms.Label
$pTitle.Text = "Projects"
$pTitle.Font = $Fonts.Title
$pTitle.ForeColor = $Theme.Text
$pTitle.Dock = "Top"
$pTitle.Height = 35

$pSub = New-Object System.Windows.Forms.Label
$pSub.Text = "Things I've built and am working on."
$pSub.Font = $Fonts.Main
$pSub.ForeColor = $Theme.TextDim
$pSub.Dock = "Top"
$pSub.Height = 30

$pCount = New-Object System.Windows.Forms.Label
$pCount.Text = "$($Projects.Count) project$(if($Projects.Count -ne 1){'s'})"
$pCount.Font = $Fonts.Tiny
$pCount.ForeColor = $Theme.TextFaint
$pCount.Dock = "Top"
$pCount.Height = 22

$pMore = New-Object System.Windows.Forms.Label
$pMore.Text = "More projects on GitHub $([char]0x2192)"
$pMore.Font = $Fonts.Small
$pMore.ForeColor = $Theme.Accent
$pMore.Dock = "Top"
$pMore.Height = 30
$pMore.Cursor = [System.Windows.Forms.Cursors]::Hand
$pMore.Padding = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
$pMore.Add_Click({ Open-Url "https://github.com/NotNahid?tab=repositories" })
$pMore.Add_MouseEnter({ $this.ForeColor = $Theme.Text }.GetNewClosure())
$pMore.Add_MouseLeave({ $this.ForeColor = $Theme.Accent }.GetNewClosure())

$projInner.Controls.Add($pMore)

for ($i = $Projects.Count - 1; $i -ge 0; $i--) {
    $pr = $Projects[$i]
    $cards = Make-Card $pr.Name $pr.Desc $pr.Url
    foreach ($c in $cards) { $projInner.Controls.Add($c) }
}

$projInner.Controls.Add($pCount)
$projInner.Controls.Add($pSub)
$projInner.Controls.Add($pTitle)

$pgProjects.Controls.Add($projInner)

# ========================
# PAGE: CONTACT
# ========================

$pgContact = New-Object System.Windows.Forms.Panel
$pgContact.Dock = "Fill"
$pgContact.AutoScroll = $true
$pgContact.BackColor = $Theme.Bg
$pgContact.Visible = $false

$conInner = New-Object System.Windows.Forms.Panel
$conInner.Dock = "Top"
$conInner.AutoSize = $true
$conInner.BackColor = $Theme.Bg
$conInner.Padding = New-Object System.Windows.Forms.Padding(30, 20, 30, 20)

$cTitle = New-Object System.Windows.Forms.Label
$cTitle.Text = "Contact"
$cTitle.Font = $Fonts.Title
$cTitle.ForeColor = $Theme.Text
$cTitle.Dock = "Top"
$cTitle.Height = 35

$cSub = New-Object System.Windows.Forms.Label
$cSub.Text = "Feel free to reach out through any of these channels."
$cSub.Font = $Fonts.Main
$cSub.ForeColor = $Theme.TextDim
$cSub.Dock = "Top"
$cSub.Height = 30

$cCount = New-Object System.Windows.Forms.Label
$cCount.Text = "$($Links.Count) channels"
$cCount.Font = $Fonts.Tiny
$cCount.ForeColor = $Theme.TextFaint
$cCount.Dock = "Top"
$cCount.Height = 22

for ($i = $Links.Count - 1; $i -ge 0; $i--) {
    $lk = $Links[$i]
    $cards = Make-Card $lk.Name $null $lk.Url
    foreach ($c in $cards) { $conInner.Controls.Add($c) }
}

$conInner.Controls.Add($cCount)
$conInner.Controls.Add($cSub)
$conInner.Controls.Add($cTitle)

$pgContact.Controls.Add($conInner)

# ========================
# ADD PAGES & SWITCHER
# ========================

$content.Controls.Add($pgAbout)
$content.Controls.Add($pgProjects)
$content.Controls.Add($pgContact)

function Switch-Page([string]$key) {
    $pgAbout.Visible    = ($key -eq "about")
    $pgProjects.Visible = ($key -eq "projects")
    $pgContact.Visible  = ($key -eq "contact")
}

# ========================
# LAUNCH
# ========================

Activate-Nav $nAbout

$form.Add_FormClosing({
    foreach ($f in $Fonts.Values) { $f.Dispose() }
    if ($iconBitmap) { $iconBitmap.Dispose() }
    if ($ms) { $ms.Dispose() }
    if ($img) { $img.Dispose() }
    if ($circularBmp) { $circularBmp.Dispose() }
})

[void]$form.ShowDialog()
