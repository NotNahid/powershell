Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ========================
# CONFIG
# ========================

$Name = "Nahid"
$Tagline = "Web Developer | Front-End | Always Learning"
$Bio = @"
A passionate self-taught developer from Bangladesh, interested in creating beautiful and engaging user experiences. Currently diving deeper into the React ecosystem and exploring full-stack development.

Website: nahid.rf.gd
Email: nahidul.live@gmail.com
Timezone: GMT+6:00 Bangladesh Standard Time
"@

$Skills = [ordered]@{
    "Languages"  = "JavaScript  |  HTML5  |  CSS3"
    "Frameworks" = "React  |  Node.js  |  Vite"
    "Tools"      = "Git  |  GitHub  |  VS Code"
    "Learning"   = "Next.js  |  Tailwind CSS  |  Firebase"
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
# COLORS & FONTS
# ========================

$Bg     = [System.Drawing.Color]::FromArgb(22, 22, 30)
$Panel  = [System.Drawing.Color]::FromArgb(32, 32, 42)
$Border = [System.Drawing.Color]::FromArgb(55, 55, 70)
$Accent = [System.Drawing.Color]::FromArgb(100, 140, 255)
$Txt    = [System.Drawing.Color]::FromArgb(230, 230, 240)
$Dim    = [System.Drawing.Color]::FromArgb(150, 150, 170)
$Hover  = [System.Drawing.Color]::FromArgb(42, 42, 55)
$Active = [System.Drawing.Color]::FromArgb(45, 50, 75)
$PillBg = [System.Drawing.Color]::FromArgb(38, 42, 58)

$FMain  = New-Object System.Drawing.Font("Segoe UI", 10)
$FBold  = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$FTitle = New-Object System.Drawing.Font("Segoe UI", 20, [System.Drawing.FontStyle]::Bold)
$FSec   = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$FSmall = New-Object System.Drawing.Font("Segoe UI", 8.5)
$FNav   = New-Object System.Drawing.Font("Segoe UI", 10.5, [System.Drawing.FontStyle]::Bold)

# ========================
# SAFE URL OPENER (won't freeze)
# ========================

function Open-Url([string]$url) {
    try {
        $si = New-Object System.Diagnostics.ProcessStartInfo
        $si.FileName = $url
        $si.UseShellExecute = $true
        [System.Diagnostics.Process]::Start($si) | Out-Null
    } catch {
        # silently fail if browser not found etc
    }
}

# ========================
# FORM
# ========================

$form = New-Object System.Windows.Forms.Form
$form.Text = "$Name - Portfolio"
$form.Size = New-Object System.Drawing.Size(900, 580)
$form.MinimumSize = New-Object System.Drawing.Size(600, 400)
$form.StartPosition = "CenterScreen"
$form.BackColor = $Bg

$prop = $form.GetType().GetProperty(
    "DoubleBuffered",
    [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic
)
$prop.SetValue($form, $true)

# ========================
# SIDEBAR
# ========================

$sidebar = New-Object System.Windows.Forms.Panel
$sidebar.Dock = "Left"
$sidebar.Width = 200
$sidebar.BackColor = $Panel

$sBorder = New-Object System.Windows.Forms.Panel
$sBorder.Dock = "Left"
$sBorder.Width = 1
$sBorder.BackColor = $Border

$lblName = New-Object System.Windows.Forms.Label
$lblName.Text = $Name
$lblName.Font = New-Object System.Drawing.Font("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
$lblName.ForeColor = $Txt
$lblName.BackColor = [System.Drawing.Color]::Transparent
$lblName.Dock = "Top"
$lblName.Height = 45
$lblName.TextAlign = [System.Drawing.ContentAlignment]::BottomCenter

$lblTag = New-Object System.Windows.Forms.Label
$lblTag.Text = $Tagline
$lblTag.Font = $FSmall
$lblTag.ForeColor = $Dim
$lblTag.BackColor = [System.Drawing.Color]::Transparent
$lblTag.Dock = "Top"
$lblTag.Height = 32
$lblTag.TextAlign = [System.Drawing.ContentAlignment]::TopCenter
$lblTag.Padding = New-Object System.Windows.Forms.Padding(5, 3, 5, 0)

$sepLine = New-Object System.Windows.Forms.Panel
$sepLine.Dock = "Top"
$sepLine.Height = 1
$sepLine.BackColor = $Border

$spacer1 = New-Object System.Windows.Forms.Panel
$spacer1.Dock = "Top"
$spacer1.Height = 12
$spacer1.BackColor = [System.Drawing.Color]::Transparent

$lblFoot = New-Object System.Windows.Forms.Label
$lblFoot.Text = "Powered by PowerShell"
$lblFoot.Font = $FSmall
$lblFoot.ForeColor = [System.Drawing.Color]::FromArgb(70, 70, 90)
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
        $b.NL.ForeColor = $Dim
        $b.NI.BackColor = [System.Drawing.Color]::Transparent
    }
    $btn.BackColor = $Active
    $btn.NL.ForeColor = $Txt
    $btn.NI.BackColor = $Accent
    $script:activeN = $btn
    Switch-Page $btn.Tag
}

function Make-Nav([string]$text, [string]$pageKey) {
    $p = New-Object System.Windows.Forms.Panel
    $p.Dock = "Top"
    $p.Height = 40
    $p.BackColor = [System.Drawing.Color]::Transparent
    $p.Cursor = [System.Windows.Forms.Cursors]::Hand
    $p.Tag = $pageKey

    $ind = New-Object System.Windows.Forms.Panel
    $ind.Dock = "Left"
    $ind.Width = 3
    $ind.BackColor = [System.Drawing.Color]::Transparent
    $p.Controls.Add($ind)

    $l = New-Object System.Windows.Forms.Label
    $l.Text = $text
    $l.Font = $FNav
    $l.ForeColor = $Dim
    $l.BackColor = [System.Drawing.Color]::Transparent
    $l.Dock = "Fill"
    $l.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $l.Padding = New-Object System.Windows.Forms.Padding(18, 0, 0, 0)
    $l.Cursor = [System.Windows.Forms.Cursors]::Hand
    $p.Controls.Add($l)

    $p | Add-Member -NotePropertyName "NL" -NotePropertyValue $l
    $p | Add-Member -NotePropertyName "NI" -NotePropertyValue $ind

    $click = {
        param($s, $e)
        $btn = if ($s -is [System.Windows.Forms.Label]) { $s.Parent } else { $s }
        if ($btn -eq $script:activeN) { return }
        Activate-Nav $btn
    }.GetNewClosure()

    $eIn  = { if ($this -ne $script:activeN) { $this.BackColor = $Hover } }.GetNewClosure()
    $eOut = { if ($this -ne $script:activeN) { $this.BackColor = [System.Drawing.Color]::Transparent } }.GetNewClosure()
    $lIn  = { if ($this.Parent -ne $script:activeN) { $this.Parent.BackColor = $Hover } }.GetNewClosure()
    $lOut = { if ($this.Parent -ne $script:activeN) { $this.Parent.BackColor = [System.Drawing.Color]::Transparent } }.GetNewClosure()

    $p.Add_Click($click); $l.Add_Click($click)
    $p.Add_MouseEnter($eIn); $p.Add_MouseLeave($eOut)
    $l.Add_MouseEnter($lIn); $l.Add_MouseLeave($lOut)

    $script:navBtns += $p
    return $p
}

$nContact  = Make-Nav "Contact"  "contact"
$nProjects = Make-Nav "Projects" "projects"
$nAbout    = Make-Nav "About"    "about"

$sidebar.Controls.Add($nContact)
$sidebar.Controls.Add($nProjects)
$sidebar.Controls.Add($nAbout)
$sidebar.Controls.Add($spacer1)
$sidebar.Controls.Add($sepLine)
$sidebar.Controls.Add($lblTag)
$sidebar.Controls.Add($lblName)

$form.Controls.Add($sBorder)
$form.Controls.Add($sidebar)

# ========================
# CONTENT HOLDER
# ========================

$content = New-Object System.Windows.Forms.Panel
$content.Dock = "Fill"
$content.BackColor = $Bg
$form.Controls.Add($content)
$form.Controls.SetChildIndex($content, 0)

# ========================
# HELPER: clickable card
# ========================

function Make-Card([string]$title, [string]$desc, [string]$url) {
    $c = New-Object System.Windows.Forms.Panel
    $c.Dock = "Top"
    $c.Height = if ($desc) { 58 } else { 42 }
    $c.BackColor = $Panel
    $c.Cursor = [System.Windows.Forms.Cursors]::Hand
    $c.Padding = New-Object System.Windows.Forms.Padding(14, 0, 10, 0)

    $sp = New-Object System.Windows.Forms.Panel
    $sp.Dock = "Top"
    $sp.Height = 6
    $sp.BackColor = $Bg

    $t = New-Object System.Windows.Forms.Label
    $t.Text = $title
    $t.Font = $FBold
    $t.ForeColor = $Accent
    $t.BackColor = [System.Drawing.Color]::Transparent
    $t.Cursor = [System.Windows.Forms.Cursors]::Hand

    $safeUrl = $url

    if ($desc) {
        $t.Dock = "Top"
        $t.Height = 28
        $t.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft

        $d = New-Object System.Windows.Forms.Label
        $d.Text = $desc
        $d.Font = $FSmall
        $d.ForeColor = $Dim
        $d.BackColor = [System.Drawing.Color]::Transparent
        $d.Dock = "Fill"
        $d.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
        $d.Cursor = [System.Windows.Forms.Cursors]::Hand

        $c.Controls.Add($d)
        $c.Controls.Add($t)

        $d.Add_Click({ Open-Url $safeUrl }.GetNewClosure())
        $d.Add_MouseEnter({ $this.Parent.BackColor = $Hover }.GetNewClosure())
        $d.Add_MouseLeave({ $this.Parent.BackColor = $Panel }.GetNewClosure())
    } else {
        $t.Dock = "Fill"
        $t.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        $c.Controls.Add($t)
    }

    $c.Add_Click({ Open-Url $safeUrl }.GetNewClosure())
    $t.Add_Click({ Open-Url $safeUrl }.GetNewClosure())

    $c.Add_MouseEnter({ $this.BackColor = $Hover }.GetNewClosure())
    $c.Add_MouseLeave({ $this.BackColor = $Panel }.GetNewClosure())
    $t.Add_MouseEnter({ $this.Parent.BackColor = $Hover }.GetNewClosure())
    $t.Add_MouseLeave({ $this.Parent.BackColor = $Panel }.GetNewClosure())

    return @($sp, $c)
}

# ========================
# PAGE: ABOUT
# ========================

$pgAbout = New-Object System.Windows.Forms.Panel
$pgAbout.Dock = "Fill"
$pgAbout.AutoScroll = $true
$pgAbout.BackColor = $Bg
$pgAbout.Visible = $false

$aboutInner = New-Object System.Windows.Forms.FlowLayoutPanel
$aboutInner.FlowDirection = "TopDown"
$aboutInner.WrapContents = $false
$aboutInner.AutoSize = $true
$aboutInner.AutoSizeMode = "GrowOnly"
$aboutInner.Dock = "Top"
$aboutInner.BackColor = $Bg
$aboutInner.Padding = New-Object System.Windows.Forms.Padding(30, 20, 30, 20)

$aTitle = New-Object System.Windows.Forms.Label
$aTitle.Text = "About Me"
$aTitle.Font = $FTitle
$aTitle.ForeColor = $Txt
$aTitle.AutoSize = $true
$aTitle.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
$aboutInner.Controls.Add($aTitle)

$aBio = New-Object System.Windows.Forms.Label
$aBio.Text = $Bio
$aBio.Font = $FMain
$aBio.ForeColor = $Dim
$aBio.AutoSize = $true
$aBio.MaximumSize = New-Object System.Drawing.Size(800, 0)
$aBio.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 15)
$aboutInner.Controls.Add($aBio)

$aSep = New-Object System.Windows.Forms.Panel
$aSep.Height = 1
$aSep.Width = 400
$aSep.BackColor = $Border
$aSep.Margin = New-Object System.Windows.Forms.Padding(0, 5, 0, 15)
$aboutInner.Controls.Add($aSep)

$aSkillsTitle = New-Object System.Windows.Forms.Label
$aSkillsTitle.Text = "Skills & Technologies"
$aSkillsTitle.Font = $FTitle
$aSkillsTitle.ForeColor = $Txt
$aSkillsTitle.AutoSize = $true
$aSkillsTitle.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
$aboutInner.Controls.Add($aSkillsTitle)

foreach ($key in $Skills.Keys) {
    $cat = New-Object System.Windows.Forms.Label
    $cat.Text = $key
    $cat.Font = $FSec
    $cat.ForeColor = $Txt
    $cat.AutoSize = $true
    $cat.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 3)
    $aboutInner.Controls.Add($cat)

    $val = New-Object System.Windows.Forms.Label
    $val.Text = $Skills[$key]
    $val.Font = $FMain
    $val.ForeColor = $Accent
    $val.BackColor = $PillBg
    $val.AutoSize = $true
    $val.Padding = New-Object System.Windows.Forms.Padding(10, 5, 10, 5)
    $val.Margin = New-Object System.Windows.Forms.Padding(5, 0, 0, 8)
    $aboutInner.Controls.Add($val)
}

$pgAbout.Controls.Add($aboutInner)

$pgAbout.Add_Resize({
    $w = $this.ClientSize.Width - 80
    if ($w -lt 200) { $w = 200 }
    $aBio.MaximumSize = New-Object System.Drawing.Size($w, 0)
    $aSep.Width = [Math]::Min($w, 600)
}.GetNewClosure())

# ========================
# PAGE: PROJECTS
# ========================

$pgProjects = New-Object System.Windows.Forms.Panel
$pgProjects.Dock = "Fill"
$pgProjects.AutoScroll = $true
$pgProjects.BackColor = $Bg
$pgProjects.Visible = $false

$projInner = New-Object System.Windows.Forms.Panel
$projInner.Dock = "Top"
$projInner.AutoSize = $true
$projInner.BackColor = $Bg
$projInner.Padding = New-Object System.Windows.Forms.Padding(30, 20, 30, 20)

$pTitle = New-Object System.Windows.Forms.Label
$pTitle.Text = "Projects"
$pTitle.Font = $FTitle
$pTitle.ForeColor = $Txt
$pTitle.Dock = "Top"
$pTitle.Height = 40
$pTitle.BackColor = [System.Drawing.Color]::Transparent

$pSub = New-Object System.Windows.Forms.Label
$pSub.Text = "Things I've built and am working on."
$pSub.Font = $FMain
$pSub.ForeColor = $Dim
$pSub.Dock = "Top"
$pSub.Height = 35
$pSub.BackColor = [System.Drawing.Color]::Transparent

$pMore = New-Object System.Windows.Forms.Label
$pMore.Text = "More projects coming soon - check my GitHub!"
$pMore.Font = $FSmall
$pMore.ForeColor = $Dim
$pMore.Dock = "Top"
$pMore.Height = 30
$pMore.BackColor = [System.Drawing.Color]::Transparent

$projInner.Controls.Add($pMore)

for ($i = $Projects.Count - 1; $i -ge 0; $i--) {
    $pr = $Projects[$i]
    $cards = Make-Card $pr.Name $pr.Desc $pr.Url
    foreach ($c in $cards) { $projInner.Controls.Add($c) }
}

$projInner.Controls.Add($pSub)
$projInner.Controls.Add($pTitle)

$pgProjects.Controls.Add($projInner)

# ========================
# PAGE: CONTACT
# ========================

$pgContact = New-Object System.Windows.Forms.Panel
$pgContact.Dock = "Fill"
$pgContact.AutoScroll = $true
$pgContact.BackColor = $Bg
$pgContact.Visible = $false

$conInner = New-Object System.Windows.Forms.Panel
$conInner.Dock = "Top"
$conInner.AutoSize = $true
$conInner.BackColor = $Bg
$conInner.Padding = New-Object System.Windows.Forms.Padding(30, 20, 30, 20)

$cTitle = New-Object System.Windows.Forms.Label
$cTitle.Text = "Contact"
$cTitle.Font = $FTitle
$cTitle.ForeColor = $Txt
$cTitle.Dock = "Top"
$cTitle.Height = 40
$cTitle.BackColor = [System.Drawing.Color]::Transparent

$cSub = New-Object System.Windows.Forms.Label
$cSub.Text = "Feel free to reach out through any of these channels."
$cSub.Font = $FMain
$cSub.ForeColor = $Dim
$cSub.Dock = "Top"
$cSub.Height = 35
$cSub.BackColor = [System.Drawing.Color]::Transparent

for ($i = $Links.Count - 1; $i -ge 0; $i--) {
    $lk = $Links[$i]
    $cards = Make-Card $lk.Name $null $lk.Url
    foreach ($c in $cards) { $conInner.Controls.Add($c) }
}

$conInner.Controls.Add($cSub)
$conInner.Controls.Add($cTitle)

$pgContact.Controls.Add($conInner)

# ========================
# ADD PAGES
# ========================

$content.Controls.Add($pgAbout)
$content.Controls.Add($pgProjects)
$content.Controls.Add($pgContact)

# ========================
# PAGE SWITCHER
# ========================

function Switch-Page([string]$key) {
    $pgAbout.Visible    = ($key -eq "about")
    $pgProjects.Visible = ($key -eq "projects")
    $pgContact.Visible  = ($key -eq "contact")
}

# ========================
# LAUNCH — show About immediately, no PerformClick
# ========================

# Directly activate About nav and show About page
Activate-Nav $nAbout

$form.Add_FormClosing({
    $FMain.Dispose(); $FBold.Dispose(); $FTitle.Dispose()
    $FSec.Dispose(); $FSmall.Dispose(); $FNav.Dispose()
})

[void]$form.ShowDialog()
