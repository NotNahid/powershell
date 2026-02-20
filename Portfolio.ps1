Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ========================
# ADVANCED GRAPHICS SETUP
# ========================

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Drawing2D;

public class GlassEffect {
    [DllImport("dwmapi.dll")]
    public static extern int DwmExtendFrameIntoClientArea(IntPtr hWnd, ref MARGINS pMarInset);
    
    [DllImport("dwmapi.dll")]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
    
    [DllImport("user32.dll")]
    public static extern bool SetLayeredWindowAttributes(IntPtr hwnd, uint crKey, byte bAlpha, uint dwFlags);

    public const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
    public const int DWMWA_MICA_EFFECT = 1029;
    
    public struct MARGINS {
        public int leftWidth;
        public int rightWidth;
        public int topHeight;
        public int bottomHeight;
    }
}

public class GraphicsExtensions {
    public static void DrawRoundedRectangle(Graphics g, Pen pen, Rectangle bounds, int radius) {
        int diameter = radius * 2;
        Rectangle arc = new Rectangle(bounds.Location, new Size(diameter, diameter));
        GraphicsPath path = new GraphicsPath();
        
        path.AddArc(arc, 180, 90);
        arc.X = bounds.Right - diameter;
        path.AddArc(arc, 270, 90);
        arc.Y = bounds.Bottom - diameter;
        path.AddArc(arc, 0, 90);
        arc.X = bounds.Left;
        path.AddArc(arc, 90, 90);
        path.CloseFigure();
        
        g.DrawPath(pen, path);
    }
    
    public static void FillRoundedRectangle(Graphics g, Brush brush, Rectangle bounds, int radius) {
        int diameter = radius * 2;
        Rectangle arc = new Rectangle(bounds.Location, new Size(diameter, diameter));
        GraphicsPath path = new GraphicsPath();
        
        path.AddArc(arc, 180, 90);
        arc.X = bounds.Right - diameter;
        path.AddArc(arc, 270, 90);
        arc.Y = bounds.Bottom - diameter;
        path.AddArc(arc, 0, 90);
        arc.X = bounds.Left;
        path.AddArc(arc, 90, 90);
        path.CloseFigure();
        
        g.FillPath(brush, path);
    }
}
"@

# ========================
# CONFIG
# ========================

$Name = "Nahidul Islam"
$Tagline = "PowerShell Architect • AI Integrator • Digital Craftsman"
$Bio = @"
Crafting the future of automation through elegant code and intelligent systems. I blend traditional scripting mastery with cutting-edge AI to build solutions that feel like magic but work like precision engineering.

Passionate about pushing PowerShell beyond its limits—creating experiences that are as beautiful as they are functional.
"@

$Skills = @(
    "PowerShell & .NET",
    "AI/ML Integration",
    "Glassmorphic UI",
    "Workflow Automation",
    "System Architecture",
    "Creative Coding"
)

$Projects = @(
    @{ 
        Name = "Quantum UI Framework"; 
        Url = "https://github.com/yourrepo1";
        Desc = "Next-gen PowerShell GUI toolkit with glassmorphism and fluid animations"
        Emoji = "✨"
    },
    @{ 
        Name = "Neural Workflow Engine"; 
        Url = "https://github.com/yourrepo2";
        Desc = "AI-powered automation that learns and adapts to your patterns"
        Emoji = "🧠"
    },
    @{ 
        Name = "Liquid Terminal"; 
        Url = "https://github.com/yourrepo3";
        Desc = "Interactive CLI with real-time gradients and smooth transitions"
        Emoji = "💎"
    },
    @{
        Name = "Hyper Orchestrator";
        Url = "https://github.com/yourrepo4";
        Desc = "Multi-dimensional task runner with visual pipeline builder"
        Emoji = "⚡"
    }
)

$Links = @(
    @{ Name = "GitHub"; Url = "https://github.com/yourusername"; Icon = "⚡" },
    @{ Name = "LinkedIn"; Url = "https://linkedin.com/in/yourprofile"; Icon = "💼" },
    @{ Name = "Portfolio"; Url = "https://yourblog.com"; Icon = "🌐" },
    @{ Name = "Email"; Url = "mailto:your.email@example.com"; Icon = "✉️" }
)

# ========================
# GLASSMORPHISM COLOR SCHEME
# ========================

# Background with gradient
$ColorBgStart = [System.Drawing.Color]::FromArgb(10, 10, 15)
$ColorBgEnd = [System.Drawing.Color]::FromArgb(25, 15, 35)

# Glass panels with transparency
$ColorGlass = [System.Drawing.Color]::FromArgb(40, 255, 255, 255)  # White with transparency
$ColorGlassDark = [System.Drawing.Color]::FromArgb(60, 20, 20, 30)  # Dark glass
$ColorGlassBorder = [System.Drawing.Color]::FromArgb(80, 255, 255, 255)

# Vibrant accents
$ColorAccent1 = [System.Drawing.Color]::FromArgb(138, 180, 248)  # Soft blue
$ColorAccent2 = [System.Drawing.Color]::FromArgb(180, 142, 248)  # Soft purple
$ColorAccent3 = [System.Drawing.Color]::FromArgb(248, 180, 218)  # Soft pink

# Text colors
$ColorText = [System.Drawing.Color]::FromArgb(245, 245, 250)
$ColorTextDim = [System.Drawing.Color]::FromArgb(180, 180, 200)
$ColorTextGlow = [System.Drawing.Color]::FromArgb(220, 220, 255)

# Glow effects
$ColorGlow = [System.Drawing.Color]::FromArgb(100, 138, 180, 248)

# ========================
# CUSTOM GLASS PANEL CLASS
# ========================

Add-Type @"
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

public class GlassPanel : Panel {
    private Color glassColor = Color.FromArgb(40, 255, 255, 255);
    private Color borderColor = Color.FromArgb(80, 255, 255, 255);
    private int borderRadius = 16;
    private bool isHovered = false;
    
    public GlassPanel() {
        this.DoubleBuffered = true;
        this.BackColor = Color.Transparent;
    }
    
    public Color GlassColor {
        get { return glassColor; }
        set { glassColor = value; Invalidate(); }
    }
    
    public int BorderRadius {
        get { return borderRadius; }
        set { borderRadius = value; Invalidate(); }
    }
    
    public bool IsHovered {
        get { return isHovered; }
        set { isHovered = value; Invalidate(); }
    }
    
    protected override void OnPaint(PaintEventArgs e) {
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        
        Rectangle bounds = new Rectangle(1, 1, this.Width - 2, this.Height - 2);
        
        // Create rounded rectangle path
        GraphicsPath path = CreateRoundedRectanglePath(bounds, borderRadius);
        
        // Glass fill with slight gradient
        Color fillColor = isHovered 
            ? Color.FromArgb(Math.Min(glassColor.A + 20, 255), glassColor.R, glassColor.G, glassColor.B)
            : glassColor;
            
        using (LinearGradientBrush brush = new LinearGradientBrush(
            bounds, 
            fillColor,
            Color.FromArgb(fillColor.A - 10, fillColor.R, fillColor.G, fillColor.B),
            LinearGradientMode.Vertical)) {
            e.Graphics.FillPath(brush, path);
        }
        
        // Border with glow effect
        Color bColor = isHovered 
            ? Color.FromArgb(Math.Min(borderColor.A + 40, 255), borderColor.R, borderColor.G, borderColor.B)
            : borderColor;
            
        using (Pen pen = new Pen(bColor, 1.5f)) {
            e.Graphics.DrawPath(pen, path);
        }
        
        // Inner highlight
        Rectangle innerBounds = new Rectangle(bounds.X + 1, bounds.Y + 1, bounds.Width - 2, bounds.Height / 3);
        GraphicsPath innerPath = CreateRoundedRectanglePath(innerBounds, borderRadius - 2);
        
        using (LinearGradientBrush innerBrush = new LinearGradientBrush(
            innerBounds,
            Color.FromArgb(30, 255, 255, 255),
            Color.FromArgb(0, 255, 255, 255),
            LinearGradientMode.Vertical)) {
            e.Graphics.FillPath(innerBrush, innerPath);
        }
        
        base.OnPaint(e);
    }
    
    private GraphicsPath CreateRoundedRectanglePath(Rectangle bounds, int radius) {
        int diameter = radius * 2;
        Rectangle arc = new Rectangle(bounds.Location, new Size(diameter, diameter));
        GraphicsPath path = new GraphicsPath();
        
        path.AddArc(arc, 180, 90);
        arc.X = bounds.Right - diameter;
        path.AddArc(arc, 270, 90);
        arc.Y = bounds.Bottom - diameter;
        path.AddArc(arc, 0, 90);
        arc.X = bounds.Left;
        path.AddArc(arc, 90, 90);
        path.CloseFigure();
        
        return path;
    }
}
"@ -ReferencedAssemblies System.Drawing, System.Windows.Forms

# ========================
# ANIMATED BACKGROUND PANEL
# ========================

$script:gradientOffset = 0
$script:animationTimer = New-Object System.Windows.Forms.Timer
$script:animationTimer.Interval = 50

# ========================
# UI SETUP
# ========================

$form = New-Object System.Windows.Forms.Form
$form.Text = "◈ $Name — Interactive Portfolio"
$form.Size = New-Object System.Drawing.Size(1200, 750)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = 'None'
$form.BackColor = $ColorBgStart
$form.Opacity = 0.98

# Custom title bar for drag
$script:isDragging = $false
$script:dragOffset = New-Object System.Drawing.Point

# Background panel with gradient
$bgPanel = New-Object System.Windows.Forms.Panel
$bgPanel.Dock = "Fill"
$bgPanel.BackColor = $ColorBgStart

$bgPanel.Add_Paint({
    param($sender, $e)
    
    $rect = $sender.ClientRectangle
    
    # Animated gradient background
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        $rect,
        $ColorBgStart,
        $ColorBgEnd,
        (45 + $script:gradientOffset) % 360
    )
    
    $e.Graphics.FillRectangle($brush, $rect)
    
    # Add floating orbs
    $e.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    
    # Orb 1
    $orb1Rect = New-Object System.Drawing.Rectangle(
        (100 + ([Math]::Sin($script:gradientOffset / 20) * 30)),
        (100 + ([Math]::Cos($script:gradientOffset / 15) * 20)),
        300, 300
    )
    $orbBrush1 = New-Object System.Drawing.Drawing2D.PathGradientBrush(
        [System.Drawing.Drawing2D.GraphicsPath]::new()
    )
    $orbBrush1.CenterColor = [System.Drawing.Color]::FromArgb(60, $ColorAccent1.R, $ColorAccent1.G, $ColorAccent1.B)
    $orbBrush1.SurroundColors = @([System.Drawing.Color]::FromArgb(0, $ColorAccent1.R, $ColorAccent1.G, $ColorAccent1.B))
    $e.Graphics.FillEllipse($orbBrush1, $orb1Rect)
    
    # Orb 2
    $orb2Rect = New-Object System.Drawing.Rectangle(
        (700 + ([Math]::Cos($script:gradientOffset / 18) * 40)),
        (400 + ([Math]::Sin($script:gradientOffset / 22) * 30)),
        400, 400
    )
    $orbBrush2 = New-Object System.Drawing.Drawing2D.PathGradientBrush(
        [System.Drawing.Drawing2D.GraphicsPath]::new()
    )
    $orbBrush2.CenterColor = [System.Drawing.Color]::FromArgb(40, $ColorAccent2.R, $ColorAccent2.G, $ColorAccent2.B)
    $orbBrush2.SurroundColors = @([System.Drawing.Color]::FromArgb(0, $ColorAccent2.R, $ColorAccent2.G, $ColorAccent2.B))
    $e.Graphics.FillEllipse($orbBrush2, $orb2Rect)
    
    $brush.Dispose()
})

$form.Controls.Add($bgPanel)

# Animation timer for gradient
$script:animationTimer.Add_Tick({
    $script:gradientOffset += 0.5
    $bgPanel.Invalidate()
})
$script:animationTimer.Start()

# ========================
# CUSTOM TITLE BAR
# ========================

$titleBar = New-Object GlassPanel
$titleBar.Size = New-Object System.Drawing.Size(1200, 50)
$titleBar.Location = New-Object System.Drawing.Point(0, 0)
$titleBar.GlassColor = [System.Drawing.Color]::FromArgb(30, 255, 255, 255)
$titleBar.BorderRadius = 0
$titleBar.Cursor = [System.Windows.Forms.Cursors]::SizeAll

# Title text
$titleText = New-Object System.Windows.Forms.Label
$titleText.Text = "◈ PORTFOLIO"
$titleText.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
$titleText.ForeColor = $ColorText
$titleText.BackColor = [System.Drawing.Color]::Transparent
$titleText.Location = New-Object System.Drawing.Point(20, 15)
$titleText.AutoSize = $true
$titleBar.Controls.Add($titleText)

# Close button
$closeBtn = New-Object System.Windows.Forms.Label
$closeBtn.Text = "✕"
$closeBtn.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$closeBtn.ForeColor = $ColorTextDim
$closeBtn.BackColor = [System.Drawing.Color]::Transparent
$closeBtn.Location = New-Object System.Drawing.Point(1160, 12)
$closeBtn.Size = New-Object System.Drawing.Size(30, 30)
$closeBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
$closeBtn.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$closeBtn.Add_Click({ $form.Close() })
$closeBtn.Add_MouseEnter({ $this.ForeColor = $ColorAccent3 })
$closeBtn.Add_MouseLeave({ $this.ForeColor = $ColorTextDim })
$titleBar.Controls.Add($closeBtn)

# Minimize button
$minBtn = New-Object System.Windows.Forms.Label
$minBtn.Text = "—"
$minBtn.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$minBtn.ForeColor = $ColorTextDim
$minBtn.BackColor = [System.Drawing.Color]::Transparent
$minBtn.Location = New-Object System.Drawing.Point(1120, 12)
$minBtn.Size = New-Object System.Drawing.Size(30, 30)
$minBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
$minBtn.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$minBtn.Add_Click({ $form.WindowState = 'Minimized' })
$minBtn.Add_MouseEnter({ $this.ForeColor = $ColorAccent1 })
$minBtn.Add_MouseLeave({ $this.ForeColor = $ColorTextDim })
$titleBar.Controls.Add($minBtn)

# Drag functionality
$titleBar.Add_MouseDown({
    param($sender, $e)
    $script:isDragging = $true
    $script:dragOffset = New-Object System.Drawing.Point($e.X, $e.Y)
})

$titleBar.Add_MouseMove({
    param($sender, $e)
    if ($script:isDragging) {
        $form.Location = New-Object System.Drawing.Point(
            ($form.Location.X + $e.X - $script:dragOffset.X),
            ($form.Location.Y + $e.Y - $script:dragOffset.Y)
        )
    }
})

$titleBar.Add_MouseUp({
    $script:isDragging = $false
})

$bgPanel.Controls.Add($titleBar)

# ========================
# GLASS SIDEBAR
# ========================

$sidebar = New-Object GlassPanel
$sidebar.Size = New-Object System.Drawing.Size(320, 650)
$sidebar.Location = New-Object System.Drawing.Point(30, 70)
$sidebar.GlassColor = [System.Drawing.Color]::FromArgb(50, 255, 255, 255)
$sidebar.BorderRadius = 20
$sidebar.Padding = New-Object System.Windows.Forms.Padding(30)

# Profile section
$profilePanel = New-Object System.Windows.Forms.Panel
$profilePanel.Size = New-Object System.Drawing.Size(260, 160)
$profilePanel.Location = New-Object System.Drawing.Point(30, 30)
$profilePanel.BackColor = [System.Drawing.Color]::Transparent

# Avatar placeholder (glowing circle)
$avatar = New-Object System.Windows.Forms.Panel
$avatar.Size = New-Object System.Drawing.Size(80, 80)
$avatar.Location = New-Object System.Drawing.Point(90, 0)
$avatar.BackColor = [System.Drawing.Color]::Transparent

$avatar.Add_Paint({
    param($sender, $e)
    $e.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    
    # Outer glow
    $glowBrush = New-Object System.Drawing.Drawing2D.PathGradientBrush(
        [System.Drawing.Drawing2D.GraphicsPath]::new()
    )
    $glowBrush.CenterColor = [System.Drawing.Color]::FromArgb(150, $ColorAccent1.R, $ColorAccent1.G, $ColorAccent1.B)
    $glowBrush.SurroundColors = @([System.Drawing.Color]::FromArgb(0, $ColorAccent1.R, $ColorAccent1.G, $ColorAccent1.B))
    $e.Graphics.FillEllipse($glowBrush, -5, -5, 90, 90)
    
    # Main circle
    $gradBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.Rectangle(0, 0, 80, 80)),
        $ColorAccent1,
        $ColorAccent2,
        45
    )
    $e.Graphics.FillEllipse($gradBrush, 0, 0, 80, 80)
    
    # Inner highlight
    $highlightBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.Rectangle(10, 10, 40, 40)),
        [System.Drawing.Color]::FromArgb(80, 255, 255, 255),
        [System.Drawing.Color]::FromArgb(0, 255, 255, 255),
        45
    )
    $e.Graphics.FillEllipse($highlightBrush, 10, 10, 40, 40)
    
    $gradBrush.Dispose()
    $highlightBrush.Dispose()
    $glowBrush.Dispose()
})

$profilePanel.Controls.Add($avatar)

# Name
$nameLabel = New-Object System.Windows.Forms.Label
$nameLabel.Text = $Name
$nameLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$nameLabel.ForeColor = $ColorText
$nameLabel.BackColor = [System.Drawing.Color]::Transparent
$nameLabel.Location = New-Object System.Drawing.Point(0, 90)
$nameLabel.Size = New-Object System.Drawing.Size(260, 30)
$nameLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$profilePanel.Controls.Add($nameLabel)

# Tagline
$taglineLabel = New-Object System.Windows.Forms.Label
$taglineLabel.Text = $Tagline
$taglineLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$taglineLabel.ForeColor = $ColorTextDim
$taglineLabel.BackColor = [System.Drawing.Color]::Transparent
$taglineLabel.Location = New-Object System.Drawing.Point(0, 125)
$taglineLabel.Size = New-Object System.Drawing.Size(260, 35)
$taglineLabel.TextAlign = [System.Drawing.ContentAlignment]::TopCenter
$profilePanel.Controls.Add($taglineLabel)

$sidebar.Controls.Add($profilePanel)

# ========================
# MAIN CONTENT AREA
# ========================

$mainContainer = New-Object GlassPanel
$mainContainer.Size = New-Object System.Drawing.Size(780, 650)
$mainContainer.Location = New-Object System.Drawing.Point(370, 70)
$mainContainer.GlassColor = [System.Drawing.Color]::FromArgb(40, 255, 255, 255)
$mainContainer.BorderRadius = 20
$mainContainer.AutoScroll = $true

$contentPanel = New-Object System.Windows.Forms.Panel
$contentPanel.Location = New-Object System.Drawing.Point(40, 30)
$contentPanel.Size = New-Object System.Drawing.Size(700, 600)
$contentPanel.BackColor = [System.Drawing.Color]::Transparent
$contentPanel.AutoSize = $true

$mainContainer.Controls.Add($contentPanel)
$bgPanel.Controls.Add($mainContainer)
$bgPanel.Controls.Add($sidebar)

# ========================
# HELPER FUNCTIONS
# ========================

function Clear-Content {
    $contentPanel.Controls.Clear()
}

function New-GlassCard([string]$title, [string]$desc, [string]$emoji, [string]$url, [int]$y) {
    $card = New-Object GlassPanel
    $card.Size = New-Object System.Drawing.Size(680, 110)
    $card.Location = New-Object System.Drawing.Point(0, $y)
    $card.GlassColor = [System.Drawing.Color]::FromArgb(30, 255, 255, 255)
    $card.BorderRadius = 16
    $card.Cursor = [System.Windows.Forms.Cursors]::Hand
    
    # Emoji
    $emojiLabel = New-Object System.Windows.Forms.Label
    $emojiLabel.Text = $emoji
    $emojiLabel.Font = New-Object System.Drawing.Font("Segoe UI Emoji", 28)
    $emojiLabel.BackColor = [System.Drawing.Color]::Transparent
    $emojiLabel.Location = New-Object System.Drawing.Point(25, 30)
    $emojiLabel.Size = New-Object System.Drawing.Size(50, 50)
    $card.Controls.Add($emojiLabel)
    
    # Title
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = $title
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = $ColorText
    $titleLabel.BackColor = [System.Drawing.Color]::Transparent
    $titleLabel.Location = New-Object System.Drawing.Point(90, 25)
    $titleLabel.Size = New-Object System.Drawing.Size(550, 25)
    $card.Controls.Add($titleLabel)
    
    # Description
    $descLabel = New-Object System.Windows.Forms.Label
    $descLabel.Text = $desc
    $descLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $descLabel.ForeColor = $ColorTextDim
    $descLabel.BackColor = [System.Drawing.Color]::Transparent
    $descLabel.Location = New-Object System.Drawing.Point(90, 55)
    $descLabel.Size = New-Object System.Drawing.Size(550, 40)
    $card.Controls.Add($descLabel)
    
    # Hover glow
    $card.Add_MouseEnter({
        $this.IsHovered = $true
    })
    $card.Add_MouseLeave({
        $this.IsHovered = $false
    })
    
    # Click
    $card.Add_Click({ Start-Process $url })
    foreach ($ctrl in $card.Controls) {
        $ctrl.Add_Click({ Start-Process $url })
    }
    
    return $card
}

function New-SkillPill([string]$text, [int]$x, [int]$y) {
    $pill = New-Object GlassPanel
    $pill.GlassColor = [System.Drawing.Color]::FromArgb(50, $ColorAccent1.R, $ColorAccent1.G, $ColorAccent1.B)
    $pill.BorderRadius = 20
    $pill.AutoSize = $false
    
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $text
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $label.ForeColor = $ColorText
    $label.BackColor = [System.Drawing.Color]::Transparent
    $label.AutoSize = $true
    $label.Padding = New-Object System.Windows.Forms.Padding(16, 8, 16, 8)
    
    $pill.Size = New-Object System.Drawing.Size(($label.PreferredWidth), ($label.PreferredHeight))
    $pill.Location = New-Object System.Drawing.Point($x, $y)
    $pill.Controls.Add($label)
    
    return $pill
}

function New-LinkCard([string]$name, [string]$icon, [string]$url, [int]$x, [int]$y) {
    $card = New-Object GlassPanel
    $card.Size = New-Object System.Drawing.Size(320, 70)
    $card.Location = New-Object System.Drawing.Point($x, $y)
    $card.GlassColor = [System.Drawing.Color]::FromArgb(30, 255, 255, 255)
    $card.BorderRadius = 14
    $card.Cursor = [System.Windows.Forms.Cursors]::Hand
    
    # Icon
    $iconLabel = New-Object System.Windows.Forms.Label
    $iconLabel.Text = $icon
    $iconLabel.Font = New-Object System.Drawing.Font("Segoe UI Emoji", 20)
    $iconLabel.BackColor = [System.Drawing.Color]::Transparent
    $iconLabel.Location = New-Object System.Drawing.Point(20, 22)
    $iconLabel.Size = New-Object System.Drawing.Size(40, 30)
    $card.Controls.Add($iconLabel)
    
    # Name
    $nameLabel = New-Object System.Windows.Forms.Label
    $nameLabel.Text = $name
    $nameLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $nameLabel.ForeColor = $ColorText
    $nameLabel.BackColor = [System.Drawing.Color]::Transparent
    $nameLabel.Location = New-Object System.Drawing.Point(70, 22)
    $nameLabel.Size = New-Object System.Drawing.Size(220, 25)
    $card.Controls.Add($nameLabel)
    
    $card.Add_MouseEnter({ $this.IsHovered = $true })
    $card.Add_MouseLeave({ $this.IsHovered = $false })
    $card.Add_Click({ Start-Process $url })
    
    foreach ($ctrl in $card.Controls) {
        $ctrl.Add_Click({ Start-Process $url })
    }
    
    return $card
}

function New-NavBtn([string]$text, [string]$icon, [int]$y, $action) {
    $btn = New-Object GlassPanel
    $btn.Size = New-Object System.Drawing.Size(260, 55)
    $btn.Location = New-Object System.Drawing.Point(30, $y)
    $btn.GlassColor = [System.Drawing.Color]::FromArgb(20, 255, 255, 255)
    $btn.BorderRadius = 12
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    
    # Icon
    $iconLabel = New-Object System.Windows.Forms.Label
    $iconLabel.Text = $icon
    $iconLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16)
    $iconLabel.ForeColor = $ColorAccent1
    $iconLabel.BackColor = [System.Drawing.Color]::Transparent
    $iconLabel.Location = New-Object System.Drawing.Point(20, 15)
    $iconLabel.Size = New-Object System.Drawing.Size(30, 25)
    $btn.Controls.Add($iconLabel)
    
    # Text
    $textLabel = New-Object System.Windows.Forms.Label
    $textLabel.Text = $text
    $textLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $textLabel.ForeColor = $ColorText
    $textLabel.BackColor = [System.Drawing.Color]::Transparent
    $textLabel.Location = New-Object System.Drawing.Point(60, 17)
    $textLabel.Size = New-Object System.Drawing.Size(180, 22)
    $btn.Controls.Add($textLabel)
    
    $btn.Add_MouseEnter({ 
        $this.IsHovered = $true
    })
    $btn.Add_MouseLeave({ 
        $this.IsHovered = $false
    })
    $btn.Add_Click($action)
    
    foreach ($ctrl in $btn.Controls) {
        $ctrl.Add_Click($action)
    }
    
    return $btn
}

function Set-ActiveNav($activeBtn) {
    foreach ($ctrl in $sidebar.Controls) {
        if ($ctrl -is [GlassPanel] -and $ctrl.Size.Height -eq 55) {
            $ctrl.GlassColor = [System.Drawing.Color]::FromArgb(20, 255, 255, 255)
        }
    }
    $activeBtn.GlassColor = [System.Drawing.Color]::FromArgb(60, $ColorAccent1.R, $ColorAccent1.G, $ColorAccent1.B)
}

# ========================
# CONTENT PAGES
# ========================

function Show-About {
    Clear-Content
    Set-ActiveNav $script:navAbout
    
    $y = 0
    
    # Animated title
    $title = New-Object System.Windows.Forms.Label
    $title.Text = "✨ About Me"
    $title.Font = New-Object System.Drawing.Font("Segoe UI", 28, [System.Drawing.FontStyle]::Bold)
    $title.ForeColor = $ColorText
    $title.BackColor = [System.Drawing.Color]::Transparent
    $title.Location = New-Object System.Drawing.Point(0, $y)
    $title.AutoSize = $true
    $contentPanel.Controls.Add($title)
    $y += 60
    
    # Bio
    $bioLabel = New-Object System.Windows.Forms.Label
    $bioLabel.Text = $Bio
    $bioLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11)
    $bioLabel.ForeColor = $ColorTextDim
    $bioLabel.BackColor = [System.Drawing.Color]::Transparent
    $bioLabel.Location = New-Object System.Drawing.Point(0, $y)
    $bioLabel.MaximumSize = New-Object System.Drawing.Size(660, 0)
    $bioLabel.AutoSize = $true
    $contentPanel.Controls.Add($bioLabel)
    $y += $bioLabel.Height + 50
    
    # Skills title
    $skillsTitle = New-Object System.Windows.Forms.Label
    $skillsTitle.Text = "⚡ Core Competencies"
    $skillsTitle.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $skillsTitle.ForeColor = $ColorText
    $skillsTitle.BackColor = [System.Drawing.Color]::Transparent
    $skillsTitle.Location = New-Object System.Drawing.Point(0, $y)
    $skillsTitle.AutoSize = $true
    $contentPanel.Controls.Add($skillsTitle)
    $y += 50
    
    # Skills
    $x = 0
    $rowY = $y
    foreach ($skill in $Skills) {
        $pill = New-SkillPill $skill $x $rowY
        $contentPanel.Controls.Add($pill)
        $x += $pill.Width + 12
        
        if ($x -gt 480) {
            $x = 0
            $rowY += 50
        }
    }
}

function Show-Projects {
    Clear-Content
    Set-ActiveNav $script:navProjects
    
    $y = 0
    
    $title = New-Object System.Windows.Forms.Label
    $title.Text = "🚀 Featured Projects"
    $title.Font = New-Object System.Drawing.Font("Segoe UI", 28, [System.Drawing.FontStyle]::Bold)
    $title.ForeColor = $ColorText
    $title.BackColor = [System.Drawing.Color]::Transparent
    $title.Location = New-Object System.Drawing.Point(0, $y)
    $title.AutoSize = $true
    $contentPanel.Controls.Add($title)
    $y += 70
    
    foreach ($project in $Projects) {
        $card = New-GlassCard $project.Name $project.Desc $project.Emoji $project.Url $y
        $contentPanel.Controls.Add($card)
        $y += 125
    }
}

function Show-Contact {
    Clear-Content
    Set-ActiveNav $script:navContact
    
    $y = 0
    
    $title = New-Object System.Windows.Forms.Label
    $title.Text = "💬 Let's Connect"
    $title.Font = New-Object System.Drawing.Font("Segoe UI", 28, [System.Drawing.FontStyle]::Bold)
    $title.ForeColor = $ColorText
    $title.BackColor = [System.Drawing.Color]::Transparent
    $title.Location = New-Object System.Drawing.Point(0, $y)
    $title.AutoSize = $true
    $contentPanel.Controls.Add($title)
    $y += 60
    
    $desc = New-Object System.Windows.Forms.Label
    $desc.Text = "Always excited to collaborate on innovative projects or discuss cutting-edge automation solutions. Drop me a line!"
    $desc.Font = New-Object System.Drawing.Font("Segoe UI", 11)
    $desc.ForeColor = $ColorTextDim
    $desc.BackColor = [System.Drawing.Color]::Transparent
    $desc.Location = New-Object System.Drawing.Point(0, $y)
    $desc.MaximumSize = New-Object System.Drawing.Size(660, 0)
    $desc.AutoSize = $true
    $contentPanel.Controls.Add($desc)
    $y += $desc.Height + 40
    
    # Links in grid
    $row = 0
    $col = 0
    foreach ($link in $Links) {
        $x = $col * 340
        $linkY = $y + ($row * 85)
        
        $card = New-LinkCard $link.Name $link.Icon $link.Url $x $linkY
        $contentPanel.Controls.Add($card)
        
        $col++
        if ($col -ge 2) {
            $col = 0
            $row++
        }
    }
}

# ========================
# NAVIGATION
# ========================

$script:navAbout = New-NavBtn "About" "👤" 220 { Show-About }
$script:navProjects = New-NavBtn "Projects" "⚡" 285 { Show-Projects }
$script:navContact = New-NavBtn "Contact" "💌" 350 { Show-Contact }

$sidebar.Controls.Add($script:navAbout)
$sidebar.Controls.Add($script:navProjects)
$sidebar.Controls.Add($script:navContact)

# Footer
$footer = New-Object System.Windows.Forms.Label
$footer.Text = "Powered by PowerShell`n& Modern Design Principles"
$footer.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$footer.ForeColor = $ColorTextDim
$footer.BackColor = [System.Drawing.Color]::Transparent
$footer.Location = New-Object System.Drawing.Point(30, 570)
$footer.Size = New-Object System.Drawing.Size(260, 40)
$footer.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$sidebar.Controls.Add($footer)

# ========================
# LAUNCH
# ========================

Show-About

$form.Add_Shown({ 
    $form.Activate()
})

$form.Add_FormClosing({
    $script:animationTimer.Stop()
    $script:animationTimer.Dispose()
})

[void]$form.ShowDialog()
