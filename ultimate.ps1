<#PSScriptInfo

.VERSION 2.2.0

.GUID 8f696a63-fc1a-416a-bc32-6b4e52898647

.AUTHOR NotNahid

.DESCRIPTION 
Dynamic utility launcher that downloads and runs various tools.

#>

Param()

# ╔════════════════════════════════════════════════════════════════════════════╗
# ║                                                                            ║
# ║   NAHID POWER UTILITIES - MAIN SCRIPT                                      ║
# ║                                                                            ║
# ║   Sections:                                                                ║
# ║     1. IMPORTS & SETUP                                                     ║
# ║     2. DATA LOADING                                                        ║
# ║     3. DATA VALIDATION                                                     ║
# ║     4. THEME CONFIGURATION                                                 ║
# ║     5. HELPER FUNCTIONS                                                    ║
# ║     6. DIALOG FUNCTIONS                                                    ║
# ║     7. UTILITY LAUNCHER (WITH REAL-TIME STATUS)                            ║
# ║     8. THEME FUNCTIONS                                                     ║
# ║     9. UI CARD BUILDER                                                     ║
# ║    10. MAIN FORM SETUP                                                     ║
# ║    11. HEADER SECTION                                                      ║
# ║    12. UTILITIES PANEL                                                     ║
# ║    13. STATUS BAR                                                          ║
# ║    14. DISPLAY UPDATE FUNCTION                                             ║
# ║    15. EVENT HANDLERS                                                      ║
# ║    16. APPLICATION LAUNCH                                                  ║
# ║                                                                            ║
# ╚════════════════════════════════════════════════════════════════════════════╝


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1: IMPORTS & SETUP
# ══════════════════════════════════════════════════════════════════════════════
# Load required .NET assemblies for Windows Forms GUI

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Enable visual styles for modern look
[System.Windows.Forms.Application]::EnableVisualStyles()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2: DATA LOADING
# ══════════════════════════════════════════════════════════════════════════════
# Fetch utilities database from remote JSON or use fallback

$configUrl = "https://raw.githubusercontent.com/NotNahid/powershell/refs/heads/main/Main%20Links%20File%20Json/utilities.json"

try {
    $Utilities = Invoke-RestMethod -Uri $configUrl -UseBasicParsing
    Write-Host "Loaded $($Utilities.Count) utilities from online database." -ForegroundColor Green
}
catch {
    Write-Warning "Failed to load online utility database. Using fallback."
    $Utilities = @(
        [PSCustomObject]@{
            Name     = "Ghost Typer"
            Link     = "https://gist.githubusercontent.com/NotNahid/70de29086ac0bbeab57a9a75a4f04a89/raw/copy-script.ps1"
            Desc     = "Auto-type text anywhere"
            Category = "Automation"
            Tags     = @("typing", "automation", "productivity")
        }
    )
}


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3: DATA VALIDATION
# ══════════════════════════════════════════════════════════════════════════════
# Validate each utility has required fields before displaying

function Test-Utilities {
    param($RawUtilities)

    $validatedUtilities = @()
    $requiredFields = @("Name", "Link", "Desc", "Category", "Tags")

    foreach ($util in $RawUtilities) {
        $isValid = $true
        $missingFields = @()

        foreach ($field in $requiredFields) {
            $prop = $util.PSObject.Properties[$field]
            $value = if ($null -ne $prop) { $prop.Value } else { $null }

            if ($null -eq $prop -or $null -eq $value -or
                ($value -is [string] -and [string]::IsNullOrWhiteSpace($value))) {
                $isValid = $false
                $missingFields += $field
            }
        }

        if (-not $isValid) {
            Write-Warning "Skipping utility: Missing fields [$($missingFields -join ', ')]"
            continue
        }

        $validatedUtilities += $util
    }

    return $validatedUtilities
}

$script:ValidatedUtilities = Test-Utilities -RawUtilities $Utilities

if ($script:ValidatedUtilities.Count -eq 0) {
    Write-Error "No valid utilities found. Exiting."
    exit 1
}


# ╔══════════════════════════════════════════════════════════╗
# ║       🎨 PREMIUM DARK THEME - Modern Fusion UI          ║
# ║   Inspired by: Catppuccin + Fluent + Discord + VS Code  ║
# ╚══════════════════════════════════════════════════════════╝

$DarkTheme = @{
    # --- Base Layers ---
    Background    = [System.Drawing.Color]::FromArgb(17, 17, 27)       # #11111B Deep crust (Catppuccin)
    Surface       = [System.Drawing.Color]::FromArgb(24, 24, 37)       # #181825 Mantle layer
    SurfaceLight  = [System.Drawing.Color]::FromArgb(30, 30, 46)       # #1E1E2E Base layer
    SurfaceHover  = [System.Drawing.Color]::FromArgb(49, 50, 68)       # #313244 Hover overlay

    # --- Accent Colors ---
    Primary       = [System.Drawing.Color]::FromArgb(137, 180, 250)    # #89B4FA Sapphire blue (Catppuccin)
    PrimaryHover  = [System.Drawing.Color]::FromArgb(116, 199, 236)    # #74C7EC Sky blue glow

    # --- Status Colors ---
    Success       = [System.Drawing.Color]::FromArgb(166, 227, 161)    # #A6E3A1 Mint green
    SuccessHover  = [System.Drawing.Color]::FromArgb(148, 226, 213)    # #94E2D5 Teal glow
    Warning       = [System.Drawing.Color]::FromArgb(249, 226, 175)    # #F9E2AF Warm peach
    Danger        = [System.Drawing.Color]::FromArgb(243, 139, 168)    # #F38BA8 Soft red/pink
    DangerHover   = [System.Drawing.Color]::FromArgb(235, 160, 172)    # #EBA0AC Flamingo

    # --- Typography ---
    Text          = [System.Drawing.Color]::FromArgb(205, 214, 244)    # #CDD6F4 Crisp white-lavender
    TextSecondary = [System.Drawing.Color]::FromArgb(166, 173, 200)    # #A6ADC8 Subtle overlay
    TextMuted     = [System.Drawing.Color]::FromArgb(108, 112, 134)    # #6C7086 Muted overlay

    # --- UI Elements ---
    Border        = [System.Drawing.Color]::FromArgb(69, 71, 90)       # #45475A Subtle separator
    SearchBox     = [System.Drawing.Color]::FromArgb(24, 24, 37)       # #181825 Input field

    # --- BONUS: Extra Accent Colors (use if needed) ---
    Accent1       = [System.Drawing.Color]::FromArgb(203, 166, 247)    # #CBA6F7 Mauve/Purple
    Accent2       = [System.Drawing.Color]::FromArgb(250, 179, 135)    # #FAB387 Peach/Orange
    Accent3       = [System.Drawing.Color]::FromArgb(137, 220, 235)    # #89DCEB Cyan
    Accent4       = [System.Drawing.Color]::FromArgb(245, 194, 231)    # #F5C2E7 Pink
    GradientStart = [System.Drawing.Color]::FromArgb(137, 180, 250)    # Blue gradient start
    GradientEnd   = [System.Drawing.Color]::FromArgb(203, 166, 247)    # Purple gradient end
}

# ╔══════════════════════════════════════════════════════════╗
# ║       ☀️ PREMIUM LIGHT THEME - Catppuccin Latte         ║
# ╚══════════════════════════════════════════════════════════╝

$LightTheme = @{
    # --- Base Layers ---
    Background    = [System.Drawing.Color]::FromArgb(239, 241, 245)    # #EFF1F5 Latte base
    Surface       = [System.Drawing.Color]::FromArgb(230, 233, 239)    # #E6E9EF Mantle
    SurfaceLight  = [System.Drawing.Color]::FromArgb(220, 224, 232)    # #DCE0E8 Crust
    SurfaceHover  = [System.Drawing.Color]::FromArgb(204, 208, 218)    # #CCD0DA Hover

    # --- Accent Colors ---
    Primary       = [System.Drawing.Color]::FromArgb(30, 102, 245)     # #1E66F5 Bold blue
    PrimaryHover  = [System.Drawing.Color]::FromArgb(4, 165, 229)      # #04A5E5 Sky blue

    # --- Status Colors ---
    Success       = [System.Drawing.Color]::FromArgb(64, 160, 43)      # #40A02B Rich green
    SuccessHover  = [System.Drawing.Color]::FromArgb(23, 146, 153)     # #179299 Teal
    Warning       = [System.Drawing.Color]::FromArgb(223, 142, 29)     # #DF8E1D Golden amber
    Danger        = [System.Drawing.Color]::FromArgb(210, 15, 57)      # #D20F39 Bold red
    DangerHover   = [System.Drawing.Color]::FromArgb(230, 69, 83)      # #E64553 Maroon

    # --- Typography ---
    Text          = [System.Drawing.Color]::FromArgb(76, 79, 105)      # #4C4F69 Rich dark text
    TextSecondary = [System.Drawing.Color]::FromArgb(92, 95, 119)      # #5C5F77 Subtle text
    TextMuted     = [System.Drawing.Color]::FromArgb(124, 127, 147)    # #7C7F93 Muted text

    # --- UI Elements ---
    Border        = [System.Drawing.Color]::FromArgb(188, 192, 204)    # #BCC0CC Clean border
    SearchBox     = [System.Drawing.Color]::FromArgb(230, 233, 239)    # #E6E9EF Input field

    # --- BONUS: Extra Accent Colors ---
    Accent1       = [System.Drawing.Color]::FromArgb(136, 57, 239)     # #8839EF Purple
    Accent2       = [System.Drawing.Color]::FromArgb(254, 100, 11)     # #FE640B Orange
    Accent3       = [System.Drawing.Color]::FromArgb(4, 165, 229)      # #04A5E5 Cyan
    Accent4       = [System.Drawing.Color]::FromArgb(234, 118, 203)    # #EA76CB Pink
    GradientStart = [System.Drawing.Color]::FromArgb(30, 102, 245)     # Blue gradient
    GradientEnd   = [System.Drawing.Color]::FromArgb(136, 57, 239)     # Purple gradient
}

# --- Set Active Theme ---
$script:Colors = $DarkTheme

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5: HELPER FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════
# Small utility functions used throughout the application

# --- Global Tooltip Object ---
$script:tooltip = New-Object System.Windows.Forms.ToolTip
$script:tooltip.AutoPopDelay = 5000
$script:tooltip.InitialDelay = 300
$script:tooltip.ReshowDelay = 100
$script:tooltip.ShowAlways = $true
$script:tooltip.IsBalloon = $false

# --- Add Tooltip to Control ---
function Add-Tooltip {
    param(
        $Control,
        [string]$Text
    )
    
    try {
        if ($null -ne $Control -and $null -ne $script:tooltip) {
            $script:tooltip.SetToolTip($Control, $Text)
        }
    }
    catch {
        Write-Warning "Tooltip error: $_"
    }
}

# --- Update Status Bar Text ---
function Update-StatusBar {
    param([string]$Message)
    
    try {
        if ($null -ne $script:statusLabel) {
            $script:statusLabel.Text = $Message
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    catch { }
}

# --- Update Button State (Thread-Safe) ---
function Set-ButtonState {
    param(
        $Button,
        [string]$Text,
        [System.Drawing.Color]$BackColor,
        [System.Drawing.Color]$ForeColor = [System.Drawing.Color]::White,
        [bool]$Enabled = $false
    )
    
    if ($null -eq $Button) { return }
    
    try {
        $Button.Text = $Text
        $Button.BackColor = $BackColor
        $Button.ForeColor = $ForeColor
        $Button.Enabled = $Enabled
        [System.Windows.Forms.Application]::DoEvents()
    }
    catch {
        Write-Warning "Button state error: $_"
    }
}


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 6: DIALOG FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════
# Custom dialog boxes for confirmations and utility details

# --- Generic Custom Dialog ---
function Show-CustomDialog {
    param(
        [string]$Title,
        [string]$Message,
        [array]$Buttons = @("OK"),
        [string]$DefaultButton = "OK",
        $ParentForm
    )

    try {
        $dialog = New-Object System.Windows.Forms.Form
        $dialog.Text = $Title
        $dialog.Size = New-Object System.Drawing.Size(480, 220)
        $dialog.StartPosition = "CenterScreen"
        $dialog.FormBorderStyle = "FixedDialog"
        $dialog.BackColor = $script:Colors.Surface
        $dialog.ShowInTaskbar = $false
        $dialog.MaximizeBox = $false
        $dialog.MinimizeBox = $false
        $dialog.KeyPreview = $true

        # --- Message Panel ---
        $messagePanel = New-Object System.Windows.Forms.Panel
        $messagePanel.Location = New-Object System.Drawing.Point(20, 20)
        $messagePanel.Size = New-Object System.Drawing.Size(440, 110)
        $messagePanel.BackColor = $script:Colors.Surface

        $messageLabel = New-Object System.Windows.Forms.Label
        $messageLabel.Text = $Message
        $messageLabel.Dock = "Fill"
        $messageLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $messageLabel.ForeColor = $script:Colors.Text
        $messageLabel.TextAlign = "MiddleLeft"
        $messageLabel.Padding = New-Object System.Windows.Forms.Padding(10)

        $messagePanel.Controls.Add($messageLabel)

        # --- Button Panel ---
        $buttonPanel = New-Object System.Windows.Forms.Panel
        $buttonPanel.Dock = "Bottom"
        $buttonPanel.Height = 60
        $buttonPanel.BackColor = $script:Colors.Surface

        $buttonWidth = 100
        $buttonSpacing = 10
        $totalButtonWidth = ($Buttons.Count * $buttonWidth) + (($Buttons.Count - 1) * $buttonSpacing)
        $startX = ($dialog.Width - $totalButtonWidth) / 2

        $firstButton = $null

        for ($i = 0; $i -lt $Buttons.Count; $i++) {
            $btn = New-Object System.Windows.Forms.Button
            $btn.Text = $Buttons[$i]
            $btn.Size = New-Object System.Drawing.Size($buttonWidth, 36)
            $btn.Location = New-Object System.Drawing.Point(($startX + ($i * ($buttonWidth + $buttonSpacing))), 12)
            $btn.FlatStyle = "Flat"
            $btn.FlatAppearance.BorderSize = 1
            $btn.Font = New-Object System.Drawing.Font("Segoe UI", 10)
            $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
            $btn.Tag = $Buttons[$i]
            $btn.TabIndex = $i

            if ($Buttons[$i] -match "Yes|OK|Confirm|Launch|Run") {
                $btn.BackColor = $script:Colors.Primary
                $btn.ForeColor = [System.Drawing.Color]::White
                $btn.FlatAppearance.BorderColor = $script:Colors.Primary
            }
            else {
                $btn.BackColor = $script:Colors.SurfaceLight
                $btn.ForeColor = $script:Colors.Text
                $btn.FlatAppearance.BorderColor = $script:Colors.Border
            }

            $capturedDialog = $dialog
            $btn.Add_Click({
                param($sender, $e)
                $capturedDialog.Tag = $sender.Tag
                $capturedDialog.Close()
            })

            $btn.Add_MouseEnter({
                if ($this.Text -match "Yes|OK|Confirm|Launch|Run") {
                    $this.BackColor = $script:Colors.PrimaryHover
                }
                else {
                    $this.BackColor = $script:Colors.SurfaceHover
                }
            })

            $btn.Add_MouseLeave({
                if ($this.Text -match "Yes|OK|Confirm|Launch|Run") {
                    $this.BackColor = $script:Colors.Primary
                }
                else {
                    $this.BackColor = $script:Colors.SurfaceLight
                }
            })

            $buttonPanel.Controls.Add($btn)

            if ($i -eq 0) { $firstButton = $btn }
            if ($Buttons[$i] -eq $DefaultButton) { $dialog.AcceptButton = $btn }
            if ($Buttons[$i] -match "Cancel|No") { $dialog.CancelButton = $btn }
        }

        $dialog.Add_KeyDown({
            param($sender, $e)
            if ($e.KeyCode -eq "Escape" -and $null -eq $dialog.CancelButton) {
                $dialog.Tag = "Cancel"
                $dialog.Close()
                $e.Handled = $true
            }
        })

        $dialog.Controls.AddRange(@($buttonPanel, $messagePanel))

        $dialog.Add_Shown({
            if ($null -ne $firstButton) { $firstButton.Focus() }
        })

        [void]$dialog.ShowDialog($ParentForm)

        return $dialog.Tag
    }
    catch {
        Write-Error "Dialog error: $_"
        return $null
    }
}

# --- Utility Details Dialog ---
function Show-UtilityDetailsDialog {
    param($Util, $ParentForm)

    try {
        if ($null -eq $Util) { return $null }

        $dialog = New-Object System.Windows.Forms.Form
        $dialog.Text = $Util.Name
        $dialog.Size = New-Object System.Drawing.Size(550, 400)
        $dialog.StartPosition = "CenterScreen"
        $dialog.FormBorderStyle = "FixedDialog"
        $dialog.BackColor = $script:Colors.Surface
        $dialog.ShowInTaskbar = $false
        $dialog.MaximizeBox = $false
        $dialog.MinimizeBox = $false
        $dialog.KeyPreview = $true

        # --- Content Panel ---
        $contentPanel = New-Object System.Windows.Forms.Panel
        $contentPanel.Location = New-Object System.Drawing.Point(30, 20)
        $contentPanel.Size = New-Object System.Drawing.Size(490, 280)
        $contentPanel.BackColor = $script:Colors.Surface

        $yPos = 0

        # Utility Name
        $nameLabel = New-Object System.Windows.Forms.Label
        $nameLabel.Text = $Util.Name
        $nameLabel.Location = New-Object System.Drawing.Point(0, $yPos)
        $nameLabel.Size = New-Object System.Drawing.Size(490, 35)
        $nameLabel.Font = $script:Fonts.Heading1
        $nameLabel.ForeColor = $script:Colors.Text
        $yPos += 45

        # Category
        $categoryLabel = New-Object System.Windows.Forms.Label
        $categoryLabel.Text = "Category: " + $Util.Category
        $categoryLabel.Location = New-Object System.Drawing.Point(0, $yPos)
        $categoryLabel.AutoSize = $true
        $categoryLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $categoryLabel.ForeColor = $script:Colors.Primary
        $yPos += 35

        # Description Header
        $descHeader = New-Object System.Windows.Forms.Label
        $descHeader.Text = "Description"
        $descHeader.Location = New-Object System.Drawing.Point(0, $yPos)
        $descHeader.Size = New-Object System.Drawing.Size(490, 20)
        $descHeader.Font = $script:Fonts.Heading3
        $descHeader.ForeColor = $script:Colors.Text

        $descText = New-Object System.Windows.Forms.Label
        $descText.Text = $Util.Desc
        $descText.Location = New-Object System.Drawing.Point(0, ($yPos + 25))
        $descText.Size = New-Object System.Drawing.Size(490, 40)
        $descText.Font = $script:Fonts.BodySmall
        $descText.ForeColor = $script:Colors.TextSecondary
        $yPos += 75

        # Tags
        $tagsHeader = New-Object System.Windows.Forms.Label
        $tagsHeader.Text = "Tags"
        $tagsHeader.Location = New-Object System.Drawing.Point(0, $yPos)
        $tagsHeader.Size = New-Object System.Drawing.Size(490, 20)
        $tagsHeader.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $tagsHeader.ForeColor = $script:Colors.Text

        $tagsValue = $Util.Tags
        $tagsString = if ($tagsValue -is [array]) { $tagsValue -join ', ' } else { "$tagsValue" }

        $tagsText = New-Object System.Windows.Forms.Label
        $tagsText.Text = $tagsString
        $tagsText.Location = New-Object System.Drawing.Point(0, ($yPos + 25))
        $tagsText.Size = New-Object System.Drawing.Size(490, 30)
        $tagsText.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $tagsText.ForeColor = $script:Colors.Primary
        $yPos += 65

        # Source
        $linkHeader = New-Object System.Windows.Forms.Label
        $linkHeader.Text = "Source"
        $linkHeader.Location = New-Object System.Drawing.Point(0, $yPos)
        $linkHeader.Size = New-Object System.Drawing.Size(490, 20)
        $linkHeader.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $linkHeader.ForeColor = $script:Colors.Text

        $linkText = New-Object System.Windows.Forms.Label
        $linkText.Text = $Util.Link
        $linkText.Location = New-Object System.Drawing.Point(0, ($yPos + 25))
        $linkText.Size = New-Object System.Drawing.Size(490, 30)
        $linkText.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $linkText.ForeColor = $script:Colors.TextSecondary

        $contentPanel.Controls.AddRange(@(
            $nameLabel, $categoryLabel,
            $descHeader, $descText,
            $tagsHeader, $tagsText,
            $linkHeader, $linkText
        ))

        # --- Button Panel ---
        $buttonPanel = New-Object System.Windows.Forms.Panel
        $buttonPanel.Dock = "Bottom"
        $buttonPanel.Height = 70
        $buttonPanel.BackColor = $script:Colors.SurfaceLight

        $runBtn = New-Object System.Windows.Forms.Button
        $runBtn.Text = "Launch"
        $runBtn.Size = New-Object System.Drawing.Size(120, 40)
        $runBtn.Location = New-Object System.Drawing.Point(215, 15)
        $runBtn.BackColor = $script:Colors.Success
        $runBtn.ForeColor = [System.Drawing.Color]::White
        $runBtn.FlatStyle = "Flat"
        $runBtn.FlatAppearance.BorderSize = 0
        $runBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $runBtn.Cursor = [System.Windows.Forms.Cursors]::Hand

        $runBtn.Add_Click({
            $dialog.Tag = "Run"
            $dialog.Close()
        })

        $runBtn.Add_MouseEnter({ $this.BackColor = $script:Colors.SuccessHover })
        $runBtn.Add_MouseLeave({ $this.BackColor = $script:Colors.Success })

        $buttonPanel.Controls.Add($runBtn)
        $dialog.Controls.AddRange(@($buttonPanel, $contentPanel))

        $dialog.Add_KeyDown({
            param($sender, $e)
            if ($e.KeyCode -eq "Enter") {
                $dialog.Tag = "Run"
                $dialog.Close()
                $e.Handled = $true
            }
            elseif ($e.KeyCode -eq "Escape") {
                $dialog.Close()
                $e.Handled = $true
            }
        })

        [void]$dialog.ShowDialog($ParentForm)

        return $dialog.Tag
    }
    catch {
        Write-Error "Details dialog error: $_"
        return $null
    }
}


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 7: UTILITY LAUNCHER (WITH REAL-TIME STATUS)
# ══════════════════════════════════════════════════════════════════════════════
# Core function to download and execute utilities with live progress updates
#
# Button States:
#   1. "Downloading..." (Yellow)  - Fetching script from URL
#   2. "Starting..."    (Orange)  - Creating background job
#   3. "Launching..."   (Blue)    - Waiting for job to initialize
#   4. "Running"        (Green)   - Job is actively running
#   5. "Done ✓"         (Green)   - Job completed successfully
#   6. "Run"            (Primary) - Ready for next launch

function Invoke-Utility {
    param(
        $UtilData,
        $SenderButton
    )

    # --- Color Definitions for Button States ---
    $ColorDownloading = [System.Drawing.Color]::FromArgb(255, 204, 0)    # Yellow
    $ColorStarting    = [System.Drawing.Color]::FromArgb(255, 149, 0)    # Orange
    $ColorLaunching   = [System.Drawing.Color]::FromArgb(0, 122, 255)    # Blue
    $ColorRunning     = [System.Drawing.Color]::FromArgb(52, 199, 89)    # Green
    $ColorSuccess     = [System.Drawing.Color]::FromArgb(52, 199, 89)    # Green
    $ColorError       = [System.Drawing.Color]::FromArgb(255, 69, 58)    # Red
    $ColorText        = [System.Drawing.Color]::FromArgb(0, 0, 0)        # Black (for yellow bg)
    $ColorWhite       = [System.Drawing.Color]::White

    try {
        if ($null -eq $UtilData) { return }

        # ═══════════════════════════════════════════════════════════════════
        # PHASE 1: DOWNLOADING
        # ═══════════════════════════════════════════════════════════════════
        
        Write-Host ""
        Write-Host "=====================================" -ForegroundColor Cyan
        Write-Host "LAUNCHING: $($UtilData.Name)" -ForegroundColor Yellow
        Write-Host "=====================================" -ForegroundColor Cyan
        Write-Host "Source: $($UtilData.Link)" -ForegroundColor Gray
        
        # Update button: Downloading
        Set-ButtonState -Button $SenderButton `
            -Text "Downloading..." `
            -BackColor $ColorDownloading `
            -ForeColor $ColorText `
            -Enabled $false

        Update-StatusBar "Downloading $($UtilData.Name)..."
        Write-Host "[1/5] Downloading script..." -ForegroundColor Yellow

        # Download the script
        $downloadStart = Get-Date
        $scriptContent = Invoke-RestMethod -Uri $UtilData.Link -UseBasicParsing
        $downloadTime = ((Get-Date) - $downloadStart).TotalSeconds

        Write-Host "[1/5] Downloaded ($([Math]::Round($downloadTime, 2))s)" -ForegroundColor Green

        # ═══════════════════════════════════════════════════════════════════
        # PHASE 2: STARTING
        # ═══════════════════════════════════════════════════════════════════
        
        # Update button: Starting
        Set-ButtonState -Button $SenderButton `
            -Text "Starting..." `
            -BackColor $ColorStarting `
            -ForeColor $ColorWhite `
            -Enabled $false

        Update-StatusBar "Starting $($UtilData.Name)..."
        Write-Host "[2/5] Creating background job..." -ForegroundColor Yellow

        # Create the background job
        $job = Start-Job -ScriptBlock {
            param($script)
            Invoke-Expression $script
        } -ArgumentList $scriptContent

        Write-Host "[2/5] Job created (ID: $($job.Id))" -ForegroundColor Green

        # ═══════════════════════════════════════════════════════════════════
        # PHASE 3: LAUNCHING (WAIT FOR JOB TO START)
        # ═══════════════════════════════════════════════════════════════════
        
        # Update button: Launching
        Set-ButtonState -Button $SenderButton `
            -Text "Launching..." `
            -BackColor $ColorLaunching `
            -ForeColor $ColorWhite `
            -Enabled $false

        Update-StatusBar "Launching $($UtilData.Name)..."
        Write-Host "[3/5] Waiting for job to start..." -ForegroundColor Yellow

        # Wait for job to leave "NotStarted" state (max 10 seconds)
        $launchTimeout = 10
        $launchStart = Get-Date
        
        while ($job.State -eq "NotStarted") {
            Start-Sleep -Milliseconds 100
            [System.Windows.Forms.Application]::DoEvents()
            
            $elapsed = ((Get-Date) - $launchStart).TotalSeconds
            if ($elapsed -gt $launchTimeout) {
                Write-Host "[3/5] Launch timeout - proceeding anyway" -ForegroundColor Yellow
                break
            }
        }

        Write-Host "[3/5] Job state: $($job.State)" -ForegroundColor Green

        # ═══════════════════════════════════════════════════════════════════
        # PHASE 4: RUNNING (MONITOR JOB)
        # ═══════════════════════════════════════════════════════════════════
        
        if ($job.State -eq "Running") {
            # Update button: Running
            Set-ButtonState -Button $SenderButton `
                -Text "Running" `
                -BackColor $ColorRunning `
                -ForeColor $ColorWhite `
                -Enabled $false

            Update-StatusBar "Running: $($UtilData.Name) [Job $($job.Id)]"
            Write-Host "[4/5] Job is running..." -ForegroundColor Green

            # Monitor job for a few seconds to catch quick completions/failures
            $monitorTime = 3  # seconds to monitor
            $monitorStart = Get-Date

            while ($job.State -eq "Running") {
                Start-Sleep -Milliseconds 200
                [System.Windows.Forms.Application]::DoEvents()

                $elapsed = ((Get-Date) - $monitorStart).TotalSeconds
                if ($elapsed -gt $monitorTime) {
                    # Job still running after monitor period - it's a long-running utility
                    Write-Host "[4/5] Job continues in background" -ForegroundColor Green
                    break
                }
            }
        }

        # ═══════════════════════════════════════════════════════════════════
        # PHASE 5: COMPLETION
        # ═══════════════════════════════════════════════════════════════════
        
        Write-Host "[5/5] Final state: $($job.State)" -ForegroundColor Cyan

        # Check final job state
        switch ($job.State) {
            "Running" {
                # Still running - show success and reset
                Set-ButtonState -Button $SenderButton `
                    -Text "Done ✓" `
                    -BackColor $ColorSuccess `
                    -ForeColor $ColorWhite `
                    -Enabled $false

                Update-StatusBar "Launched: $($UtilData.Name) [Job $($job.Id) running]"
                Write-Host "Status: LAUNCHED SUCCESSFULLY" -ForegroundColor Green
            }
            "Completed" {
                # Job finished quickly
                Set-ButtonState -Button $SenderButton `
                    -Text "Done ✓" `
                    -BackColor $ColorSuccess `
                    -ForeColor $ColorWhite `
                    -Enabled $false

                Update-StatusBar "Completed: $($UtilData.Name)"
                Write-Host "Status: COMPLETED" -ForegroundColor Green

                # Check for any output
                $output = Receive-Job -Job $job -ErrorAction SilentlyContinue
                if ($output) {
                    Write-Host "Output: $output" -ForegroundColor Gray
                }
            }
            "Failed" {
                # Job failed
                Set-ButtonState -Button $SenderButton `
                    -Text "Failed ✗" `
                    -BackColor $ColorError `
                    -ForeColor $ColorWhite `
                    -Enabled $false

                $errorMsg = $job.ChildJobs[0].JobStateInfo.Reason.Message
                Update-StatusBar "Failed: $($UtilData.Name)"
                Write-Host "Status: FAILED - $errorMsg" -ForegroundColor Red

                # Show error dialog
                Show-CustomDialog `
                    -Title "Launch Failed" `
                    -Message "The utility '$($UtilData.Name)' failed to run:`n`n$errorMsg" `
                    -Buttons @("OK") `
                    -ParentForm $script:mainForm
            }
            default {
                # Unknown state
                Set-ButtonState -Button $SenderButton `
                    -Text "Done" `
                    -BackColor $ColorSuccess `
                    -ForeColor $ColorWhite `
                    -Enabled $false

                Update-StatusBar "Launched: $($UtilData.Name) [$($job.State)]"
                Write-Host "Status: $($job.State)" -ForegroundColor Yellow
            }
        }

        Write-Host "=====================================" -ForegroundColor Cyan
        Write-Host ""

        # ═══════════════════════════════════════════════════════════════════
        # PHASE 6: RESET BUTTON (AFTER DELAY)
        # ═══════════════════════════════════════════════════════════════════
        
        # Keep "Done ✓" visible for 2 seconds, then reset
        Start-Sleep -Milliseconds 500
        [System.Windows.Forms.Application]::DoEvents()
        
        Start-Sleep -Milliseconds 500
        [System.Windows.Forms.Application]::DoEvents()
        
        Start-Sleep -Milliseconds 500
        [System.Windows.Forms.Application]::DoEvents()
        
        Start-Sleep -Milliseconds 500
        [System.Windows.Forms.Application]::DoEvents()

        # Reset button to default state
        Set-ButtonState -Button $SenderButton `
            -Text "Run" `
            -BackColor $script:Colors.Primary `
            -ForeColor $ColorWhite `
            -Enabled $true

    }
    catch {
        # ═══════════════════════════════════════════════════════════════════
        # ERROR HANDLING
        # ═══════════════════════════════════════════════════════════════════
        
        Write-Host "Status: FAILED - $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "=====================================" -ForegroundColor Cyan

        Update-StatusBar "Error: $($UtilData.Name)"

        # Show error state
        Set-ButtonState -Button $SenderButton `
            -Text "Error ✗" `
            -BackColor $ColorError `
            -ForeColor $ColorWhite `
            -Enabled $false

        # Show error dialog
        Show-CustomDialog `
            -Title "Launch Error" `
            -Message "Failed to launch '$($UtilData.Name)':`n`n$($_.Exception.Message)" `
            -Buttons @("OK") `
            -ParentForm $script:mainForm

        # Wait then reset
        Start-Sleep -Milliseconds 1500
        [System.Windows.Forms.Application]::DoEvents()

        # Reset button
        Set-ButtonState -Button $SenderButton `
            -Text "Run" `
            -BackColor $script:Colors.Primary `
            -ForeColor $ColorWhite `
            -Enabled $true
    }
}


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 8: THEME FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════
# Toggle and apply themes to all UI elements

function Switch-Theme {
    if ($script:CurrentTheme -eq "Dark") {
        $script:CurrentTheme = "Light"
        $script:Colors = $LightTheme
    }
    else {
        $script:CurrentTheme = "Dark"
        $script:Colors = $DarkTheme
    }
    Set-Theme
}

function Set-Theme {
    try {
        $script:mainForm.BackColor = $script:Colors.Background
        $script:headerPanel.BackColor = $script:Colors.Surface
        $script:titleLabel.ForeColor = $script:Colors.Text
        $script:subtitleLabel.ForeColor = $script:Colors.TextSecondary
        $script:searchContainer.BackColor = $script:Colors.SearchBox
        $script:searchBox.BackColor = $script:Colors.SearchBox
        $script:searchBox.ForeColor = $script:Colors.Text
        $script:clearBtn.BackColor = $script:Colors.SurfaceLight
        $script:clearBtn.ForeColor = $script:Colors.TextSecondary
        $script:refreshBtn.BackColor = $script:Colors.SurfaceLight
        $script:refreshBtn.ForeColor = $script:Colors.Primary
        $script:themeBtn.BackColor = $script:Colors.SurfaceLight
        $script:themeBtn.ForeColor = $script:Colors.Text
        $script:utilitiesPanel.BackColor = $script:Colors.Background
        $script:statusBar.BackColor = $script:Colors.Surface
        $script:statusLabel.ForeColor = $script:Colors.Text
        $script:versionLabel.ForeColor = $script:Colors.Primary

        Update-UtilityDisplay
    }
    catch {
        Write-Warning "Theme apply error: $_"
    }
}


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 9: UI CARD BUILDER
# ══════════════════════════════════════════════════════════════════════════════
# Creates individual utility cards for the main panel

function New-UtilityCard {
    param(
        $Util,
        [int]$ContainerWidth
    )

    try {
        if ($null -eq $Util) { return $null }

        # --- Card Dimensions ---
        $cardWidth = [Math]::Max(400, $ContainerWidth - 60)
        $cardHeight = 90

        # --- Card Container ---
        $card = New-Object System.Windows.Forms.Panel
        $card.Size = New-Object System.Drawing.Size($cardWidth, $cardHeight)
        $card.BackColor = $script:Colors.Surface
        $card.Tag = $Util
        $card.Cursor = [System.Windows.Forms.Cursors]::Hand
        $card.Anchor = "Top,Left,Right"
        $card.BorderStyle = "FixedSingle"

        # --- Title Label ---
        $titleLbl = New-Object System.Windows.Forms.Label
        $titleLbl.Text = $Util.Name
        $titleLbl.Location = New-Object System.Drawing.Point(20, 15)
        $titleLbl.Size = New-Object System.Drawing.Size(($cardWidth - 180), 24)
        $titleLbl.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
        $titleLbl.ForeColor = $script:Colors.Text
        $titleLbl.BackColor = [System.Drawing.Color]::Transparent
        $titleLbl.Anchor = "Top,Left,Right"
        Add-Tooltip -Control $titleLbl -Text "Click to view details"

        # --- Description Label ---
        $descLbl = New-Object System.Windows.Forms.Label
        $descLbl.Text = $Util.Desc
        $descLbl.Location = New-Object System.Drawing.Point(20, 42)
        $descLbl.Size = New-Object System.Drawing.Size(($cardWidth - 180), 18)
        $descLbl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $descLbl.ForeColor = $script:Colors.TextSecondary
        $descLbl.BackColor = [System.Drawing.Color]::Transparent
        $descLbl.Anchor = "Top,Left,Right"

        # --- Category Label ---
        $catLbl = New-Object System.Windows.Forms.Label
        $catLbl.Text = $Util.Category
        $catLbl.Location = New-Object System.Drawing.Point(20, 63)
        $catLbl.AutoSize = $true
        $catLbl.Font = New-Object System.Drawing.Font("Segoe UI", 8)
        $catLbl.ForeColor = $script:Colors.Primary
        $catLbl.BackColor = [System.Drawing.Color]::Transparent

        # --- Run Button ---
        $runBtn = New-Object System.Windows.Forms.Button
        $runBtn.Text = "Run"
        $runBtn.Size = New-Object System.Drawing.Size(100, 35)
        $runBtn.Location = New-Object System.Drawing.Point(($cardWidth - 120), 28)
        $runBtn.BackColor = $script:Colors.Primary
        $runBtn.ForeColor = [System.Drawing.Color]::White
        $runBtn.FlatStyle = "Flat"
        $runBtn.FlatAppearance.BorderSize = 0
        $runBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $runBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
        $runBtn.Anchor = "Top,Right"
        Add-Tooltip -Control $runBtn -Text "Launch $($Util.Name)"

        # --- Run Button Click ---
        $runBtn.Add_Click({
            param($sender, $e)
            try {
                # Don't process if button is already in a loading state
                if ($sender.Text -ne "Run") { return }

                $utilData = $sender.Parent.Tag
                if ($null -eq $utilData) { return }

                $confirm = Show-CustomDialog `
                    -Title "Launch Utility" `
                    -Message "Launch '$($utilData.Name)'?`n`n$($utilData.Desc)`n`nSource: $($utilData.Link)" `
                    -Buttons @("Launch", "Cancel") `
                    -DefaultButton "Launch" `
                    -ParentForm $script:mainForm

                if ($confirm -eq "Launch") {
                    Invoke-Utility -UtilData $utilData -SenderButton $sender
                }
            }
            catch {
                Write-Error "Run button error: $_"
            }
        })

        # --- Card Click (Show Details) ---
        $card.Add_Click({
            try {
                $utilData = $this.Tag
                if ($null -eq $utilData) { return }

                $result = Show-UtilityDetailsDialog -Util $utilData -ParentForm $script:mainForm

                if ($result -eq "Run") {
                    foreach ($ctrl in $this.Controls) {
                        if ($ctrl -is [System.Windows.Forms.Button] -and $ctrl.Text -eq "Run") {
                            $ctrl.PerformClick()
                            break
                        }
                    }
                }
            }
            catch {
                Write-Error "Card click error: $_"
            }
        })

        # --- Hover Effects ---
        $card.Add_MouseEnter({ $this.BackColor = $script:Colors.SurfaceHover })
        $card.Add_MouseLeave({ $this.BackColor = $script:Colors.Surface })

        $runBtn.Add_MouseEnter({
            # Only show hover effect if button is in default "Run" state
            if ($this.Enabled -and $this.Text -eq "Run") {
                $this.BackColor = $script:Colors.PrimaryHover
            }
        })
        $runBtn.Add_MouseLeave({
            # Only reset if button is in default "Run" state
            if ($this.Enabled -and $this.Text -eq "Run") {
                $this.BackColor = $script:Colors.Primary
            }
        })

        # --- Add Controls to Card ---
        $card.Controls.AddRange(@($titleLbl, $descLbl, $catLbl, $runBtn))

        return $card
    }
    catch {
        Write-Error "Card creation error: $_"
        return $null
    }
}


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 10: MAIN FORM SETUP
# ══════════════════════════════════════════════════════════════════════════════
# Create and configure the main application window

$script:mainForm = New-Object System.Windows.Forms.Form
$script:mainForm.Text = "Nahid Power Utilities"
$script:mainForm.StartPosition = "CenterScreen"
$script:mainForm.WindowState = "Maximized"
$script:mainForm.BackColor = $script:Colors.Background
$script:mainForm.FormBorderStyle = "Sizable"
$script:mainForm.MinimumSize = New-Object System.Drawing.Size(700, 550)
$script:mainForm.KeyPreview = $true


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 11: HEADER SECTION
# ══════════════════════════════════════════════════════════════════════════════
# Top section with title, search bar, and control buttons

# --- Header Panel ---
$script:headerPanel = New-Object System.Windows.Forms.Panel
$script:headerPanel.Dock = "Top"
$script:headerPanel.Height = 130
$script:headerPanel.BackColor = $script:Colors.Surface

# --- Title ---
$script:titleLabel = New-Object System.Windows.Forms.Label
$script:titleLabel.Text = "Nahid Power Utilities"
$script:titleLabel.Location = New-Object System.Drawing.Point(30, 20)
$script:titleLabel.AutoSize = $true
$script:titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 20, [System.Drawing.FontStyle]::Bold)
$script:titleLabel.ForeColor = $script:Colors.Text
Add-Tooltip -Control $script:titleLabel -Text "Dashboard v2.2"

# --- Subtitle ---
$script:subtitleLabel = New-Object System.Windows.Forms.Label
$script:subtitleLabel.Text = "Your productivity toolkit"
$script:subtitleLabel.Location = New-Object System.Drawing.Point(30, 55)
$script:subtitleLabel.AutoSize = $true
$script:subtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$script:subtitleLabel.ForeColor = $script:Colors.TextSecondary

# --- Search Container ---
$script:searchContainer = New-Object System.Windows.Forms.Panel
$script:searchContainer.Location = New-Object System.Drawing.Point(25, 85)
$script:searchContainer.Size = New-Object System.Drawing.Size(400, 35)
$script:searchContainer.BackColor = $script:Colors.SearchBox
$script:searchContainer.Anchor = "Top,Left,Right"
$script:searchContainer.BorderStyle = "FixedSingle"

# --- Search Icon ---
$searchIcon = New-Object System.Windows.Forms.Label
$searchIcon.Text = "Search"
$searchIcon.Location = New-Object System.Drawing.Point(10, 8)
$searchIcon.Size = New-Object System.Drawing.Size(45, 20)
$searchIcon.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$searchIcon.ForeColor = $script:Colors.TextSecondary

# --- Search TextBox ---
$script:searchBox = New-Object System.Windows.Forms.TextBox
$script:searchBox.Location = New-Object System.Drawing.Point(60, 7)
$script:searchBox.Size = New-Object System.Drawing.Size(330, 22)
$script:searchBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$script:searchBox.BackColor = $script:Colors.SearchBox
$script:searchBox.ForeColor = $script:Colors.Text
$script:searchBox.BorderStyle = "None"
$script:searchBox.Anchor = "Top,Left,Right"
Add-Tooltip -Control $script:searchBox -Text "Search utilities (Press / to focus)"

$script:searchContainer.Controls.AddRange(@($searchIcon, $script:searchBox))

# --- Theme Button ---
$script:themeBtn = New-Object System.Windows.Forms.Button
$script:themeBtn.Text = "Theme"
$script:themeBtn.Size = New-Object System.Drawing.Size(60, 35)
$script:themeBtn.Location = New-Object System.Drawing.Point(440, 85)
$script:themeBtn.BackColor = $script:Colors.SurfaceLight
$script:themeBtn.ForeColor = $script:Colors.Text
$script:themeBtn.FlatStyle = "Flat"
$script:themeBtn.FlatAppearance.BorderSize = 0
$script:themeBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$script:themeBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
$script:themeBtn.Anchor = "Top,Right"
Add-Tooltip -Control $script:themeBtn -Text "Toggle dark/light mode (Ctrl+T)"

$script:themeBtn.Add_Click({ Switch-Theme })
$script:themeBtn.Add_MouseEnter({ $this.BackColor = $script:Colors.SurfaceHover })
$script:themeBtn.Add_MouseLeave({ $this.BackColor = $script:Colors.SurfaceLight })

# --- Clear Button ---
$script:clearBtn = New-Object System.Windows.Forms.Button
$script:clearBtn.Text = "Clear"
$script:clearBtn.Size = New-Object System.Drawing.Size(60, 35)
$script:clearBtn.Location = New-Object System.Drawing.Point(510, 85)
$script:clearBtn.BackColor = $script:Colors.SurfaceLight
$script:clearBtn.ForeColor = $script:Colors.TextSecondary
$script:clearBtn.FlatStyle = "Flat"
$script:clearBtn.FlatAppearance.BorderSize = 0
$script:clearBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$script:clearBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
$script:clearBtn.Anchor = "Top,Right"
Add-Tooltip -Control $script:clearBtn -Text "Clear search (Esc)"

$script:clearBtn.Add_Click({
    $script:searchBox.Text = ""
    $script:searchBox.Focus()
    Update-UtilityDisplay
})
$script:clearBtn.Add_MouseEnter({ $this.BackColor = $script:Colors.Danger; $this.ForeColor = [System.Drawing.Color]::White })
$script:clearBtn.Add_MouseLeave({ $this.BackColor = $script:Colors.SurfaceLight; $this.ForeColor = $script:Colors.TextSecondary })

# --- Refresh Button ---
$script:refreshBtn = New-Object System.Windows.Forms.Button
$script:refreshBtn.Text = "Refresh"
$script:refreshBtn.Size = New-Object System.Drawing.Size(65, 35)
$script:refreshBtn.Location = New-Object System.Drawing.Point(580, 85)
$script:refreshBtn.BackColor = $script:Colors.SurfaceLight
$script:refreshBtn.ForeColor = $script:Colors.Primary
$script:refreshBtn.FlatStyle = "Flat"
$script:refreshBtn.FlatAppearance.BorderSize = 0
$script:refreshBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$script:refreshBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
$script:refreshBtn.Anchor = "Top,Right"
Add-Tooltip -Control $script:refreshBtn -Text "Refresh (F5)"

$script:refreshBtn.Add_Click({
    Update-UtilityDisplay
    Update-StatusBar "Refreshed"
})
$script:refreshBtn.Add_MouseEnter({ $this.BackColor = $script:Colors.Primary; $this.ForeColor = [System.Drawing.Color]::White })
$script:refreshBtn.Add_MouseLeave({ $this.BackColor = $script:Colors.SurfaceLight; $this.ForeColor = $script:Colors.Primary })

# --- Add Controls to Header ---
$script:headerPanel.Controls.AddRange(@(
    $script:titleLabel,
    $script:subtitleLabel,
    $script:searchContainer,
    $script:themeBtn,
    $script:clearBtn,
    $script:refreshBtn
))


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 12: UTILITIES PANEL
# ══════════════════════════════════════════════════════════════════════════════
# Scrollable panel that contains all utility cards

$script:utilitiesPanel = New-Object System.Windows.Forms.Panel
$script:utilitiesPanel.Dock = "Fill"
$script:utilitiesPanel.AutoScroll = $true
$script:utilitiesPanel.BackColor = $script:Colors.Background
$script:utilitiesPanel.Padding = New-Object System.Windows.Forms.Padding(25, 25, 25, 100)


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 13: STATUS BAR
# ══════════════════════════════════════════════════════════════════════════════
# Bottom status bar showing current state and version

$script:statusBar = New-Object System.Windows.Forms.StatusStrip
$script:statusBar.BackColor = $script:Colors.Surface
$script:statusBar.Dock = "Bottom"

$script:statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:statusLabel.Text = "Ready | Total: $($script:ValidatedUtilities.Count) utilities"
$script:statusLabel.ForeColor = $script:Colors.Text
$script:statusLabel.Spring = $true
$script:statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$script:versionLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:versionLabel.Text = "v2.2"
$script:versionLabel.ForeColor = $script:Colors.Primary
$script:versionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$script:statusBar.Items.AddRange(@($script:statusLabel, $script:versionLabel)) | Out-Null


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 14: DISPLAY UPDATE FUNCTION
# ══════════════════════════════════════════════════════════════════════════════
# Rebuilds the utility cards based on current search filter

function Update-UtilityDisplay {
    param([string]$SearchText = "")

    try {
        $script:utilitiesPanel.SuspendLayout()
        $script:utilitiesPanel.Controls.Clear()

        $filtered = $script:ValidatedUtilities

        if ($SearchText -eq "") {
            $SearchText = $script:searchBox.Text
        }

        if ($SearchText -ne "") {
            $filtered = $filtered | Where-Object {
                $_.Name -like "*$SearchText*" -or
                $_.Desc -like "*$SearchText*" -or
                $_.Category -like "*$SearchText*" -or
                ($_.Tags | Where-Object { $_ -like "*$SearchText*" }).Count -gt 0
            }
        }

        # --- Count Header ---
        $countPanel = New-Object System.Windows.Forms.Panel
        $countPanel.Height = 40
        $countPanel.Width = $script:utilitiesPanel.Width - 50
        $countPanel.BackColor = [System.Drawing.Color]::Transparent
        $countPanel.Anchor = "Top,Left,Right"

        $countLabel = New-Object System.Windows.Forms.Label
        $countLabel.Text = "Showing $($filtered.Count) of $($script:ValidatedUtilities.Count) utilities"
        $countLabel.Location = New-Object System.Drawing.Point(10, 10)
        $countLabel.AutoSize = $true
        $countLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $countLabel.ForeColor = $script:Colors.Text

        $countPanel.Controls.Add($countLabel)
        $script:utilitiesPanel.Controls.Add($countPanel)

        # --- Build Cards ---
        $yPos = 50
        $cardSpacing = 15

        foreach ($util in $filtered) {
            $card = New-UtilityCard -Util $util -ContainerWidth $script:utilitiesPanel.ClientSize.Width
            if ($null -ne $card) {
                $card.Location = New-Object System.Drawing.Point(15, $yPos)
                $script:utilitiesPanel.Controls.Add($card)
                $yPos += $card.Height + $cardSpacing
            }
        }

        # --- Bottom Spacer ---
        $bottomSpacer = New-Object System.Windows.Forms.Panel
        $bottomSpacer.Size = New-Object System.Drawing.Size(10, 120)
        $bottomSpacer.Location = New-Object System.Drawing.Point(0, $yPos)
        $bottomSpacer.BackColor = [System.Drawing.Color]::Transparent
        $script:utilitiesPanel.Controls.Add($bottomSpacer)

        # --- No Results ---
        if ($filtered.Count -eq 0) {
            $noResultsPanel = New-Object System.Windows.Forms.Panel
            $noResultsPanel.Size = New-Object System.Drawing.Size(($script:utilitiesPanel.Width - 50), 180)
            $noResultsPanel.Location = New-Object System.Drawing.Point(25, 60)
            $noResultsPanel.BackColor = $script:Colors.Surface
            $noResultsPanel.Anchor = "Top,Left,Right"

            $noResultsLabel = New-Object System.Windows.Forms.Label
            $noResultsLabel.Text = "No utilities found"
            $noResultsLabel.Size = New-Object System.Drawing.Size($noResultsPanel.Width, 35)
            $noResultsLabel.Location = New-Object System.Drawing.Point(0, 60)
            $noResultsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
            $noResultsLabel.ForeColor = $script:Colors.Text
            $noResultsLabel.TextAlign = "MiddleCenter"

            $noResultsDesc = New-Object System.Windows.Forms.Label
            $noResultsDesc.Text = "Try different search terms"
            $noResultsDesc.Size = New-Object System.Drawing.Size($noResultsPanel.Width, 25)
            $noResultsDesc.Location = New-Object System.Drawing.Point(0, 100)
            $noResultsDesc.Font = New-Object System.Drawing.Font("Segoe UI", 10)
            $noResultsDesc.ForeColor = $script:Colors.TextSecondary
            $noResultsDesc.TextAlign = "MiddleCenter"

            $noResultsPanel.Controls.AddRange(@($noResultsLabel, $noResultsDesc))
            $script:utilitiesPanel.Controls.Add($noResultsPanel)
        }

        Update-StatusBar "Showing $($filtered.Count) of $($script:ValidatedUtilities.Count) utilities"
    }
    catch {
        Write-Error "Display update error: $_"
    }
    finally {
        $script:utilitiesPanel.ResumeLayout($true)
    }
}


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 15: EVENT HANDLERS
# ══════════════════════════════════════════════════════════════════════════════
# Handle form resize, search, keyboard shortcuts, and cleanup

# --- Form Resize ---
$script:mainForm.Add_Resize({
    try {
        $rightEdge = $script:headerPanel.ClientSize.Width
        
        $script:refreshBtn.Left = $rightEdge - 75
        $script:clearBtn.Left = $rightEdge - 145
        $script:themeBtn.Left = $rightEdge - 215
        
        $script:searchContainer.Width = $script:themeBtn.Left - 50
        $script:searchBox.Width = $script:searchContainer.Width - 70

        Update-UtilityDisplay
    }
    catch {
        Write-Warning "Resize error: $_"
    }
})

# --- Search Text Changed ---
$script:searchBox.Add_TextChanged({
    try {
        Update-UtilityDisplay
    }
    catch {
        Write-Warning "Search error: $_"
    }
})

# --- Global Keyboard Shortcuts ---
$script:mainForm.Add_KeyDown({
    param($sender, $e)

    if ($e.KeyCode -eq "OemQuestion" -or $e.KeyCode -eq "Divide" -or
        ($e.Control -and $e.KeyCode -eq "F")) {
        $script:searchBox.Focus()
        $script:searchBox.SelectAll()
        $e.Handled = $true
    }

    if ($e.KeyCode -eq "Escape") {
        if ($script:searchBox.Text -ne "") {
            $script:searchBox.Text = ""
        }
        $e.Handled = $true
    }

    if ($e.KeyCode -eq "F5") {
        Update-UtilityDisplay
        Update-StatusBar "Refreshed"
        $e.Handled = $true
    }

    if ($e.Control -and $e.KeyCode -eq "T") {
        Switch-Theme
        $e.Handled = $true
    }
})

# --- Form Closing (Cleanup) ---
$script:mainForm.Add_FormClosing({
    try {
        Get-Job | Remove-Job -Force -ErrorAction SilentlyContinue
        
        if ($null -ne $script:tooltip) {
            $script:tooltip.Dispose()
        }
    }
    catch {
        Write-Warning "Cleanup error: $_"
    }
})


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 16: APPLICATION LAUNCH
# ══════════════════════════════════════════════════════════════════════════════
# Assemble the form and show it to the user

$script:mainForm.Controls.AddRange(@(
    $script:statusBar,
    $script:utilitiesPanel,
    $script:headerPanel
))

Update-UtilityDisplay

# --- Console Welcome Message ---
Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "   Nahid Power Utilities Dashboard v2.2        " -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "Loaded  : $($script:ValidatedUtilities.Count) utilities" -ForegroundColor Green
Write-Host ""
Write-Host "Button States:" -ForegroundColor Yellow
Write-Host "  Downloading... : Fetching script" -ForegroundColor DarkYellow
Write-Host "  Starting...    : Creating job" -ForegroundColor DarkYellow
Write-Host "  Launching...   : Initializing" -ForegroundColor Blue
Write-Host "  Running        : Actively running" -ForegroundColor Green
Write-Host "  Done           : Completed" -ForegroundColor Green
Write-Host ""
Write-Host "Shortcuts:" -ForegroundColor Yellow
Write-Host "  / or Ctrl+F : Focus search" -ForegroundColor Gray
Write-Host "  Esc         : Clear search" -ForegroundColor Gray
Write-Host "  F5          : Refresh" -ForegroundColor Gray
Write-Host "  Ctrl+T      : Toggle theme" -ForegroundColor Gray
Write-Host ""

try {
    [void]$script:mainForm.ShowDialog()
}
catch {
    Write-Error "Application error: $_"
}
