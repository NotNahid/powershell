#requires -Version 5.1
using namespace System.Windows.Forms
using namespace System.Drawing
using namespace System.Runtime.InteropServices

# --- 1. Define Native API and the Main Form Class ---
Add-Type -TypeDefinition @"
using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;

public class AppForm : Form {
    [DllImport("user32.dll")]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);
    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")]
    public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll")]
    public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    public const uint WM_HOTKEY = 0x0312;
    public const uint MOUSEEVENTF_LEFTDOWN = 0x02;
    public const uint MOUSEEVENTF_LEFTUP = 0x04;
    public const uint MOUSEEVENTF_RIGHTDOWN = 0x08;
    public const uint MOUSEEVENTF_RIGHTUP = 0x10;
    public const int HOTKEY_START = 9000;
    public const int HOTKEY_STOP = 9001;

    public Action OnStart;
    public Action OnStop;

    public AppForm() {
        this.FormClosing += (s, e) => {
            UnregisterHotKey(this.Handle, HOTKEY_START);
            UnregisterHotKey(this.Handle, HOTKEY_STOP);
        };
    }

    protected override void OnLoad(EventArgs e) {
        base.OnLoad(e);
        RegisterHotKey(this.Handle, HOTKEY_START, 0, (uint)Keys.F6);
        RegisterHotKey(this.Handle, HOTKEY_STOP, 0, (uint)Keys.F7);
    }

    protected override void WndProc(ref Message m) {
        base.WndProc(ref m);
        if (m.Msg == WM_HOTKEY) {
            int id = m.WParam.ToInt32();
            if (id == HOTKEY_START && OnStart != null) { OnStart.Invoke(); }
            if (id == HOTKEY_STOP && OnStop != null) { OnStop.Invoke(); }
        }
    }
}
"@ -ReferencedAssemblies System.Windows.Forms, System.Drawing

# --- 2. Script State ---
$syncHash = [hashtable]::Synchronized(@{
    Running      = $false
    Stop         = $false
    Interval     = 100
    ClickType    = "Left"
    PosMode      = "Current"
    FixedX       = 0
    FixedY       = 0
    Jitter       = $false
    JitterAmount = 15
    ClickCount   = 0
    PSInstance   = $null
    Runspace     = $null
    AsyncResult  = $null
})

# --- 3. Aesthetic Definitions (Perfectly Balanced) ---
$ColorBg      = [ColorTranslator]::FromHtml("#F8F8F8")
$ColorSurface = [Color]::White
$ColorText    = [ColorTranslator]::FromHtml("#1A1A1A")
$ColorMuted   = [ColorTranslator]::FromHtml("#858585")
$ColorAccent  = [ColorTranslator]::FromHtml("#0078D4")
$ColorDanger  = [ColorTranslator]::FromHtml("#C42B1C")
$ColorBorder  = [ColorTranslator]::FromHtml("#CCCCCC") # Slightly more visible border for inputs/buttons
$FontUI       = [Font]::new("Segoe UI", 10)
$FontTitle    = [Font]::new("Segoe UI", 16, [FontStyle]::Bold)
$FontSection  = [Font]::new("Segoe UI", 10, [FontStyle]::Bold)
$FontStatus   = [Font]::new("Segoe UI", 10)

# --- 4. Build the UI ---
$form = [AppForm]::new()
$form.BackColor = $ColorBg
$form.Text = "Autoclicker"
$form.Size = [Size]::new(360, 430) 
$form.FormBorderStyle = 'FixedSingle' # Forces a clean, even border all the way around
$form.MaximizeBox = $false
$form.StartPosition = 'CenterScreen'
$form.Font = $FontUI

$marginLeft = 30

# Title
$lblTitle = [Label]@{
    Location = [Point]::new($marginLeft, 20)
    Size = [Size]::new(200, 35)
    Text = "Autoclicker"
    ForeColor = $ColorText
    Font = $FontTitle
}

# --- TIMING SECTION ---
$yPos = 65

$lblTimingSec = [Label]@{Location=[Point]::new($marginLeft, $yPos); Size=[Size]::new(100, 20); Text="TIMING"; ForeColor=$ColorMuted; Font=$FontSection}
$yPos += 28

$lblInt = [Label]@{Location=[Point]::new($marginLeft, $yPos); Size=[Size]::new(65, 25); Text="Interval:"; ForeColor=$ColorText}
$numInt = [NumericUpDown]@{Location=[Point]::new(95, $yPos-2); Size=[Size]::new(75, 26); BackColor=$ColorSurface; ForeColor=$ColorText; BorderStyle='FixedSingle'; Minimum=1; Maximum=9999; Value=100; TextAlign='Center'}
$lblMs  = [Label]@{Location=[Point]::new(175, $yPos); Size=[Size]::new(25, 25); Text="ms"; ForeColor=$ColorMuted}

$chkJitter = [CheckBox]@{Location=[Point]::new(205, $yPos); Size=[Size]::new(65, 25); Text="Jitter"; ForeColor=$ColorText; BackColor=$ColorBg}
$numJitter = [NumericUpDown]@{Location=[Point]::new(275, $yPos-2); Size=[Size]::new(45, 26); BackColor=$ColorSurface; ForeColor=$ColorMuted; BorderStyle='FixedSingle'; Minimum=1; Maximum=500; Value=15; Enabled=$false; TextAlign='Center'}

$numInt.Add_ValueChanged({ $syncHash.Interval = $numInt.Value })
$chkJitter.Add_CheckedChanged({ 
    $syncHash.Jitter = $chkJitter.Checked
    $numJitter.Enabled = $chkJitter.Checked
    if($chkJitter.Checked){ $numJitter.BackColor = $ColorAccent; $numJitter.ForeColor = [Color]::White }
    else { $numJitter.BackColor = $ColorSurface; $numJitter.ForeColor = $ColorMuted }
})

# --- TARGET SECTION ---
$yPos += 45

$lblTargetSec = [Label]@{Location=[Point]::new($marginLeft, $yPos); Size=[Size]::new(100, 20); Text="TARGET"; ForeColor=$ColorMuted; Font=$FontSection}
$yPos += 28

$rbCurr = [RadioButton]@{Location=[Point]::new($marginLeft, $yPos); Size=[Size]::new(75, 25); Text="Current"; ForeColor=$ColorText; BackColor=$ColorBg; Checked=$true}
$rbFix = [RadioButton]@{Location=[Point]::new(110, $yPos); Size=[Size]::new(60, 25); Text="Fixed"; ForeColor=$ColorText; BackColor=$ColorBg}

$txtX = [TextBox]@{Location=[Point]::new(240, $yPos-2); Size=[Size]::new(40, 26); Text="0"; BackColor=$ColorSurface; ForeColor=$ColorMuted; Enabled=$false; TextAlign='Center'; BorderStyle='FixedSingle'}
$txtY = [TextBox]@{Location=[Point]::new(285, $yPos-2); Size=[Size]::new(40, 26); Text="0"; BackColor=$ColorSurface; ForeColor=$ColorMuted; Enabled=$false; TextAlign='Center'; BorderStyle='FixedSingle'}

$rbCurr.Add_CheckedChanged({ 
    if($rbCurr.Checked) { 
        $syncHash.PosMode = "Current"; 
        $txtX.Enabled = $false; $txtX.BackColor = $ColorSurface; $txtX.ForeColor = $ColorMuted
        $txtY.Enabled = $false; $txtY.BackColor = $ColorSurface; $txtY.ForeColor = $ColorMuted
    }
})
$rbFix.Add_CheckedChanged({ 
    if($rbFix.Checked) { 
        $syncHash.PosMode = "Fixed"; 
        $txtX.Enabled = $true; $txtX.BackColor = [Color]::White; $txtX.ForeColor = $ColorText
        $txtY.Enabled = $true; $txtY.BackColor = [Color]::White; $txtY.ForeColor = $ColorText
    }
})

# --- OPTIONS SECTION ---
$yPos += 45

$lblOptionsSec = [Label]@{Location=[Point]::new($marginLeft, $yPos); Size=[Size]::new(100, 20); Text="OPTIONS"; ForeColor=$ColorMuted; Font=$FontSection}
$yPos += 28

$rbLeft = [RadioButton]@{Location=[Point]::new($marginLeft, $yPos); Size=[Size]::new(60, 25); Text="Left"; ForeColor=$ColorText; BackColor=$ColorBg; Checked=$true}
$rbRight = [RadioButton]@{Location=[Point]::new(95, $yPos); Size=[Size]::new(65, 25); Text="Right"; ForeColor=$ColorText; BackColor=$ColorBg}
$chkTop = [CheckBox]@{Location=[Point]::new(230, $yPos); Size=[Size]::new(80, 25); Text="On Top"; ForeColor=$ColorText; BackColor=$ColorBg}

$rbLeft.Add_CheckedChanged({ if($rbLeft.Checked){ $syncHash.ClickType = "Left" } })
$rbRight.Add_CheckedChanged({ if($rbRight.Checked){ $syncHash.ClickType = "Right" } })
$chkTop.Add_CheckedChanged({ $form.TopMost = $chkTop.Checked })

# --- STATUS & COUNTER ---
$yPos += 50

$lblStatus = [Label]@{
    Location = [Point]::new($marginLeft, $yPos)
    Size = [Size]::new(150, 25)
    Text = "Status: Idle"
    ForeColor = $ColorMuted
    Font = $FontStatus
}
$lblCounter = [Label]@{
    Location = [Point]::new(200, $yPos)
    Size = [Size]::new(130, 25)
    Text = "Clicks: 0"
    ForeColor = $ColorMuted
    Font = $FontStatus
    TextAlign = 'MiddleRight'
}

$yPos += 45

# --- ACTION BUTTONS ---
$btnStart = [Button]@{
    Location = [Point]::new($marginLeft, $yPos)
    Size = [Size]::new(145, 45)
    Text = "Start (F6)"
    BackColor = $ColorAccent
    ForeColor = [Color]::White
    FlatStyle = 'Flat'
    Font = [Font]::new("Segoe UI", 11, [FontStyle]::Bold)
    Cursor = [Cursors]::Hand
}
$btnStart.FlatAppearance.BorderSize = 0

$btnStop = [Button]@{
    Location = [Point]::new(180, $yPos)
    Size = [Size]::new(145, 45)
    Text = "Stop (F7)"
    BackColor = $ColorSurface
    ForeColor = $ColorText
    FlatStyle = 'Flat'
    Font = [Font]::new("Segoe UI", 11)
    Cursor = [Cursors]::Hand
    Enabled = $false
}
$btnStop.FlatAppearance.BorderSize = 1
$btnStop.FlatAppearance.BorderColor = $ColorBorder

$form.Controls.AddRange(@($lblTitle, $lblTimingSec, $lblInt, $numInt, $lblMs, $chkJitter, $numJitter, $lblTargetSec, $rbCurr, $rbFix, $txtX, $txtY, $lblOptionsSec, $rbLeft, $rbRight, $chkTop, $lblStatus, $lblCounter, $btnStart, $btnStop))

# Timer to safely update UI from Runspace
$uiTimer = [Timer]::new()
$uiTimer.Interval = 150
$uiTimer.Add_Tick({
    if ($syncHash.Running) {
        $lblCounter.Text = "Clicks: $($syncHash.ClickCount)"
    }
})
$uiTimer.Start()

# --- 5. Logic ---
$allConfigControls = @($numInt, $chkJitter, $numJitter, $rbCurr, $rbFix, $txtX, $txtY, $rbLeft, $rbRight, $chkTop)

function Toggle-Config($enabled) {
    foreach($ctrl in $allConfigControls) {
        if($ctrl -is [RadioButton] -or $ctrl -is [CheckBox]) { $ctrl.Enabled = $enabled; continue }
        if($ctrl -is [TextBox] -and $ctrl.Name -in @('txtX','txtY') -and $syncHash.PosMode -eq 'Current') { continue }
        if($ctrl -is [NumericUpDown] -and $ctrl.Name -eq 'numJitter' -and -not $syncHash.Jitter) { continue }
        $ctrl.Enabled = $enabled
    }
}

function Start-Clicker {
    if ($syncHash.Running) { return }

    if ($syncHash.PosMode -eq "Fixed") {
        try {
            $syncHash.FixedX = [int]$txtX.Text
            $syncHash.FixedY = [int]$txtY.Text
        } catch {
            [MessageBox]::Show("Invalid X/Y coordinates.", "Error")
            return
        }
    }
    
    $syncHash.JitterAmount = $numJitter.Value
    $syncHash.ClickCount = 0
    $syncHash.Running = $true
    $syncHash.Stop = $false

    # Update UI State
    Toggle-Config $false
    $btnStart.Enabled = $false
    $btnStart.BackColor = $ColorSurface
    $btnStart.ForeColor = $ColorMuted
    
    $btnStop.Enabled = $true
    $btnStop.BackColor = $ColorDanger
    $btnStop.ForeColor = [Color]::White
    
    $lblStatus.Text = "Status: Running"
    $lblStatus.ForeColor = $ColorAccent
    
    # Create and start Runspace
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("syncHash", $syncHash)
    
    $psCmd = [powershell]::Create()
    $psCmd.Runspace = $runspace
    $psCmd.AddScript({
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $nextTick = 0
        
        while (-not $syncHash.Stop) {
            if ($stopwatch.ElapsedMilliseconds -ge $nextTick) {
                
                if ($syncHash.PosMode -eq "Fixed") {
                    [AppForm]::SetCursorPos($syncHash.FixedX, $syncHash.FixedY)
                }

                if ($syncHash.ClickType -eq "Left") {
                    [AppForm]::mouse_event([AppForm]::MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                    [AppForm]::mouse_event([AppForm]::MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                } else {
                    [AppForm]::mouse_event([AppForm]::MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
                    [AppForm]::mouse_event([AppForm]::MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
                }

                $syncHash.ClickCount++
                
                # Calculate next tick with Jitter
                $baseInterval = $syncHash.Interval
                if ($syncHash.Jitter) {
                    $jitterAmt = $syncHash.JitterAmount
                    $actualInterval = $baseInterval + (Get-Random -Minimum -$jitterAmt -Maximum ($jitterAmt + 1))
                    if ($actualInterval -lt 10) { $actualInterval = 10 }
                } else {
                    $actualInterval = $baseInterval
                }

                $nextTick = $stopwatch.ElapsedMilliseconds + $actualInterval
            }
            Start-Sleep -Milliseconds 1
        }
        $syncHash.Running = $false
    })
    
    $syncHash.PSInstance = $psCmd
    $syncHash.Runspace = $runspace
    $syncHash.AsyncResult = $psCmd.BeginInvoke()
}

function Stop-Clicker {
    if (-not $syncHash.Running) { return }
    $syncHash.Stop = $true
    
    # Properly dispose of the runspace
    if ($syncHash.PSInstance) {
        $syncHash.PSInstance.EndInvoke($syncHash.AsyncResult)
        $syncHash.PSInstance.Dispose()
    }
    if ($syncHash.Runspace) {
        $syncHash.Runspace.Close()
    }
    $syncHash.PSInstance = $null
    $syncHash.Runspace = $null
    $syncHash.AsyncResult = $null

    # Update UI
    Toggle-Config $true
    
    $btnStart.Enabled = $true
    $btnStart.BackColor = $ColorAccent
    $btnStart.ForeColor = [Color]::White
    
    $btnStop.Enabled = $false
    $btnStop.BackColor = $ColorSurface
    $btnStop.ForeColor = $ColorText
    
    $lblStatus.Text = "Status: Idle"
    $lblStatus.ForeColor = $ColorMuted
}

# Wire Events
$form.OnStart = { Start-Clicker }
$form.OnStop = { Stop-Clicker }
$btnStart.Add_Click({ Start-Clicker })
$btnStop.Add_Click({ Stop-Clicker })

# Start
[void]$form.ShowDialog()
$uiTimer.Stop()