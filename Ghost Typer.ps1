Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Ghost Typer"
$form.Size = New-Object System.Drawing.Size(500, 400)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Text input label
$label = New-Object System.Windows.Forms.Label
$label.Text = "Enter text to auto-type:"
$label.Location = New-Object System.Drawing.Point(10, 10)
$label.AutoSize = $true
$label.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($label)

# Text box with character counter
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Multiline = $true
$textBox.ScrollBars = "Vertical"
$textBox.Size = New-Object System.Drawing.Size(460, 180)
$textBox.Location = New-Object System.Drawing.Point(10, 35)
$textBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($textBox)

# Character counter
$charLabel = New-Object System.Windows.Forms.Label
$charLabel.Text = "Characters: 0"
$charLabel.Location = New-Object System.Drawing.Point(10, 220)
$charLabel.AutoSize = $true
$charLabel.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($charLabel)

$textBox.Add_TextChanged({
    $charLabel.Text = "Characters: $($textBox.Text.Length)"
})

# Speed section
$speedLabel = New-Object System.Windows.Forms.Label
$speedLabel.Text = "Typing speed (milliseconds per character):"
$speedLabel.Location = New-Object System.Drawing.Point(10, 245)
$speedLabel.AutoSize = $true
$speedLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($speedLabel)

# Speed display label
$speedValueLabel = New-Object System.Windows.Forms.Label
$speedValueLabel.Text = "50 ms"
$speedValueLabel.Location = New-Object System.Drawing.Point(410, 245)
$speedValueLabel.AutoSize = $true
$speedValueLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($speedValueLabel)

# Speed slider
$speedSlider = New-Object System.Windows.Forms.TrackBar
$speedSlider.Minimum = 5
$speedSlider.Maximum = 200
$speedSlider.Value = 50
$speedSlider.TickFrequency = 25
$speedSlider.Size = New-Object System.Drawing.Size(460, 45)
$speedSlider.Location = New-Object System.Drawing.Point(10, 270)
$form.Controls.Add($speedSlider)

# Update speed display
$speedSlider.Add_ValueChanged({
    $speedValueLabel.Text = "$($speedSlider.Value) ms"
})

# Start button
$startButton = New-Object System.Windows.Forms.Button
$startButton.Text = "START TYPING (3 second delay)"
$startButton.Size = New-Object System.Drawing.Size(460, 40)
$startButton.Location = New-Object System.Drawing.Point(10, 320)
$startButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$startButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$startButton.ForeColor = [System.Drawing.Color]::White
$startButton.FlatStyle = "Flat"
$form.Controls.Add($startButton)

# Escape handling
$form.KeyPreview = $true
$script:shouldStop = $false

$form.Add_KeyDown({
    if ($_.KeyCode -eq "Escape") {
        $script:shouldStop = $true
    }
})

# Start typing functionality
$startButton.Add_Click({
    $text = $textBox.Text
    
    # Validation
    if ([string]::IsNullOrWhiteSpace($text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter some text to type.",
            "No Text Entered",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $delay = $speedSlider.Value
    $script:shouldStop = $false
    
    # Hide form and show countdown
    $form.WindowState = "Minimized"
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Click in your target window now.`n`nTyping will start in 3 seconds.`n`nPress ESC to cancel during typing.",
        "Get Ready",
        [System.Windows.Forms.MessageBoxButtons]::OKCancel,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
    
    if ($result -eq "Cancel") {
        $form.WindowState = "Normal"
        return
    }
    
    Start-Sleep -Seconds 3

    # Type each character with escape handling
    $typedChars = 0
    foreach ($char in $text.ToCharArray()) {
        if ($script:shouldStop) {
            [System.Media.SystemSounds]::Asterisk.Play()
            [System.Windows.Forms.MessageBox]::Show(
                "Typing cancelled. Typed $typedChars of $($text.Length) characters.",
                "Cancelled",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            $form.WindowState = "Normal"
            return
        }
        
        # Escape special characters for SendKeys
        $charToSend = switch ($char) {
            '+' { '{+}' }
            '^' { '{^}' }
            '%' { '{%}' }
            '~' { '{~}' }
            '(' { '{(}' }
            ')' { '{)}' }
            '[' { '{[}' }
            ']' { '{]}' }
            '{' { '{{}' }
            '}' { '{}}' }
            default { $char }
        }
        
        try {
            [System.Windows.Forms.SendKeys]::SendWait($charToSend)
            $typedChars++
            Start-Sleep -Milliseconds $delay
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Error typing character: $char`n`nTyped $typedChars of $($text.Length) characters.",
                "Typing Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            $form.WindowState = "Normal"
            return
        }
    }

    # Success
    [System.Media.SystemSounds]::Beep.Play()
    [System.Windows.Forms.MessageBox]::Show(
        "Successfully typed $typedChars characters!",
        "Typing Completed",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
    
    $form.WindowState = "Normal"
})

# Show form
[void]$form.ShowDialog()
