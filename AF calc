Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# --- Global variables to store the detailed results for clipboard ---
$script:chadsvascResultText = ""
$script:hasbledResultText = ""

# --- Create the main form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Cardiovascular Risk Calculator"
$form.Size = New-Object System.Drawing.Size(500, 820)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false

# --- Patient Demographics Section ---
$demographicsGroup = New-Object System.Windows.Forms.GroupBox
$demographicsGroup.Location = New-Object System.Drawing.Point(15, 15)
$demographicsGroup.Size = New-Object System.Drawing.Size(460, 60)
$demographicsGroup.Text = "Patient Info"
$form.Controls.Add($demographicsGroup)

$ageLabel = New-Object System.Windows.Forms.Label
$ageLabel.Text = "Age:"
$ageLabel.Location = New-Object System.Drawing.Point(20, 25)
$ageLabel.AutoSize = $true
$demographicsGroup.Controls.Add($ageLabel)

$ageTextBox = New-Object System.Windows.Forms.TextBox
$ageTextBox.Location = New-Object System.Drawing.Point(55, 22)
$ageTextBox.Size = New-Object System.Drawing.Size(50, 20)
$demographicsGroup.Controls.Add($ageTextBox)

$sexLabel = New-Object System.Windows.Forms.Label
$sexLabel.Text = "Sex:"
$sexLabel.Location = New-Object System.Drawing.Point(150, 25)
$sexLabel.AutoSize = $true
$demographicsGroup.Controls.Add($sexLabel)

$sexComboBox = New-Object System.Windows.Forms.ComboBox
$sexComboBox.Location = New-Object System.Drawing.Point(185, 22)
$sexComboBox.Size = New-Object System.Drawing.Size(80, 20)
$sexComboBox.Items.AddRange(@("Male", "Female"))
$sexComboBox.SelectedIndex = 0 # Default to Male
$demographicsGroup.Controls.Add($sexComboBox)

# --- Risk Factors Section ---
$riskFactorsGroup = New-Object System.Windows.Forms.GroupBox
$riskFactorsGroup.Location = New-Object System.Drawing.Point(15, 85)
$riskFactorsGroup.Size = New-Object System.Drawing.Size(460, 300)
$riskFactorsGroup.Text = "Risk Factors"
$form.Controls.Add($riskFactorsGroup)

# Create and position all checkboxes in two columns
$checkboxes = @{
    "CHF (Congestive Heart Failure)" = New-Object System.Drawing.Point(20, 30);
    "Hypertension" = New-Object System.Drawing.Point(20, 55);
    "Diabetes" = New-Object System.Drawing.Point(20, 80);
    "Stroke / TIA / Thromboembolism" = New-Object System.Drawing.Point(20, 105);
    "Vascular Disease (Prior MI, PAD, Aortic Plaque)" = New-Object System.Drawing.Point(20, 130);
    "Abnormal Renal Function" = New-Object System.Drawing.Point(20, 155);
    "Abnormal Liver Function" = New-Object System.Drawing.Point(20, 180);
    "Bleeding History or Predisposition" = New-Object System.Drawing.Point(20, 205);
    "Labile INR" = New-Object System.Drawing.Point(20, 230);
    "Concomitant Drugs or Alcohol Use" = New-Object System.Drawing.Point(20, 255);
}

$controls = @{}
foreach ($text in $checkboxes.Keys) {
    $checkbox = New-Object System.Windows.Forms.CheckBox
    $checkbox.Text = $text
    $checkbox.Location = $checkboxes[$text]
    $checkbox.AutoSize = $true
    $riskFactorsGroup.Controls.Add($checkbox)
    $controls[$text] = $checkbox
}

# --- Calculate Button ---
$calculateButton = New-Object System.Windows.Forms.Button
$calculateButton.Location = New-Object System.Drawing.Point(180, 395)
$calculateButton.Size = New-Object System.Drawing.Size(120, 30)
$calculateButton.Text = "Calculate Scores"
$form.Controls.Add($calculateButton)

# --- Results Section ---
$resultsGroup = New-Object System.Windows.Forms.GroupBox
$resultsGroup.Location = New-Object System.Drawing.Point(15, 435)
$resultsGroup.Size = New-Object System.Drawing.Size(460, 260)
$resultsGroup.Text = "Results"
$form.Controls.Add($resultsGroup)

$resultsTextBox = New-Object System.Windows.Forms.RichTextBox
$resultsTextBox.Location = New-Object System.Drawing.Point(10, 20)
$resultsTextBox.Size = New-Object System.Drawing.Size(440, 230)
$resultsTextBox.ReadOnly = $true
$resultsTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$resultsGroup.Controls.Add($resultsTextBox)

# --- Copy Buttons Section ---
$copyCHADSButton = New-Object System.Windows.Forms.Button
$copyCHADSButton.Location = New-Object System.Drawing.Point(15, 710)
$copyCHADSButton.Size = New-Object System.Drawing.Size(140, 30)
$copyCHADSButton.Text = "Copy CHADS-VASc"
$form.Controls.Add($copyCHADSButton)

$copyHASBLEDButton = New-Object System.Windows.Forms.Button
$copyHASBLEDButton.Location = New-Object System.Drawing.Point(170, 710)
$copyHASBLEDButton.Size = New-Object System.Drawing.Size(140, 30)
$copyHASBLEDButton.Text = "Copy HAS-BLED"
$form.Controls.Add($copyHASBLEDButton)

$copyAllButton = New-Object System.Windows.Forms.Button
$copyAllButton.Location = New-Object System.Drawing.Point(325, 710)
$copyAllButton.Size = New-Object System.Drawing.Size(150, 30)
$copyAllButton.Text = "Copy Both Scores"
$form.Controls.Add($copyAllButton)

# --- Event Handlers ---

# Calculate Button Click Logic
$calculateButton.Add_Click({
    [int]$age = 0
    [void][int]::TryParse($ageTextBox.Text, [ref]$age)
    $isFemale = $sexComboBox.SelectedItem -eq "Female"

    # --- CHADS-VASc Calculation ---
    $chadsvascScore = 0
    $chadsvascDisplay = New-Object System.Text.StringBuilder
    $chadsvascDisplay.AppendLine("CHADS-VASc Score Components:") | Out-Null
    $chadsvascComponents = New-Object System.Collections.Generic.List[string]
    
    if ($controls["CHF (Congestive Heart Failure)"].Checked) { $chadsvascScore += 1; $chadsvascDisplay.AppendLine("- Congestive Heart Failure (+1)") | Out-Null; $chadsvascComponents.Add("Congestive Heart Failure (+1)") }
    if ($controls["Hypertension"].Checked) { $chadsvascScore += 1; $chadsvascDisplay.AppendLine("- Hypertension (+1)") | Out-Null; $chadsvascComponents.Add("Hypertension (+1)") }
    if ($age -ge 75) { $chadsvascScore += 2; $chadsvascDisplay.AppendLine("- Age >= 75 years (+2)") | Out-Null; $chadsvascComponents.Add("Age >= 75 years (+2)") }
    elseif ($age -ge 65) { $chadsvascScore += 1; $chadsvascDisplay.AppendLine("- Age 65-74 years (+1)") | Out-Null; $chadsvascComponents.Add("Age 65-74 years (+1)") }
    if ($controls["Diabetes"].Checked) { $chadsvascScore += 1; $chadsvascDisplay.AppendLine("- Diabetes Mellitus (+1)") | Out-Null; $chadsvascComponents.Add("Diabetes Mellitus (+1)") }
    if ($controls["Stroke / TIA / Thromboembolism"].Checked) { $chadsvascScore += 2; $chadsvascDisplay.AppendLine("- Stroke/TIA/Thromboembolism (+2)") | Out-Null; $chadsvascComponents.Add("Stroke/TIA/Thromboembolism (+2)") }
    if ($controls["Vascular Disease (Prior MI, PAD, Aortic Plaque)"].Checked) { $chadsvascScore += 1; $chadsvascDisplay.AppendLine("- Vascular Disease (+1)") | Out-Null; $chadsvascComponents.Add("Vascular Disease (+1)") }
    if ($isFemale) { $chadsvascScore += 1; $chadsvascDisplay.AppendLine("- Female Sex (+1)") | Out-Null; $chadsvascComponents.Add("Female Sex (+1)") }
    
    $chadsvascClipboardText = "CHADS-VASc Score: $chadsvascScore"
    if ($chadsvascComponents.Count -gt 0) {
        $chadsvascClipboardText += " [" + ($chadsvascComponents -join ', ') + "]"
    }
    $script:chadsvascResultText = $chadsvascClipboardText

    # --- HAS-BLED Calculation ---
    $hasbledScore = 0
    $hasbledDisplay = New-Object System.Text.StringBuilder
    $hasbledDisplay.AppendLine("HAS-BLED Score Components:") | Out-Null
    $hasbledComponents = New-Object System.Collections.Generic.List[string]

    if ($controls["Hypertension"].Checked) { $hasbledScore += 1; $hasbledDisplay.AppendLine("- Hypertension (+1)") | Out-Null; $hasbledComponents.Add("Hypertension (+1)") }
    if ($controls["Abnormal Renal Function"].Checked) { $hasbledScore += 1; $hasbledDisplay.AppendLine("- Abnormal Renal Function (+1)") | Out-Null; $hasbledComponents.Add("Abnormal Renal Function (+1)") }
    if ($controls["Abnormal Liver Function"].Checked) { $hasbledScore += 1; $hasbledDisplay.AppendLine("- Abnormal Liver Function (+1)") | Out-Null; $hasbledComponents.Add("Abnormal Liver Function (+1)") }
    if ($controls["Stroke / TIA / Thromboembolism"].Checked) { $hasbledScore += 1; $hasbledDisplay.AppendLine("- Stroke (+1)") | Out-Null; $hasbledComponents.Add("Stroke (+1)") }
    if ($controls["Bleeding History or Predisposition"].Checked) { $hasbledScore += 1; $hasbledDisplay.AppendLine("- Bleeding History/Predisposition (+1)") | Out-Null; $hasbledComponents.Add("Bleeding History/Predisposition (+1)") }
    if ($controls["Labile INR"].Checked) { $hasbledScore += 1; $hasbledDisplay.AppendLine("- Labile INR (+1)") | Out-Null; $hasbledComponents.Add("Labile INR (+1)") }
    if ($age -gt 65) { $hasbledScore += 1; $hasbledDisplay.AppendLine("- Elderly (Age > 65) (+1)") | Out-Null; $hasbledComponents.Add("Elderly (Age > 65) (+1)") }
    if ($controls["Concomitant Drugs or Alcohol Use"].Checked) { $hasbledScore += 1; $hasbledDisplay.AppendLine("- Drugs or Alcohol Use (+1)") | Out-Null; $hasbledComponents.Add("Drugs or Alcohol Use (+1)") }

    $hasbledClipboardText = "HAS-BLED Score: $hasbledScore"
    if ($hasbledComponents.Count -gt 0) {
        $hasbledClipboardText += " [" + ($hasbledComponents -join ', ') + "]"
    }
    $script:hasbledResultText = $hasbledClipboardText

    # --- Update Display ---
    $fullDisplayText = "CHADS-VASc Score: $chadsvascScore`n" + $chadsvascDisplay.ToString() + "`nHAS-BLED Score: $hasbledScore`n" + $hasbledDisplay.ToString()
    $resultsTextBox.Text = $fullDisplayText
})

# Copy Button Click Logic
$copyCHADSButton.Add_Click({
    if (-not [string]::IsNullOrEmpty($script:chadsvascResultText)) {
        [System.Windows.Forms.Clipboard]::SetText($script:chadsvascResultText)
    }
})

$copyHASBLEDButton.Add_Click({
    if (-not [string]::IsNullOrEmpty($script:hasbledResultText)) {
        [System.Windows.Forms.Clipboard]::SetText($script:hasbledResultText)
    }
})

$copyAllButton.Add_Click({
    if (-not [string]::IsNullOrEmpty($script:chadsvascResultText)) {
        $allText = "$($script:chadsvascResultText)`n$($script:hasbledResultText)"
        [System.Windows.Forms.Clipboard]::SetText($allText)
    }
})

# --- Display the form ---
$form.ShowDialog()

