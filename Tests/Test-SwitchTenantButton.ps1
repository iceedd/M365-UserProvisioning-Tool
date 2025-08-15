#Requires -Version 7.0

<#
.SYNOPSIS
    Quick test to verify Switch Tenant button creation
.DESCRIPTION
    This script tests just the button creation part to debug visibility issues
#>

# Initialize Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Create a simple test form
$TestForm = New-Object System.Windows.Forms.Form
$TestForm.Text = "Switch Tenant Button Test"
$TestForm.Size = New-Object System.Drawing.Size(600, 200)
$TestForm.StartPosition = "CenterScreen"

# Create the connection panel (same as in main app)
$ConnectionPanel = New-Object System.Windows.Forms.Panel
$ConnectionPanel.Height = 70
$ConnectionPanel.Dock = "Top"
$ConnectionPanel.BackColor = [System.Drawing.Color]::LightSteelBlue
$ConnectionPanel.Padding = New-Object System.Windows.Forms.Padding(10)

# Create Connect button (same as main app)
$ConnectButton = New-Object System.Windows.Forms.Button
$ConnectButton.Text = "🔗 Connect to M365"
$ConnectButton.Size = New-Object System.Drawing.Size(180, 35)
$ConnectButton.Location = New-Object System.Drawing.Point(10, 15)
$ConnectButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$ConnectButton.BackColor = [System.Drawing.Color]::LightBlue

# Create Switch Tenant button (same as main app)
$SwitchTenantButton = New-Object System.Windows.Forms.Button
$SwitchTenantButton.Text = "🔄 Switch Tenant"
$SwitchTenantButton.Size = New-Object System.Drawing.Size(160, 35)
$SwitchTenantButton.Location = New-Object System.Drawing.Point(200, 15)  # Right next to Connect button
$SwitchTenantButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$SwitchTenantButton.Enabled = $true  # Enabled for testing
$SwitchTenantButton.BackColor = [System.Drawing.Color]::Orange  # High visibility color
$SwitchTenantButton.ForeColor = [System.Drawing.Color]::White   # White text for contrast

# Add click event for testing
$SwitchTenantButton.Add_Click({
    [System.Windows.Forms.MessageBox]::Show(
        "Switch Tenant button clicked successfully!",
        "Button Test",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
})

# Create Refresh button for comparison
$RefreshButton = New-Object System.Windows.Forms.Button
$RefreshButton.Text = "🔄 Refresh Data"
$RefreshButton.Size = New-Object System.Drawing.Size(140, 35)
$RefreshButton.Location = New-Object System.Drawing.Point(370, 15)
$RefreshButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Add all buttons to panel
$ConnectionPanel.Controls.AddRange(@($ConnectButton, $SwitchTenantButton, $RefreshButton))

# Add panel to form
$TestForm.Controls.Add($ConnectionPanel)

# Add a label for instructions
$InstructionLabel = New-Object System.Windows.Forms.Label
$InstructionLabel.Text = "Test Form: If you can see the orange 'Switch Tenant' button between Connect and Refresh, the button creation is working correctly."
$InstructionLabel.Location = New-Object System.Drawing.Point(10, 80)
$InstructionLabel.Size = New-Object System.Drawing.Size(580, 40)
$InstructionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$TestForm.Controls.Add($InstructionLabel)

Write-Host "🧪 Opening Switch Tenant button test form..." -ForegroundColor Cyan
Write-Host "Expected layout: [Connect] [Switch Tenant] [Refresh]" -ForegroundColor Yellow
Write-Host "The Switch Tenant button should be ORANGE with white text" -ForegroundColor Yellow

# Show the test form
$result = $TestForm.ShowDialog()

Write-Host "✅ Test completed" -ForegroundColor Green