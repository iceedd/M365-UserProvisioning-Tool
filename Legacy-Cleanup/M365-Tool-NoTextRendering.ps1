#Requires -Version 7.0

<#
.SYNOPSIS
    M365 User Provisioning Tool - No SetCompatibleTextRenderingDefault Version
    
.DESCRIPTION
    This version completely bypasses the SetCompatibleTextRenderingDefault call,
    which resolves the initialization error. Most PowerShell Windows Forms apps
    work perfectly fine without this call.
    
.NOTES
    This version intentionally skips SetCompatibleTextRenderingDefault to avoid
    the timing issues in your environment.
#>

[CmdletBinding()]
param()

Write-Host "M365 User Provisioning Tool - Bypass Edition" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "ℹ️  This version bypasses SetCompatibleTextRenderingDefault" -ForegroundColor Yellow
Write-Host ""

try {
    Write-Host "🔧 Initializing Windows Forms (bypass mode)..." -ForegroundColor Cyan
    
    # Load Windows Forms assemblies
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    Write-Host "   ✅ Assemblies loaded" -ForegroundColor Green
    
    # Enable visual styles
    [System.Windows.Forms.Application]::EnableVisualStyles()
    Write-Host "   ✅ Visual styles enabled" -ForegroundColor Green
    
    # SKIP SetCompatibleTextRenderingDefault - this is what was causing the error!
    Write-Host "   ⏭️  Skipping SetCompatibleTextRenderingDefault (not required)" -ForegroundColor Yellow
    
    Write-Host "✅ Windows Forms ready!" -ForegroundColor Green
    Write-Host ""

    # Create main form
    Write-Host "🖥️  Creating main application window..." -ForegroundColor Green
    
    $MainForm = New-Object System.Windows.Forms.Form
    $MainForm.Text = "M365 User Provisioning Tool - Bypass Edition"
    $MainForm.Size = New-Object System.Drawing.Size(1000, 700)
    $MainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $MainForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    # Create connection panel
    $ConnectionPanel = New-Object System.Windows.Forms.Panel
    $ConnectionPanel.Height = 80
    $ConnectionPanel.Dock = [System.Windows.Forms.DockStyle]::Top
    $ConnectionPanel.BackColor = [System.Drawing.Color]::LightBlue
    $ConnectionPanel.Padding = New-Object System.Windows.Forms.Padding(10)
    
    $ConnectButton = New-Object System.Windows.Forms.Button
    $ConnectButton.Text = "🔗 Connect to Microsoft 365"
    $ConnectButton.Size = New-Object System.Drawing.Size(200, 35)
    $ConnectButton.Location = New-Object System.Drawing.Point(10, 20)
    $ConnectButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    
    $StatusLabel = New-Object System.Windows.Forms.Label
    $StatusLabel.Text = "Ready - Click Connect to start"
    $StatusLabel.Location = New-Object System.Drawing.Point(220, 25)
    $StatusLabel.Size = New-Object System.Drawing.Size(400, 25)
    $StatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    
    $ConnectionPanel.Controls.Add($ConnectButton)
    $ConnectionPanel.Controls.Add($StatusLabel)
    
    # Create main content area
    $ContentPanel = New-Object System.Windows.Forms.Panel
    $ContentPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $ContentPanel.Padding = New-Object System.Windows.Forms.Padding(20)
    
    # Create tab control
    $TabControl = New-Object System.Windows.Forms.TabControl
    $TabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    
    # User Creation Tab
    $UserTab = New-Object System.Windows.Forms.TabPage
    $UserTab.Text = "👤 Create User"
    $UserTab.Padding = New-Object System.Windows.Forms.Padding(10)
    
    # Simple user creation form
    $UserForm = New-Object System.Windows.Forms.TableLayoutPanel
    $UserForm.ColumnCount = 2
    $UserForm.RowCount = 6
    $UserForm.Dock = [System.Windows.Forms.DockStyle]::Top
    $UserForm.Height = 200
    
    # Add form controls
    $FirstNameLabel = New-Object System.Windows.Forms.Label
    $FirstNameLabel.Text = "First Name:"
    $FirstNameLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    
    $FirstNameBox = New-Object System.Windows.Forms.TextBox
    $FirstNameBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    
    $LastNameLabel = New-Object System.Windows.Forms.Label
    $LastNameLabel.Text = "Last Name:"
    $LastNameLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    
    $LastNameBox = New-Object System.Windows.Forms.TextBox
    $LastNameBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    
    $EmailLabel = New-Object System.Windows.Forms.Label
    $EmailLabel.Text = "Email:"
    $EmailLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    
    $EmailBox = New-Object System.Windows.Forms.TextBox
    $EmailBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    
    $CreateButton = New-Object System.Windows.Forms.Button
    $CreateButton.Text = "Create User"
    $CreateButton.Enabled = $false
    $CreateButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    
    # Add controls to table
    $UserForm.Controls.Add($FirstNameLabel, 0, 0)
    $UserForm.Controls.Add($FirstNameBox, 1, 0)
    $UserForm.Controls.Add($LastNameLabel, 0, 1)
    $UserForm.Controls.Add($LastNameBox, 1, 1)
    $UserForm.Controls.Add($EmailLabel, 0, 2)
    $UserForm.Controls.Add($EmailBox, 1, 2)
    $UserForm.Controls.Add($CreateButton, 1, 3)
    
    $UserTab.Controls.Add($UserForm)
    $TabControl.TabPages.Add($UserTab)
    
    # Activity Log Tab
    $LogTab = New-Object System.Windows.Forms.TabPage
    $LogTab.Text = "📋 Activity Log"
    
    $LogTextBox = New-Object System.Windows.Forms.TextBox
    $LogTextBox.Multiline = $true
    $LogTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $LogTextBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $LogTextBox.ReadOnly = $true
    $LogTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $LogTextBox.Text = "$(Get-Date): Application started successfully`r`n$(Get-Date): Windows Forms initialized without SetCompatibleTextRenderingDefault`r`n$(Get-Date): Ready for user input"
    
    $LogTab.Controls.Add($LogTextBox)
    $TabControl.TabPages.Add($LogTab)
    
    $ContentPanel.Controls.Add($TabControl)
    
    # Add event handlers
    $ConnectButton.Add_Click({
        try {
            $StatusLabel.Text = "Connecting to Microsoft 365..."
            $StatusLabel.ForeColor = [System.Drawing.Color]::Blue
            [System.Windows.Forms.Application]::DoEvents()
            
            # Simulate connection (replace with actual Microsoft Graph connection)
            Start-Sleep 2
            
            $ConnectButton.Text = "✅ Connected to M365"
            $ConnectButton.BackColor = [System.Drawing.Color]::LightGreen
            $StatusLabel.Text = "Connected successfully! You can now create users."
            $StatusLabel.ForeColor = [System.Drawing.Color]::DarkGreen
            $CreateButton.Enabled = $true
            
            $LogTextBox.AppendText("$(Get-Date): Connected to Microsoft 365`r`n")
        }
        catch {
            $StatusLabel.Text = "Connection failed: $($_.Exception.Message)"
            $StatusLabel.ForeColor = [System.Drawing.Color]::Red
            $LogTextBox.AppendText("$(Get-Date): Connection failed - $($_.Exception.Message)`r`n")
        }
    })
    
    $CreateButton.Add_Click({
        if ($FirstNameBox.Text -and $LastNameBox.Text -and $EmailBox.Text) {
            $LogTextBox.AppendText("$(Get-Date): Creating user - $($FirstNameBox.Text) $($LastNameBox.Text) ($($EmailBox.Text))`r`n")
            [System.Windows.Forms.MessageBox]::Show(
                "User creation would be processed here:`n`nName: $($FirstNameBox.Text) $($LastNameBox.Text)`nEmail: $($EmailBox.Text)", 
                "User Creation", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            # Clear form
            $FirstNameBox.Clear()
            $LastNameBox.Clear()
            $EmailBox.Clear()
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please fill in all fields.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })
    
    # Add panels to main form
    $MainForm.Controls.Add($ContentPanel)
    $MainForm.Controls.Add($ConnectionPanel)
    
    # Form events
    $MainForm.Add_Load({
        Write-Host "✅ Application window loaded successfully" -ForegroundColor Green
        $LogTextBox.AppendText("$(Get-Date): Application window loaded`r`n")
    })
    
    $MainForm.Add_FormClosing({
        $Result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to exit?", "Confirm Exit", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($Result -eq [System.Windows.Forms.DialogResult]::No) {
            $_.Cancel = $true
        }
    })
    
    Write-Host "🚀 Launching M365 User Provisioning Tool..." -ForegroundColor Green
    Write-Host ""
    
    # Show the form
    [System.Windows.Forms.Application]::Run($MainForm)
    
    Write-Host ""
    Write-Host "👋 Application closed successfully" -ForegroundColor Yellow
    
} catch {
    Write-Host ""
    Write-Host "❌ ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Stack Trace:" -ForegroundColor Yellow
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
    
    if ($_.Exception.InnerException) {
        Write-Host ""
        Write-Host "Inner Exception: $($_.Exception.InnerException.Message)" -ForegroundColor Yellow
    }
    
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    Read-Host
}