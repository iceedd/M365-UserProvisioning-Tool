#Requires -Version 7.0
<#
.SYNOPSIS
    M365 GUI Module - Complete Windows Forms Interface
.DESCRIPTION
    Contains all GUI components for M365 User Provisioning Tool
.NOTES
    Version: 1.0.0 - Extracted from Legacy Script
    Author: Tom Mortiboys
#>

# Load Windows Forms assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Module-scoped variables for UI components
$Script:MainForm = $null
$Script:StatusLabel = $null
$Script:StatusStrip = $null
$Script:TabControl = $null

# Form controls
$Script:ConnectButton = $null
$Script:DisconnectButton = $null
$Script:RefreshButton = $null
$Script:CreateUserButton = $null
$Script:ImportCSVButton = $null

# Input controls
$Script:DisplayNameTextBox = $null
$Script:FirstNameTextBox = $null
$Script:LastNameTextBox = $null
$Script:UsernameTextBox = $null
$Script:PasswordTextBox = $null
$Script:DepartmentTextBox = $null
$Script:JobTitleTextBox = $null

# Dropdowns
$Script:DomainDropdown = $null
$Script:ManagerDropdown = $null
$Script:OfficeDropdown = $null
$Script:LicenseDropdown = $null

# Lists and grids
$Script:GroupsCheckedListBox = $null
$Script:UsersDataGridView = $null
$Script:GroupsDataGridView = $null
$Script:LicensesDataGridView = $null

# File browser
$Script:FilePathTextBox = $null

# Progress and logging
$Script:ProgressBar = $null
$Script:ProgressLabel = $null
$Script:LogTextBox = $null

# Connection info labels
$Script:ConnectionInfoLabel = $null
$Script:ConnectionDetailsLabel = $null

# UK-based locations configuration
$Script:UKLocations = @(
    "United Kingdom - London",
    "United Kingdom - Manchester", 
    "United Kingdom - Birmingham",
    "United Kingdom - Leeds",
    "United Kingdom - Glasgow",
    "United Kingdom - Edinburgh",
    "United Kingdom - Bristol",
    "United Kingdom - Liverpool",
    "Remote/Home Working - UK",
    "Office - Head Office",
    "Office - Branch Office"
)

function Start-M365ProvisioningTool {
    <#
    .SYNOPSIS
        Main application launcher - starts the GUI interface
    #>
    try {
        Write-Host "🚀 Starting M365 User Provisioning Tool GUI..." -ForegroundColor Cyan
        
        [System.Windows.Forms.Application]::EnableVisualStyles()
        [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

        # Create and show main form
        $Script:MainForm = New-MainForm
        
        Write-Host "🖥️ Launching GUI interface..." -ForegroundColor Green

        
        # Show the form and wait for it to close
        $Result = $Script:MainForm.ShowDialog()
        
        Write-Host "📱 Application closed normally" -ForegroundColor Gray
        return $Result
        
    }
    catch {
        $ErrorMessage = "GUI failed to start: $($_.Exception.Message)"
        Write-Error $ErrorMessage
        
        try {
            [System.Windows.Forms.MessageBox]::Show(
                $ErrorMessage,
                "Application Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
        catch {
            Write-Host "Error: $ErrorMessage" -ForegroundColor Red
        }
        
        throw
    }
}

function New-MainForm {
    <#
    .SYNOPSIS
        Creates the main application form with tabbed interface
    #>
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "M365 User Provisioning Tool - Enterprise Edition 2025"
    $Form.Size = New-Object System.Drawing.Size(1400, 900)
    $Form.StartPosition = "CenterScreen"
    $Form.MinimumSize = New-Object System.Drawing.Size(1200, 800)
    $Form.MaximizeBox = $true
    $Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("$env:SystemRoot\System32\shell32.dll")
    $Form.WindowState = "Maximized"
    
    # Status strip at bottom
    $Script:StatusStrip = New-Object System.Windows.Forms.StatusStrip
    $Script:StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $Script:StatusLabel.Text = "Ready - Click 'Connect to M365' to begin"
    $Script:StatusLabel.Spring = $true
    $null = $Script:StatusStrip.Items.Add($Script:StatusLabel)
    $null = $Form.Controls.Add($Script:StatusStrip)
    
    # Connection panel at top
    $ConnectionPanel = New-Object System.Windows.Forms.Panel
    $ConnectionPanel.Height = 90
    $ConnectionPanel.Dock = "Top"
    $ConnectionPanel.BackColor = [System.Drawing.Color]::LightSteelBlue
    
    # Connection buttons
    $Script:ConnectButton = New-Object System.Windows.Forms.Button
    $Script:ConnectButton.Text = "Connect to M365"
    $Script:ConnectButton.Size = New-Object System.Drawing.Size(120, 30)
    $Script:ConnectButton.Location = New-Object System.Drawing.Point(10, 10)
    $Script:ConnectButton.BackColor = [System.Drawing.Color]::LightGreen
    $Script:ConnectButton.Add_Click({
        try {
            $Script:ConnectButton.Enabled = $false
            $Script:ConnectButton.Text = "Connecting..."
            $Script:ConnectButton.Refresh()
            
            Update-StatusLabel "Connecting to Microsoft Graph..."
            $Result = Connect-ToMicrosoftGraph
            
            if ($Result.Success) {
                Update-UIAfterConnection
                Update-StatusLabel "Connected successfully to $($Result.TenantId)"
                
                [System.Windows.Forms.MessageBox]::Show(
                    "Successfully connected to Microsoft Graph!`n`nTenant: $($Result.TenantId)`nAccount: $($Result.Account)",
                    "Connection Successful",
                    "OK",
                    "Information"
                )
            }
            else {
                Update-StatusLabel "Connection failed: $($Result.Message)"
                [System.Windows.Forms.MessageBox]::Show(
                    "Failed to connect: $($Result.Message)",
                    "Connection Failed",
                    "OK",
                    "Error"
                )
                $Script:ConnectButton.Enabled = $true
                $Script:ConnectButton.Text = "Connect to M365"
            }
        }
        catch {
            Update-StatusLabel "Connection error occurred"
            [System.Windows.Forms.MessageBox]::Show(
                "Connection error: $($_.Exception.Message)",
                "Error",
                "OK",
                "Error"
            )
            $Script:ConnectButton.Enabled = $true
            $Script:ConnectButton.Text = "Connect to M365"
        }
    })
    
    $Script:DisconnectButton = New-Object System.Windows.Forms.Button
    $Script:DisconnectButton.Text = "Disconnect"
    $Script:DisconnectButton.Size = New-Object System.Drawing.Size(100, 30)
    $Script:DisconnectButton.Location = New-Object System.Drawing.Point(140, 10)
    $Script:DisconnectButton.Enabled = $false
    $Script:DisconnectButton.BackColor = [System.Drawing.Color]::LightCoral
    $Script:DisconnectButton.Add_Click({
        $Result = Disconnect-FromMicrosoftGraph
        if ($Result.Success) {
            Update-UIAfterDisconnection
            Update-StatusLabel "Disconnected from Microsoft Graph"
        }
    })
    
    $Script:RefreshButton = New-Object System.Windows.Forms.Button
    $Script:RefreshButton.Text = "Refresh Data"
    $Script:RefreshButton.Size = New-Object System.Drawing.Size(100, 30)
    $Script:RefreshButton.Location = New-Object System.Drawing.Point(250, 10)
    $Script:RefreshButton.Enabled = $false
    $Script:RefreshButton.Add_Click({
        Update-StatusLabel "Refreshing tenant data..."
        Start-TenantDiscovery
        Refresh-TenantDataViews
        Update-StatusLabel "Tenant data refreshed"
    })
    
    # Connection info labels
    $Script:ConnectionInfoLabel = New-Object System.Windows.Forms.Label
    $Script:ConnectionInfoLabel.Text = "Status: Disconnected"
    $Script:ConnectionInfoLabel.Location = New-Object System.Drawing.Point(10, 45)
    $Script:ConnectionInfoLabel.Size = New-Object System.Drawing.Size(600, 20)
    $Script:ConnectionInfoLabel.ForeColor = [System.Drawing.Color]::DarkBlue
    
    $Script:ConnectionDetailsLabel = New-Object System.Windows.Forms.Label
    $Script:ConnectionDetailsLabel.Text = "Click 'Connect to M365' to begin"
    $Script:ConnectionDetailsLabel.Location = New-Object System.Drawing.Point(10, 65)
    $Script:ConnectionDetailsLabel.Size = New-Object System.Drawing.Size(600, 20)
    $Script:ConnectionDetailsLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    
    $null = $ConnectionPanel.Controls.Add($Script:ConnectButton)
    $null = $ConnectionPanel.Controls.Add($Script:DisconnectButton)
    $null = $ConnectionPanel.Controls.Add($Script:RefreshButton)
    $null = $ConnectionPanel.Controls.Add($Script:ConnectionInfoLabel)
    $null = $ConnectionPanel.Controls.Add($Script:ConnectionDetailsLabel)
    $null = $Form.Controls.Add($ConnectionPanel)
    
    # Tab control
    $Script:TabControl = New-Object System.Windows.Forms.TabControl
    $Script:TabControl.Location = New-Object System.Drawing.Point(0, 90)
    $Script:TabControl.Size = New-Object System.Drawing.Size($Form.ClientSize.Width, ($Form.ClientSize.Height - 120))
    $Script:TabControl.Anchor = "Top,Bottom,Left,Right"
    $Script:TabControl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    # Create tabs
    New-UserCreationTab
    New-BulkImportTab
    New-TenantDataTab
    New-ActivityLogTab
    
    $null = $Form.Controls.Add($Script:TabControl)
    
    # Handle form closing
    $Form.Add_FormClosing({
        param($Sender, $e)
        
        $Status = Get-M365AuthenticationStatus
        if ($Status.GraphConnected) {
            $Result = [System.Windows.Forms.MessageBox]::Show(
                "You are still connected to Microsoft Graph. Do you want to disconnect before closing?",
                "Confirm Exit",
                "YesNoCancel",
                "Question"
            )
            
            if ($Result -eq "Yes") {
                Disconnect-FromMicrosoftGraph | Out-Null
            }
            elseif ($Result -eq "Cancel") {
                $e.Cancel = $true
                return
            }
        }
    })
    
    return $Form
}

function New-UserCreationTab {
    <#
    .SYNOPSIS
        Creates the user creation tab interface
    #>
    $Tab = New-Object System.Windows.Forms.TabPage
    $Tab.Text = "Create User"
    $Tab.BackColor = [System.Drawing.Color]::White
    
    $MainPanel = New-Object System.Windows.Forms.Panel
    $MainPanel.Dock = "Fill"
    $MainPanel.AutoScroll = $true
    $MainPanel.Padding = New-Object System.Windows.Forms.Padding(15)
    
    # User details group (left side)
    $UserDetailsGroup = New-Object System.Windows.Forms.GroupBox
    $UserDetailsGroup.Text = "User Details"
    $UserDetailsGroup.Location = New-Object System.Drawing.Point(20, 20)
    $UserDetailsGroup.Size = New-Object System.Drawing.Size(480, 350)
    
    # Display Name
    $DisplayNameLabel = New-Object System.Windows.Forms.Label
    $DisplayNameLabel.Text = "Display Name *:"
    $DisplayNameLabel.Location = New-Object System.Drawing.Point(10, 30)
    $DisplayNameLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:DisplayNameTextBox = New-Object System.Windows.Forms.TextBox
    $Script:DisplayNameTextBox.Location = New-Object System.Drawing.Point(120, 28)
    $Script:DisplayNameTextBox.Size = New-Object System.Drawing.Size(200, 20)
    
    # First Name
    $FirstNameLabel = New-Object System.Windows.Forms.Label
    $FirstNameLabel.Text = "First Name:"
    $FirstNameLabel.Location = New-Object System.Drawing.Point(10, 60)
    $FirstNameLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:FirstNameTextBox = New-Object System.Windows.Forms.TextBox
    $Script:FirstNameTextBox.Location = New-Object System.Drawing.Point(120, 58)
    $Script:FirstNameTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $Script:FirstNameTextBox.Add_TextChanged({
        if (-not [string]::IsNullOrWhiteSpace($Script:FirstNameTextBox.Text) -or -not [string]::IsNullOrWhiteSpace($Script:LastNameTextBox.Text)) {
            $FirstName = $Script:FirstNameTextBox.Text.Trim()
            $LastName = $Script:LastNameTextBox.Text.Trim()
            $Script:DisplayNameTextBox.Text = "$FirstName $LastName".Trim()
        }
    })
    
    # Last Name
    $LastNameLabel = New-Object System.Windows.Forms.Label
    $LastNameLabel.Text = "Last Name:"
    $LastNameLabel.Location = New-Object System.Drawing.Point(10, 90)
    $LastNameLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:LastNameTextBox = New-Object System.Windows.Forms.TextBox
    $Script:LastNameTextBox.Location = New-Object System.Drawing.Point(120, 88)
    $Script:LastNameTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $Script:LastNameTextBox.Add_TextChanged({
        if (-not [string]::IsNullOrWhiteSpace($Script:FirstNameTextBox.Text) -or -not [string]::IsNullOrWhiteSpace($Script:LastNameTextBox.Text)) {
            $FirstName = $Script:FirstNameTextBox.Text.Trim()
            $LastName = $Script:LastNameTextBox.Text.Trim()
            $Script:DisplayNameTextBox.Text = "$FirstName $LastName".Trim()
        }
    })
    
    # Username
    $UsernameLabel = New-Object System.Windows.Forms.Label
    $UsernameLabel.Text = "Username *:"
    $UsernameLabel.Location = New-Object System.Drawing.Point(10, 120)
    $UsernameLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:UsernameTextBox = New-Object System.Windows.Forms.TextBox
    $Script:UsernameTextBox.Location = New-Object System.Drawing.Point(120, 118)
    $Script:UsernameTextBox.Size = New-Object System.Drawing.Size(130, 20)
    
    # Domain dropdown
    $DomainLabel = New-Object System.Windows.Forms.Label
    $DomainLabel.Text = "@"
    $DomainLabel.Location = New-Object System.Drawing.Point(255, 120)
    $DomainLabel.Size = New-Object System.Drawing.Size(15, 20)
    
    $Script:DomainDropdown = New-Object System.Windows.Forms.ComboBox
    $Script:DomainDropdown.Location = New-Object System.Drawing.Point(270, 118)
    $Script:DomainDropdown.Size = New-Object System.Drawing.Size(180, 20)
    $Script:DomainDropdown.DropDownStyle = "DropDownList"
    
    # Password
    $PasswordLabel = New-Object System.Windows.Forms.Label
    $PasswordLabel.Text = "Password *:"
    $PasswordLabel.Location = New-Object System.Drawing.Point(10, 150)
    $PasswordLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:PasswordTextBox = New-Object System.Windows.Forms.TextBox
    $Script:PasswordTextBox.Location = New-Object System.Drawing.Point(120, 148)
    $Script:PasswordTextBox.Size = New-Object System.Drawing.Size(150, 20)
    $Script:PasswordTextBox.UseSystemPasswordChar = $true
    
    # Generate password button
    $GeneratePasswordButton = New-Object System.Windows.Forms.Button
    $GeneratePasswordButton.Text = "Generate"
    $GeneratePasswordButton.Location = New-Object System.Drawing.Point(280, 148)
    $GeneratePasswordButton.Size = New-Object System.Drawing.Size(70, 22)
    $GeneratePasswordButton.Add_Click({
        $GeneratedPassword = New-SecurePassword
        $Script:PasswordTextBox.Text = $GeneratedPassword
        [System.Windows.Forms.MessageBox]::Show("Generated password: $GeneratedPassword`n`nPlease save this password securely!", "Password Generated", "OK", "Information")
    })
    
    # Department
    $DepartmentLabel = New-Object System.Windows.Forms.Label
    $DepartmentLabel.Text = "Department:"
    $DepartmentLabel.Location = New-Object System.Drawing.Point(10, 180)
    $DepartmentLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:DepartmentTextBox = New-Object System.Windows.Forms.TextBox
    $Script:DepartmentTextBox.Location = New-Object System.Drawing.Point(120, 178)
    $Script:DepartmentTextBox.Size = New-Object System.Drawing.Size(200, 20)
    
    # Job Title
    $JobTitleLabel = New-Object System.Windows.Forms.Label
    $JobTitleLabel.Text = "Job Title:"
    $JobTitleLabel.Location = New-Object System.Drawing.Point(10, 210)
    $JobTitleLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:JobTitleTextBox = New-Object System.Windows.Forms.TextBox
    $Script:JobTitleTextBox.Location = New-Object System.Drawing.Point(120, 208)
    $Script:JobTitleTextBox.Size = New-Object System.Drawing.Size(200, 20)
    
    # Office Location
    $OfficeLabel = New-Object System.Windows.Forms.Label
    $OfficeLabel.Text = "Office Location:"
    $OfficeLabel.Location = New-Object System.Drawing.Point(10, 240)
    $OfficeLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:OfficeDropdown = New-Object System.Windows.Forms.ComboBox
    $Script:OfficeDropdown.Location = New-Object System.Drawing.Point(120, 238)
    $Script:OfficeDropdown.Size = New-Object System.Drawing.Size(250, 20)
    $Script:OfficeDropdown.DropDownStyle = "DropDownList"
    $Script:UKLocations | ForEach-Object { $null = $Script:OfficeDropdown.Items.Add($_) }
    
    # Add all user detail controls
    $UserDetailsGroup.Controls.AddRange(@(
        $DisplayNameLabel, $Script:DisplayNameTextBox,
        $FirstNameLabel, $Script:FirstNameTextBox,
        $LastNameLabel, $Script:LastNameTextBox,
        $UsernameLabel, $Script:UsernameTextBox,
        $DomainLabel, $Script:DomainDropdown,
        $PasswordLabel, $Script:PasswordTextBox, $GeneratePasswordButton,
        $DepartmentLabel, $Script:DepartmentTextBox,
        $JobTitleLabel, $Script:JobTitleTextBox,
        $OfficeLabel, $Script:OfficeDropdown
    ))
    
    # Management group (right side)
    $ManagementGroup = New-Object System.Windows.Forms.GroupBox
    $ManagementGroup.Text = "Management & Licensing"
    $ManagementGroup.Location = New-Object System.Drawing.Point(520, 20)
    $ManagementGroup.Size = New-Object System.Drawing.Size(450, 350)
    
    # Manager dropdown
    $ManagerLabel = New-Object System.Windows.Forms.Label
    $ManagerLabel.Text = "Manager:"
    $ManagerLabel.Location = New-Object System.Drawing.Point(10, 30)
    $ManagerLabel.Size = New-Object System.Drawing.Size(80, 20)
    
    $Script:ManagerDropdown = New-Object System.Windows.Forms.ComboBox
    $Script:ManagerDropdown.Location = New-Object System.Drawing.Point(100, 28)
    $Script:ManagerDropdown.Size = New-Object System.Drawing.Size(330, 20)
    $Script:ManagerDropdown.DropDownStyle = "DropDownList"
    
    # License dropdown
    $LicenseLabel = New-Object System.Windows.Forms.Label
    $LicenseLabel.Text = "License Type:"
    $LicenseLabel.Location = New-Object System.Drawing.Point(10, 60)
    $LicenseLabel.Size = New-Object System.Drawing.Size(80, 20)
    
    $Script:LicenseDropdown = New-Object System.Windows.Forms.ComboBox
    $Script:LicenseDropdown.Location = New-Object System.Drawing.Point(100, 58)
    $Script:LicenseDropdown.Size = New-Object System.Drawing.Size(330, 20)
    $Script:LicenseDropdown.DropDownStyle = "DropDownList"
    
    $ManagementGroup.Controls.AddRange(@(
        $ManagerLabel, $Script:ManagerDropdown,
        $LicenseLabel, $Script:LicenseDropdown
    ))
    
    # Groups selection
    $GroupsGroup = New-Object System.Windows.Forms.GroupBox
    $GroupsGroup.Text = "Group Membership"
    $GroupsGroup.Location = New-Object System.Drawing.Point(20, 390)
    $GroupsGroup.Size = New-Object System.Drawing.Size(950, 200)
    
    $Script:GroupsCheckedListBox = New-Object System.Windows.Forms.CheckedListBox
    $Script:GroupsCheckedListBox.Location = New-Object System.Drawing.Point(10, 20)
    $Script:GroupsCheckedListBox.Size = New-Object System.Drawing.Size(930, 170)
    $Script:GroupsCheckedListBox.CheckOnClick = $true
    
    $GroupsGroup.Controls.Add($Script:GroupsCheckedListBox)
    
    # Create user button
    $Script:CreateUserButton = New-Object System.Windows.Forms.Button
    $Script:CreateUserButton.Text = "Create User"
    $Script:CreateUserButton.Size = New-Object System.Drawing.Size(120, 35)
    $Script:CreateUserButton.Location = New-Object System.Drawing.Point(20, 610)
    $Script:CreateUserButton.BackColor = [System.Drawing.Color]::LightGreen
    $Script:CreateUserButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $Script:CreateUserButton.Enabled = $false
    $Script:CreateUserButton.Add_Click({
        try {
            # Validate required fields
            if ([string]::IsNullOrWhiteSpace($Script:DisplayNameTextBox.Text)) {
                [System.Windows.Forms.MessageBox]::Show("Display Name is required", "Validation Error", "OK", "Warning")
                return
            }
            
            if ([string]::IsNullOrWhiteSpace($Script:UsernameTextBox.Text)) {
                [System.Windows.Forms.MessageBox]::Show("Username is required", "Validation Error", "OK", "Warning")
                return
            }
            
            if ([string]::IsNullOrWhiteSpace($Script:PasswordTextBox.Text)) {
                [System.Windows.Forms.MessageBox]::Show("Password is required", "Validation Error", "OK", "Warning")
                return
            }
            
            if (-not $Script:DomainDropdown.SelectedItem) {
                [System.Windows.Forms.MessageBox]::Show("Please select a domain", "Validation Error", "OK", "Warning")
                return
            }
            
            $UserPrincipalName = "$($Script:UsernameTextBox.Text)@$($Script:DomainDropdown.SelectedItem)"
            
            # Get selected groups
            $SelectedGroups = @()
            for ($i = 0; $i -lt $Script:GroupsCheckedListBox.CheckedItems.Count; $i++) {
                $SelectedGroups += $Script:GroupsCheckedListBox.CheckedItems[$i].ToString()
            }
            
            # Get manager
            $Manager = $null
            if ($Script:ManagerDropdown.SelectedIndex -gt 0) {
                $ManagerText = $Script:ManagerDropdown.SelectedItem
                if ($ManagerText -match '\(([^)]+)\)$') {
                    $Manager = $Matches[1]
                }
            }
            
            # Get license type
            $LicenseType = $null
            if ($Script:LicenseDropdown.SelectedIndex -gt 0) {
                $LicenseType = $Script:LicenseDropdown.SelectedItem
            }
            
            $Script:CreateUserButton.Enabled = $false
            $Script:CreateUserButton.Text = "Creating..."
            Update-StatusLabel "Creating user: $($Script:DisplayNameTextBox.Text)..."
            
            $NewUserResult = New-M365User -DisplayName $Script:DisplayNameTextBox.Text -UserPrincipalName $UserPrincipalName -FirstName $Script:FirstNameTextBox.Text -LastName $Script:LastNameTextBox.Text -Department $Script:DepartmentTextBox.Text -JobTitle $Script:JobTitleTextBox.Text -Office $Script:OfficeDropdown.SelectedItem -Manager $Manager -LicenseType $LicenseType -Groups $SelectedGroups -Password $Script:PasswordTextBox.Text
            
            [System.Windows.Forms.MessageBox]::Show("User created successfully!`n`nDisplay Name: $($NewUserResult.DisplayName)`nUPN: $($NewUserResult.UserPrincipalName)", "User Created", "OK", "Information")
            
            Clear-UserCreationForm
            Update-StatusLabel "User created successfully"
            
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to create user: $($_.Exception.Message)", "Error", "OK", "Error")
            Update-StatusLabel "User creation failed"
        }
        finally {
            $Script:CreateUserButton.Enabled = $true
            $Script:CreateUserButton.Text = "Create User"
        }
    })
    
    # Clear form button
    $ClearFormButton = New-Object System.Windows.Forms.Button
    $ClearFormButton.Text = "Clear Form"
    $ClearFormButton.Size = New-Object System.Drawing.Size(100, 35)
    $ClearFormButton.Location = New-Object System.Drawing.Point(150, 610)
    $ClearFormButton.Add_Click({
        Clear-UserCreationForm
    })
    
    $MainPanel.Controls.AddRange(@(
        $UserDetailsGroup,
        $ManagementGroup,
        $GroupsGroup,
        $Script:CreateUserButton,
        $ClearFormButton
    ))
    
    $Tab.Controls.Add($MainPanel)
    $null = $Script:TabControl.TabPages.Add($Tab)
}

function New-BulkImportTab {
    <#
    .SYNOPSIS
        Creates the bulk CSV import tab interface
    #>
    $Tab = New-Object System.Windows.Forms.TabPage
    $Tab.Text = "Bulk Import"
    $Tab.BackColor = [System.Drawing.Color]::White
    
    # Instructions
    $InstructionsGroup = New-Object System.Windows.Forms.GroupBox
    $InstructionsGroup.Text = "CSV Import Instructions"
    $InstructionsGroup.Location = New-Object System.Drawing.Point(10, 10)
    $InstructionsGroup.Size = New-Object System.Drawing.Size(900, 120)
    
    $InstructionsText = New-Object System.Windows.Forms.RichTextBox
    $InstructionsText.Dock = "Fill"
    $InstructionsText.ReadOnly = $true
    $InstructionsText.Text = "CSV Format Requirements:`n• Required columns: DisplayName, UserPrincipalName, Password`n• Optional columns: FirstName, LastName, Department, JobTitle, Office, Manager, LicenseType, Groups`n• Groups column should contain comma-separated group names`n• Manager should be the UPN of the manager`n• Office should match one of the UK locations`n• LicenseType should be one of: BusinessBasic, BusinessPremium, BusinessStandard"
    $InstructionsText.BackColor = [System.Drawing.Color]::LightYellow
    
    $InstructionsGroup.Controls.Add($InstructionsText)
    
    # File selection
    $FileSelectionGroup = New-Object System.Windows.Forms.GroupBox
    $FileSelectionGroup.Text = "File Selection"
    $FileSelectionGroup.Location = New-Object System.Drawing.Point(10, 140)
    $FileSelectionGroup.Size = New-Object System.Drawing.Size(900, 80)
    
    $FilePathLabel = New-Object System.Windows.Forms.Label
    $FilePathLabel.Text = "CSV File:"
    $FilePathLabel.Location = New-Object System.Drawing.Point(10, 30)
    $FilePathLabel.Size = New-Object System.Drawing.Size(60, 20)
    
    $Script:FilePathTextBox = New-Object System.Windows.Forms.TextBox
    $Script:FilePathTextBox.Location = New-Object System.Drawing.Point(80, 28)
    $Script:FilePathTextBox.Size = New-Object System.Drawing.Size(600, 20)
    $Script:FilePathTextBox.ReadOnly = $true
    
    $BrowseButton = New-Object System.Windows.Forms.Button
    $BrowseButton.Text = "Browse..."
    $BrowseButton.Location = New-Object System.Drawing.Point(690, 28)
    $BrowseButton.Size = New-Object System.Drawing.Size(80, 23)
    $BrowseButton.Add_Click({
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $OpenFileDialog.FilterIndex = 1
        
        if ($OpenFileDialog.ShowDialog() -eq "OK") {
            $Script:FilePathTextBox.Text = $OpenFileDialog.FileName
        }
    })
    
    $FileSelectionGroup.Controls.AddRange(@(
        $FilePathLabel, $Script:FilePathTextBox, $BrowseButton
    ))
    
    # Progress
    $ProgressGroup = New-Object System.Windows.Forms.GroupBox
    $ProgressGroup.Text = "Import Progress"
    $ProgressGroup.Location = New-Object System.Drawing.Point(10, 230)
    $ProgressGroup.Size = New-Object System.Drawing.Size(900, 100)
    
    $Script:ProgressBar = New-Object System.Windows.Forms.ProgressBar
    $Script:ProgressBar.Location = New-Object System.Drawing.Point(10, 30)
    $Script:ProgressBar.Size = New-Object System.Drawing.Size(880, 25)
    
    $Script:ProgressLabel = New-Object System.Windows.Forms.Label
    $Script:ProgressLabel.Text = "Ready to import"
    $Script:ProgressLabel.Location = New-Object System.Drawing.Point(10, 60)
    $Script:ProgressLabel.Size = New-Object System.Drawing.Size(880, 20)
    
    $ProgressGroup.Controls.AddRange(@(
        $Script:ProgressBar, $Script:ProgressLabel
    ))
    
    # Import button
    $Script:ImportCSVButton = New-Object System.Windows.Forms.Button
    $Script:ImportCSVButton.Text = "Import Users from CSV"
    $Script:ImportCSVButton.Size = New-Object System.Drawing.Size(150, 35)
    $Script:ImportCSVButton.Location = New-Object System.Drawing.Point(10, 340)
    $Script:ImportCSVButton.BackColor = [System.Drawing.Color]::LightBlue
    $Script:ImportCSVButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $Script:ImportCSVButton.Enabled = $false
    $Script:ImportCSVButton.Add_Click({
        if ([string]::IsNullOrWhiteSpace($Script:FilePathTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please select a CSV file first", "No File Selected", "OK", "Warning")
            return
        }
        
        if (-not (Test-Path $Script:FilePathTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("The selected file does not exist", "File Not Found", "OK", "Error")
            return
        }
        
        $Confirmation = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to import users from the selected CSV file?", "Confirm Import", "YesNo", "Question")
        
        if ($Confirmation -eq "Yes") {
            try {
                $Script:ImportCSVButton.Enabled = $false
                $Script:ImportCSVButton.Text = "Importing..."
                $Script:ProgressBar.Value = 0
                $Script:ProgressLabel.Text = "Starting import..."
                
                $Result = Import-UsersFromCSV -CSVPath $Script:FilePathTextBox.Text
                
                [System.Windows.Forms.MessageBox]::Show("Import completed!`n`nTotal: $($Result.TotalUsers)`nSuccess: $($Result.Successful)`nFailed: $($Result.Failed)", "Import Complete", "OK", "Information")
                
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Import failed: $($_.Exception.Message)", "Import Error", "OK", "Error")
            }
            finally {
                $Script:ImportCSVButton.Enabled = $true
                $Script:ImportCSVButton.Text = "Import Users from CSV"
                $Script:ProgressBar.Value = 0
                $Script:ProgressLabel.Text = "Import completed"
            }
        }
    })
    
    $Tab.Controls.AddRange(@(
        $InstructionsGroup,
        $FileSelectionGroup,
        $ProgressGroup,
        $Script:ImportCSVButton
    ))
    
    $null = $Script:TabControl.TabPages.Add($Tab)
}

function New-TenantDataTab {
    <#
    .SYNOPSIS
        Creates the tenant data viewing tab
    #>
    $Tab = New-Object System.Windows.Forms.TabPage
    $Tab.Text = "Tenant Data"
    $Tab.BackColor = [System.Drawing.Color]::White
    
    $SubTabControl = New-Object System.Windows.Forms.TabControl
    $SubTabControl.Dock = "Fill"
    
    # Users sub-tab
    $UsersTab = New-Object System.Windows.Forms.TabPage
    $UsersTab.Text = "Users"
    
    $Script:UsersDataGridView = New-Object System.Windows.Forms.DataGridView
    $Script:UsersDataGridView.Dock = "Fill"
    $Script:UsersDataGridView.ReadOnly = $true
    $Script:UsersDataGridView.AutoSizeColumnsMode = "AllCells"
    $Script:UsersDataGridView.AllowUserToAddRows = $false
    
    $UsersTab.Controls.Add($Script:UsersDataGridView)
    
    # Groups sub-tab
    $GroupsTab = New-Object System.Windows.Forms.TabPage
    $GroupsTab.Text = "Groups"
    
    $Script:GroupsDataGridView = New-Object System.Windows.Forms.DataGridView
    $Script:GroupsDataGridView.Dock = "Fill"
    $Script:GroupsDataGridView.ReadOnly = $true
    $Script:GroupsDataGridView.AutoSizeColumnsMode = "AllCells"
    $Script:GroupsDataGridView.AllowUserToAddRows = $false
    
    $GroupsTab.Controls.Add($Script:GroupsDataGridView)
    
    # Licenses sub-tab
    $LicensesTab = New-Object System.Windows.Forms.TabPage
    $LicensesTab.Text = "Licenses"
    
    $Script:LicensesDataGridView = New-Object System.Windows.Forms.DataGridView
    $Script:LicensesDataGridView.Dock = "Fill"
    $Script:LicensesDataGridView.ReadOnly = $true
    $Script:LicensesDataGridView.AutoSizeColumnsMode = "AllCells"
    $Script:LicensesDataGridView.AllowUserToAddRows = $false
    
    $LicensesTab.Controls.Add($Script:LicensesDataGridView)
    
    $SubTabControl.TabPages.AddRange(@($UsersTab, $GroupsTab, $LicensesTab))
    $Tab.Controls.Add($SubTabControl)
    $null = $Script:TabControl.TabPages.Add($Tab)
}

function New-ActivityLogTab {
    <#
    .SYNOPSIS
        Creates the activity log tab
    #>
    $Tab = New-Object System.Windows.Forms.TabPage
    $Tab.Text = "Activity Log"
    $Tab.BackColor = [System.Drawing.Color]::White
    
    $Script:LogTextBox = New-Object System.Windows.Forms.RichTextBox
    $Script:LogTextBox.Dock = "Fill"
    $Script:LogTextBox.ReadOnly = $true
    $Script:LogTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $Script:LogTextBox.BackColor = [System.Drawing.Color]::Black
    $Script:LogTextBox.ForeColor = [System.Drawing.Color]::White
    
    $Tab.Controls.Add($Script:LogTextBox)
    $null = $Script:TabControl.TabPages.Add($Tab)
}

function Update-StatusLabel {
    param([string]$Message)
    
    if ($Script:StatusLabel) {
        $Script:StatusLabel.Text = $Message
        if ($Script:StatusStrip) {
            $Script:StatusStrip.Refresh()
        }
    }
}

function Update-UIAfterConnection {
    if ($Script:ConnectButton) { 
        $Script:ConnectButton.Enabled = $false 
        $Script:ConnectButton.Text = "Connected"
        $Script:ConnectButton.BackColor = [System.Drawing.Color]::LightGray
    }
    if ($Script:DisconnectButton) { $Script:DisconnectButton.Enabled = $true }
    if ($Script:RefreshButton) { $Script:RefreshButton.Enabled = $true }
    if ($Script:CreateUserButton) { $Script:CreateUserButton.Enabled = $true }
    if ($Script:ImportCSVButton) { $Script:ImportCSVButton.Enabled = $true }
    
    if ($Script:ConnectionInfoLabel) {
        $Script:ConnectionInfoLabel.Text = "Status: Connected to Microsoft Graph"
        $Script:ConnectionInfoLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    
    # Update dropdowns and lists with tenant data
    Update-DomainDropdown
    Update-ManagerDropdown
    Update-GroupsList
    Update-LicenseDropdown
    Refresh-TenantDataViews
}

function Update-UIAfterDisconnection {
    if ($Script:ConnectButton) { 
        $Script:ConnectButton.Enabled = $true 
        $Script:ConnectButton.Text = "Connect to M365"
        $Script:ConnectButton.BackColor = [System.Drawing.Color]::LightGreen
    }
    if ($Script:DisconnectButton) { $Script:DisconnectButton.Enabled = $false }
    if ($Script:RefreshButton) { $Script:RefreshButton.Enabled = $false }
    if ($Script:CreateUserButton) { $Script:CreateUserButton.Enabled = $false }
    if ($Script:ImportCSVButton) { $Script:ImportCSVButton.Enabled = $false }
    
    if ($Script:ConnectionInfoLabel) {
        $Script:ConnectionInfoLabel.Text = "Status: Disconnected"
        $Script:ConnectionInfoLabel.ForeColor = [System.Drawing.Color]::DarkBlue
    }
    
    Clear-AllDropdowns
}

function Update-DomainDropdown {
    if ($Script:DomainDropdown) {
        $Script:DomainDropdown.Items.Clear()
        $TenantData = Get-M365TenantData
        foreach ($Domain in $TenantData.AcceptedDomains) {
            $null = $Script:DomainDropdown.Items.Add($Domain.DomainName)
            if ($Domain.IsDefault) {
                $Script:DomainDropdown.SelectedItem = $Domain.DomainName
            }
        }
    }
}

function Update-ManagerDropdown {
    if ($Script:ManagerDropdown) {
        $Script:ManagerDropdown.Items.Clear()
        $null = $Script:ManagerDropdown.Items.Add("(No Manager)")
        $TenantData = Get-M365TenantData
        foreach ($User in ($TenantData.AvailableUsers | Sort-Object DisplayName)) {
            $DisplayText = "$($User.DisplayName) ($($User.UserPrincipalName))"
            $null = $Script:ManagerDropdown.Items.Add($DisplayText)
        }
        $Script:ManagerDropdown.SelectedIndex = 0
    }
}

function Update-GroupsList {
    if ($Script:GroupsCheckedListBox) {
        $Script:GroupsCheckedListBox.Items.Clear()
        $TenantData = Get-M365TenantData
        foreach ($Group in ($TenantData.AvailableGroups | Sort-Object GroupType, DisplayName)) {
            $DisplayText = "$($Group.DisplayName) [$($Group.GroupType)]"
            if ($Group.Mail) {
                $DisplayText += " - $($Group.Mail)"
            }
            $null = $Script:GroupsCheckedListBox.Items.Add($DisplayText)
        }
    }
}

function Update-LicenseDropdown {
    if ($Script:LicenseDropdown) {
        $Script:LicenseDropdown.Items.Clear()
        $null = $Script:LicenseDropdown.Items.Add("(No License Assignment)")
        $LicenseTypes = @("BusinessBasic", "BusinessPremium", "BusinessStandard", "ExchangeOnline1", "ExchangeOnline2")
        foreach ($LicenseType in $LicenseTypes) {
            $null = $Script:LicenseDropdown.Items.Add($LicenseType)
        }
        $Script:LicenseDropdown.SelectedIndex = 0
    }
}

function Clear-AllDropdowns {
    if ($Script:DomainDropdown) { $Script:DomainDropdown.Items.Clear() }
    if ($Script:ManagerDropdown) { $Script:ManagerDropdown.Items.Clear() }
    if ($Script:GroupsCheckedListBox) { $Script:GroupsCheckedListBox.Items.Clear() }
    if ($Script:LicenseDropdown) { $Script:LicenseDropdown.Items.Clear() }
}

function Refresh-TenantDataViews {
    if ($Script:UsersDataGridView) {
        $TenantData = Get-M365TenantData
        $Script:UsersDataGridView.DataSource = $TenantData.AvailableUsers
    }
    
    if ($Script:GroupsDataGridView) {
        $TenantData = Get-M365TenantData
        $Script:GroupsDataGridView.DataSource = $TenantData.AvailableGroups
    }
    
    if ($Script:LicensesDataGridView) {
        $TenantData = Get-M365TenantData
        $Script:LicensesDataGridView.DataSource = $TenantData.AvailableLicenses
    }
}

function Clear-UserCreationForm {
    if ($Script:DisplayNameTextBox) { $Script:DisplayNameTextBox.Clear() }
    if ($Script:FirstNameTextBox) { $Script:FirstNameTextBox.Clear() }
    if ($Script:LastNameTextBox) { $Script:LastNameTextBox.Clear() }
    if ($Script:UsernameTextBox) { $Script:UsernameTextBox.Clear() }
    if ($Script:PasswordTextBox) { $Script:PasswordTextBox.Clear() }
    if ($Script:DepartmentTextBox) { $Script:DepartmentTextBox.Clear() }
    if ($Script:JobTitleTextBox) { $Script:JobTitleTextBox.Clear() }
    
    if ($Script:OfficeDropdown -and $Script:OfficeDropdown.Items.Count -gt 0) {
        $Script:OfficeDropdown.SelectedIndex = -1
    }
    
    if ($Script:ManagerDropdown -and $Script:ManagerDropdown.Items.Count -gt 0) {
        $Script:ManagerDropdown.SelectedIndex = 0
    }
    
    if ($Script:LicenseDropdown -and $Script:LicenseDropdown.Items.Count -gt 0) {
        $Script:LicenseDropdown.SelectedIndex = 0
    }
    
    if ($Script:GroupsCheckedListBox) {
        for ($i = 0; $i -lt $Script:GroupsCheckedListBox.Items.Count; $i++) {
            $Script:GroupsCheckedListBox.SetItemChecked($i, $false)
        }
    }
}

# Export all GUI functions
Export-ModuleMember -Function @(
    'Start-M365ProvisioningTool',
    'New-MainForm',
    'New-UserCreationTab',
    'New-BulkImportTab',
    'New-TenantDataTab',
    'New-ActivityLogTab',
    'Update-StatusLabel',
    'Update-UIAfterConnection',
    'Update-UIAfterDisconnection',
    'Update-DomainDropdown',
    'Update-ManagerDropdown',
    'Update-GroupsList',
    'Update-LicenseDropdown',
    'Clear-AllDropdowns',
    'Refresh-TenantDataViews',
    'Clear-UserCreationForm'
)


