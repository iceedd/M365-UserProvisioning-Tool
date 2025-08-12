#Requires -Version 7.0

<#
.SYNOPSIS
    M365 GUI Module - Complete Enterprise Edition
    
.DESCRIPTION
    Full-featured Windows Forms interface for M365 User Provisioning Tool with:
    - Comprehensive tenant data discovery
    - Complete user creation form with all fields
    - License assignment via CustomAttribute1
    - Manual office location input
    - Manager and group selection
    - Real-time tenant data updates
    
.NOTES
    Module: M365.GUI
    Version: 2.0.0-Enterprise
    Author: Tom Mortiboys
    
    FIXED: No SetCompatibleTextRenderingDefault - resolves initialization errors
#>

# Script-level variables for form controls
$Script:MainForm = $null
$Script:StatusLabel = $null
$Script:StatusStrip = $null
$Script:TabControl = $null

# Connection controls
$Script:ConnectButton = $null
$Script:DisconnectButton = $null
$Script:RefreshDataButton = $null

# User creation form controls - COMPLETE SET
$Script:FirstNameTextBox = $null
$Script:LastNameTextBox = $null
$Script:DisplayNameTextBox = $null
$Script:UsernameTextBox = $null
$Script:PasswordTextBox = $null
$Script:DepartmentTextBox = $null
$Script:JobTitleTextBox = $null
$Script:OfficePhoneTextBox = $null
$Script:MobilePhoneTextBox = $null
$Script:StreetAddressTextBox = $null
$Script:CityTextBox = $null
$Script:PostalCodeTextBox = $null

# Dropdown controls
$Script:DomainDropdown = $null
$Script:OfficeLocationTextBox = $null  # CHANGED: Now a TextBox instead of Dropdown
$Script:ManagerDropdown = $null
$Script:LicenseTypeDropdown = $null
$Script:CountryDropdown = $null
$Script:StateDropdown = $null

# Group selection
$Script:GroupsCheckedListBox = $null

# Tenant data controls
$Script:TenantDataTextBox = $null
$Script:UsersListBox = $null
$Script:GroupsListBox = $null
$Script:LicensesListBox = $null
$Script:DomainsListBox = $null

# Activity log
$Script:ActivityLogTextBox = $null

# Global tenant data (populated after connection)
$Script:TenantUsers = @()
$Script:TenantGroups = @()
$Script:TenantLicenses = @()
$Script:TenantDomains = @()
$Script:TenantInfo = $null

# License types for CustomAttribute1
$Script:LicenseTypes = @(
    "BusinessBasic",
    "BusinessPremium", 
    "BusinessStandard",
    "E3",
    "E5",
    "ExchangeOnline1",
    "ExchangeOnline2"
)

# Track initialization state
$Script:WindowsFormsInitialized = $false
$Script:TenantDataLoaded = $false

function Initialize-WindowsFormsIfNeeded {
    <#
    .SYNOPSIS
        Initializes Windows Forms if not already done - ENTERPRISE FIXED VERSION
    #>
    
    if ($Script:WindowsFormsInitialized) {
        return $true
    }
    
    try {
        Write-Verbose "Initializing Windows Forms (Enterprise Fixed Mode)..."
        
        # Load assemblies
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        
        # Enable visual styles only - SKIP SetCompatibleTextRenderingDefault
        [System.Windows.Forms.Application]::EnableVisualStyles()
        
        # Mark as initialized
        $Script:WindowsFormsInitialized = $true
        
        Write-Verbose "Windows Forms initialized successfully (Enterprise Mode)"
        return $true
    }
    catch {
        Write-Error "Failed to initialize Windows Forms: $($_.Exception.Message)"
        return $false
    }
}

function Start-M365ProvisioningTool {
    <#
    .SYNOPSIS
        Main application launcher - Enterprise Edition
    #>
    
    try {
        Write-Host "🚀 Starting M365 User Provisioning Tool GUI (Enterprise)..." -ForegroundColor Cyan
        
        # Initialize Windows Forms first
        if (-not (Initialize-WindowsFormsIfNeeded)) {
            throw "Failed to initialize Windows Forms"
        }
        
        # Create and show main form
        $Script:MainForm = New-MainForm
        
        Write-Host "🖥️ Launching Enterprise GUI interface..." -ForegroundColor Green
        
        # Show the form and wait for it to close
        $Result = $Script:MainForm.ShowDialog()
        
        Write-Host "📱 Enterprise application closed" -ForegroundColor Gray
        return $Result
        
    }
    catch {
        $ErrorMessage = "Enterprise GUI failed to start: $($_.Exception.Message)"
        Write-Error $ErrorMessage
        
        if ($Script:WindowsFormsInitialized) {
            try {
                [System.Windows.Forms.MessageBox]::Show(
                    $ErrorMessage,
                    "Enterprise Application Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
            catch {
                Write-Host "Error: $ErrorMessage" -ForegroundColor Red
            }
        }
        
        throw
    }
}

function New-MainForm {
    <#
    .SYNOPSIS
        Creates the main enterprise application form
    #>
    
    if (-not (Initialize-WindowsFormsIfNeeded)) {
        throw "Windows Forms initialization failed"
    }
    
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "M365 User Provisioning Tool - Enterprise Edition 2025"
    $Form.Size = New-Object System.Drawing.Size(1600, 1000)
    $Form.StartPosition = "CenterScreen"
    $Form.MinimumSize = New-Object System.Drawing.Size(1400, 900)
    $Form.MaximizeBox = $true
    $Form.WindowState = "Maximized"
    $Form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    # Set application icon
    try {
        $Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("$env:SystemRoot\System32\shell32.dll")
    }
    catch {
        Write-Verbose "Could not set application icon"
    }
    
    # Status strip at bottom
    $Script:StatusStrip = New-Object System.Windows.Forms.StatusStrip
    $Script:StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $Script:StatusLabel.Text = "Ready - Click Connect to begin tenant discovery"
    $Script:StatusLabel.Spring = $true
    $Script:StatusLabel.TextAlign = "MiddleLeft"
    $Script:StatusStrip.Items.Add($Script:StatusLabel) | Out-Null
    
    # Create tab control
    $Script:TabControl = New-Object System.Windows.Forms.TabControl
    $Script:TabControl.Dock = "Fill"
    $Script:TabControl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    # Create tabs
    $UserCreationTab = New-UserCreationTab
    $BulkImportTab = New-BulkImportTab
    $TenantDataTab = New-TenantDataTab
    $ActivityLogTab = New-ActivityLogTab
    
    # Add tabs to control
    $Script:TabControl.TabPages.Add($UserCreationTab)
    $Script:TabControl.TabPages.Add($BulkImportTab) 
    $Script:TabControl.TabPages.Add($TenantDataTab)
    $Script:TabControl.TabPages.Add($ActivityLogTab)
    
    # Add controls to form
    $Form.Controls.Add($Script:TabControl)
    $Form.Controls.Add($Script:StatusStrip)
    
    # Form event handlers
    $Form.Add_Load({
        Write-Host "📋 Enterprise form loaded successfully" -ForegroundColor Green
        Update-StatusLabel "Enterprise application started - Ready for M365 connection"
        Add-ActivityLog "Application started - Enterprise Edition"
    })
    
    $Form.Add_FormClosing({
        param($sender, $e)
        
        if ($Global:AppState -and $Global:AppState.Connected) {
            $Result = [System.Windows.Forms.MessageBox]::Show(
                "You are currently connected to Microsoft 365. Disconnect and exit?",
                "Confirm Exit",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            
            if ($Result -eq [System.Windows.Forms.DialogResult]::No) {
                $e.Cancel = $true
                return
            }
            
            # Disconnect if connected
            try {
                if (Get-Command "Disconnect-FromMicrosoftGraph" -ErrorAction SilentlyContinue) {
                    Disconnect-FromMicrosoftGraph
                }
            }
            catch {
                Write-Verbose "Error during disconnect: $($_.Exception.Message)"
            }
        }
        
        Add-ActivityLog "Application closing"
        Write-Host "👋 Enterprise application shutting down..." -ForegroundColor Yellow
    })
    
    return $Form
}

function New-UserCreationTab {
    <#
    .SYNOPSIS
        Creates the comprehensive user creation tab
    #>
    
    $Tab = New-Object System.Windows.Forms.TabPage
    $Tab.Text = "👤 Create User"
    $Tab.Name = "UserCreationTab"
    $Tab.Padding = New-Object System.Windows.Forms.Padding(10)
    $Tab.AutoScroll = $true
    
    # Connection Panel (Top)
    $ConnectionPanel = New-Object System.Windows.Forms.Panel
    $ConnectionPanel.Height = 70
    $ConnectionPanel.Dock = "Top"
    $ConnectionPanel.BackColor = [System.Drawing.Color]::LightSteelBlue
    $ConnectionPanel.Padding = New-Object System.Windows.Forms.Padding(10)
    
    # Connection buttons
    $Script:ConnectButton = New-Object System.Windows.Forms.Button
    $Script:ConnectButton.Text = "🔗 Connect to M365"
    $Script:ConnectButton.Size = New-Object System.Drawing.Size(180, 35)
    $Script:ConnectButton.Location = New-Object System.Drawing.Point(10, 15)
    $Script:ConnectButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $Script:ConnectButton.BackColor = [System.Drawing.Color]::LightBlue
    
    $Script:RefreshDataButton = New-Object System.Windows.Forms.Button
    $Script:RefreshDataButton.Text = "🔄 Refresh Tenant Data"
    $Script:RefreshDataButton.Size = New-Object System.Drawing.Size(180, 35)
    $Script:RefreshDataButton.Location = New-Object System.Drawing.Point(200, 15)
    $Script:RefreshDataButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $Script:RefreshDataButton.Enabled = $false
    
    $Script:DisconnectButton = New-Object System.Windows.Forms.Button
    $Script:DisconnectButton.Text = "❌ Disconnect"
    $Script:DisconnectButton.Size = New-Object System.Drawing.Size(120, 35)
    $Script:DisconnectButton.Location = New-Object System.Drawing.Point(390, 15)
    $Script:DisconnectButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $Script:DisconnectButton.Enabled = $false
    
    # Connection event handlers
    $Script:ConnectButton.Add_Click({
        try {
            Update-StatusLabel "🔗 Connecting to Microsoft 365..."
            
            # Call authentication function
            if (Get-Command "Connect-ToMicrosoftGraph" -ErrorAction SilentlyContinue) {
                $Connected = Connect-ToMicrosoftGraph
                if ($Connected) {
                    Update-UIAfterConnection
                    Start-TenantDataDiscovery
                }
            } else {
                throw "Connect-ToMicrosoftGraph function not available. Check M365.Authentication module."
            }
        }
        catch {
            Update-StatusLabel "❌ Connection failed: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to connect to Microsoft 365:`n`n$($_.Exception.Message)",
                "Connection Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    })
    
    $Script:RefreshDataButton.Add_Click({
        Start-TenantDataDiscovery
    })
    
    $Script:DisconnectButton.Add_Click({
        try {
            if (Get-Command "Disconnect-FromMicrosoftGraph" -ErrorAction SilentlyContinue) {
                Disconnect-FromMicrosoftGraph
            }
            Update-UIAfterDisconnection
        }
        catch {
            Write-Warning "Disconnect failed: $($_.Exception.Message)"
        }
    })
    
    $ConnectionPanel.Controls.AddRange(@(
        $Script:ConnectButton, $Script:RefreshDataButton, $Script:DisconnectButton
    ))
    
    # Main Content Panel
    $ContentPanel = New-Object System.Windows.Forms.Panel
    $ContentPanel.Dock = "Fill"
    $ContentPanel.AutoScroll = $true
    $ContentPanel.Padding = New-Object System.Windows.Forms.Padding(10)
    
    # User Details Group (Left Side)
    $UserDetailsGroup = New-Object System.Windows.Forms.GroupBox
    $UserDetailsGroup.Text = "User Details"
    $UserDetailsGroup.Location = New-Object System.Drawing.Point(10, 10)
    $UserDetailsGroup.Size = New-Object System.Drawing.Size(480, 450)
    $UserDetailsGroup.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    
    # Create comprehensive user form
    $y = 25
    $spacing = 30
    $labelWidth = 100
    $textBoxWidth = 200
    
    # First Name
    $FirstNameLabel = New-Object System.Windows.Forms.Label
    $FirstNameLabel.Text = "First Name: *"
    $FirstNameLabel.Location = New-Object System.Drawing.Point(10, $y)
    $FirstNameLabel.Size = New-Object System.Drawing.Size($labelWidth, 20)
    
    $Script:FirstNameTextBox = New-Object System.Windows.Forms.TextBox
    $Script:FirstNameTextBox.Location = New-Object System.Drawing.Point(120, ($y-2))
    $Script:FirstNameTextBox.Size = New-Object System.Drawing.Size($textBoxWidth, 20)
    
    $y += $spacing
    
    # Last Name
    $LastNameLabel = New-Object System.Windows.Forms.Label
    $LastNameLabel.Text = "Last Name: *"
    $LastNameLabel.Location = New-Object System.Drawing.Point(10, $y)
    $LastNameLabel.Size = New-Object System.Drawing.Size($labelWidth, 20)
    
    $Script:LastNameTextBox = New-Object System.Windows.Forms.TextBox
    $Script:LastNameTextBox.Location = New-Object System.Drawing.Point(120, ($y-2))
    $Script:LastNameTextBox.Size = New-Object System.Drawing.Size($textBoxWidth, 20)
    
    $y += $spacing
    
    # Auto-generate display name when first/last name changes
    $Script:FirstNameTextBox.Add_TextChanged({
        Update-DisplayName
    })
    
    $Script:LastNameTextBox.Add_TextChanged({
        Update-DisplayName
    })
    
    # Display Name
    $DisplayNameLabel = New-Object System.Windows.Forms.Label
    $DisplayNameLabel.Text = "Display Name:"
    $DisplayNameLabel.Location = New-Object System.Drawing.Point(10, $y)
    $DisplayNameLabel.Size = New-Object System.Drawing.Size($labelWidth, 20)
    
    $Script:DisplayNameTextBox = New-Object System.Windows.Forms.TextBox
    $Script:DisplayNameTextBox.Location = New-Object System.Drawing.Point(120, ($y-2))
    $Script:DisplayNameTextBox.Size = New-Object System.Drawing.Size($textBoxWidth, 20)
    
    $y += $spacing
    
    # Username
    $UsernameLabel = New-Object System.Windows.Forms.Label
    $UsernameLabel.Text = "Username: *"
    $UsernameLabel.Location = New-Object System.Drawing.Point(10, $y)
    $UsernameLabel.Size = New-Object System.Drawing.Size($labelWidth, 20)
    
    $Script:UsernameTextBox = New-Object System.Windows.Forms.TextBox
    $Script:UsernameTextBox.Location = New-Object System.Drawing.Point(120, ($y-2))
    $Script:UsernameTextBox.Size = New-Object System.Drawing.Size(150, 20)
    
    # Domain dropdown next to username
    $AtLabel = New-Object System.Windows.Forms.Label
    $AtLabel.Text = "@"
    $AtLabel.Location = New-Object System.Drawing.Point(275, $y)
    $AtLabel.Size = New-Object System.Drawing.Size(15, 20)
    
    $Script:DomainDropdown = New-Object System.Windows.Forms.ComboBox
    $Script:DomainDropdown.Location = New-Object System.Drawing.Point(290, ($y-2))
    $Script:DomainDropdown.Size = New-Object System.Drawing.Size(150, 20)
    $Script:DomainDropdown.DropDownStyle = "DropDownList"
    
    $y += $spacing
    
    # Password
    $PasswordLabel = New-Object System.Windows.Forms.Label
    $PasswordLabel.Text = "Password: *"
    $PasswordLabel.Location = New-Object System.Drawing.Point(10, $y)
    $PasswordLabel.Size = New-Object System.Drawing.Size($labelWidth, 20)
    
    $Script:PasswordTextBox = New-Object System.Windows.Forms.TextBox
    $Script:PasswordTextBox.Location = New-Object System.Drawing.Point(120, ($y-2))
    $Script:PasswordTextBox.Size = New-Object System.Drawing.Size(150, 20)
    $Script:PasswordTextBox.UseSystemPasswordChar = $true
    
    $GeneratePasswordButton = New-Object System.Windows.Forms.Button
    $GeneratePasswordButton.Text = "Generate"
    $GeneratePasswordButton.Location = New-Object System.Drawing.Point(280, ($y-3))
    $GeneratePasswordButton.Size = New-Object System.Drawing.Size(80, 22)
    $GeneratePasswordButton.Add_Click({
        $Script:PasswordTextBox.Text = Generate-SecurePassword
    })
    
    $y += $spacing
    
    # Department
    $DepartmentLabel = New-Object System.Windows.Forms.Label
    $DepartmentLabel.Text = "Department:"
    $DepartmentLabel.Location = New-Object System.Drawing.Point(10, $y)
    $DepartmentLabel.Size = New-Object System.Drawing.Size($labelWidth, 20)
    
    $Script:DepartmentTextBox = New-Object System.Windows.Forms.TextBox
    $Script:DepartmentTextBox.Location = New-Object System.Drawing.Point(120, ($y-2))
    $Script:DepartmentTextBox.Size = New-Object System.Drawing.Size($textBoxWidth, 20)
    
    $y += $spacing
    
    # Job Title
    $JobTitleLabel = New-Object System.Windows.Forms.Label
    $JobTitleLabel.Text = "Job Title:"
    $JobTitleLabel.Location = New-Object System.Drawing.Point(10, $y)
    $JobTitleLabel.Size = New-Object System.Drawing.Size($labelWidth, 20)
    
    $Script:JobTitleTextBox = New-Object System.Windows.Forms.TextBox
    $Script:JobTitleTextBox.Location = New-Object System.Drawing.Point(120, ($y-2))
    $Script:JobTitleTextBox.Size = New-Object System.Drawing.Size($textBoxWidth, 20)
    
    $y += $spacing
    
    # Office Location - CHANGED TO TEXT BOX
    $OfficeLocationLabel = New-Object System.Windows.Forms.Label
    $OfficeLocationLabel.Text = "Office Location:"
    $OfficeLocationLabel.Location = New-Object System.Drawing.Point(10, $y)
    $OfficeLocationLabel.Size = New-Object System.Drawing.Size($labelWidth, 20)
    
    $Script:OfficeLocationTextBox = New-Object System.Windows.Forms.TextBox
    $Script:OfficeLocationTextBox.Location = New-Object System.Drawing.Point(120, ($y-2))
    $Script:OfficeLocationTextBox.Size = New-Object System.Drawing.Size($textBoxWidth, 20)
    $Script:OfficeLocationTextBox.PlaceholderText = "Enter office location"
    
    $y += $spacing
    
    # Office Phone
    $OfficePhoneLabel = New-Object System.Windows.Forms.Label
    $OfficePhoneLabel.Text = "Office Phone:"
    $OfficePhoneLabel.Location = New-Object System.Drawing.Point(10, $y)
    $OfficePhoneLabel.Size = New-Object System.Drawing.Size($labelWidth, 20)
    
    $Script:OfficePhoneTextBox = New-Object System.Windows.Forms.TextBox
    $Script:OfficePhoneTextBox.Location = New-Object System.Drawing.Point(120, ($y-2))
    $Script:OfficePhoneTextBox.Size = New-Object System.Drawing.Size($textBoxWidth, 20)
    
    $y += $spacing
    
    # Mobile Phone
    $MobilePhoneLabel = New-Object System.Windows.Forms.Label
    $MobilePhoneLabel.Text = "Mobile Phone:"
    $MobilePhoneLabel.Location = New-Object System.Drawing.Point(10, $y)
    $MobilePhoneLabel.Size = New-Object System.Drawing.Size($labelWidth, 20)
    
    $Script:MobilePhoneTextBox = New-Object System.Windows.Forms.TextBox
    $Script:MobilePhoneTextBox.Location = New-Object System.Drawing.Point(120, ($y-2))
    $Script:MobilePhoneTextBox.Size = New-Object System.Drawing.Size($textBoxWidth, 20)
    
    # Add all controls to user details group
    $UserDetailsGroup.Controls.AddRange(@(
        $FirstNameLabel, $Script:FirstNameTextBox,
        $LastNameLabel, $Script:LastNameTextBox,
        $DisplayNameLabel, $Script:DisplayNameTextBox,
        $UsernameLabel, $Script:UsernameTextBox, $AtLabel, $Script:DomainDropdown,
        $PasswordLabel, $Script:PasswordTextBox, $GeneratePasswordButton,
        $DepartmentLabel, $Script:DepartmentTextBox,
        $JobTitleLabel, $Script:JobTitleTextBox,
        $OfficeLocationLabel, $Script:OfficeLocationTextBox,
        $OfficePhoneLabel, $Script:OfficePhoneTextBox,
        $MobilePhoneLabel, $Script:MobilePhoneTextBox
    ))
    
    # Management & Licensing Group (Right Side)
    $ManagementGroup = New-Object System.Windows.Forms.GroupBox
    $ManagementGroup.Text = "Management & Licensing"
    $ManagementGroup.Location = New-Object System.Drawing.Point(500, 10)
    $ManagementGroup.Size = New-Object System.Drawing.Size(480, 450)
    $ManagementGroup.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    
    $y = 25
    
    # Manager
    $ManagerLabel = New-Object System.Windows.Forms.Label
    $ManagerLabel.Text = "Manager:"
    $ManagerLabel.Location = New-Object System.Drawing.Point(10, $y)
    $ManagerLabel.Size = New-Object System.Drawing.Size($labelWidth, 20)
    
    $Script:ManagerDropdown = New-Object System.Windows.Forms.ComboBox
    $Script:ManagerDropdown.Location = New-Object System.Drawing.Point(120, ($y-2))
    $Script:ManagerDropdown.Size = New-Object System.Drawing.Size(340, 20)
    $Script:ManagerDropdown.DropDownStyle = "DropDownList"
    
    $y += $spacing
    
    # License Type (CustomAttribute1)
    $LicenseTypeLabel = New-Object System.Windows.Forms.Label
    $LicenseTypeLabel.Text = "License Type:"
    $LicenseTypeLabel.Location = New-Object System.Drawing.Point(10, $y)
    $LicenseTypeLabel.Size = New-Object System.Drawing.Size($labelWidth, 20)
    
    $Script:LicenseTypeDropdown = New-Object System.Windows.Forms.ComboBox
    $Script:LicenseTypeDropdown.Location = New-Object System.Drawing.Point(120, ($y-2))
    $Script:LicenseTypeDropdown.Size = New-Object System.Drawing.Size(200, 20)
    $Script:LicenseTypeDropdown.DropDownStyle = "DropDownList"
    
    # Populate license types
    foreach ($LicenseType in $Script:LicenseTypes) {
        $Script:LicenseTypeDropdown.Items.Add($LicenseType) | Out-Null
    }
    
    $y += $spacing
    
    # License info
    $LicenseInfoLabel = New-Object System.Windows.Forms.Label
    $LicenseInfoLabel.Text = "Note: License assignment handled via CustomAttribute1`nThis value determines automatic license assignment"
    $LicenseInfoLabel.Location = New-Object System.Drawing.Point(10, $y)
    $LicenseInfoLabel.Size = New-Object System.Drawing.Size(450, 40)
    $LicenseInfoLabel.ForeColor = [System.Drawing.Color]::DarkBlue
    $LicenseInfoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    
    $y += 50
    
    # Available license types display
    $AvailableLicensesLabel = New-Object System.Windows.Forms.Label
    $AvailableLicensesLabel.Text = "Available License Types:"
    $AvailableLicensesLabel.Location = New-Object System.Drawing.Point(10, $y)
    $AvailableLicensesLabel.Size = New-Object System.Drawing.Size(200, 20)
    $AvailableLicensesLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    
    $y += 25
    
    $LicenseTypesListLabel = New-Object System.Windows.Forms.Label
    $LicenseTypesListLabel.Text = "• BusinessBasic`n• BusinessPremium`n• BusinessStandard`n• E3`n• E5`n• ExchangeOnline1`n• ExchangeOnline2"
    $LicenseTypesListLabel.Location = New-Object System.Drawing.Point(10, $y)
    $LicenseTypesListLabel.Size = New-Object System.Drawing.Size(450, 140)
    $LicenseTypesListLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    $LicenseTypesListLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    
    $ManagementGroup.Controls.AddRange(@(
        $ManagerLabel, $Script:ManagerDropdown,
        $LicenseTypeLabel, $Script:LicenseTypeDropdown,
        $LicenseInfoLabel, $AvailableLicensesLabel, $LicenseTypesListLabel
    ))
    
    # Groups Selection (Full Width Below)
    $GroupsGroup = New-Object System.Windows.Forms.GroupBox
    $GroupsGroup.Text = "Group Membership (Select groups to add user to)"
    $GroupsGroup.Location = New-Object System.Drawing.Point(10, 470)
    $GroupsGroup.Size = New-Object System.Drawing.Size(970, 250)
    $GroupsGroup.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    
    $Script:GroupsCheckedListBox = New-Object System.Windows.Forms.CheckedListBox
    $Script:GroupsCheckedListBox.Location = New-Object System.Drawing.Point(10, 20)
    $Script:GroupsCheckedListBox.Size = New-Object System.Drawing.Size(950, 220)
    $Script:GroupsCheckedListBox.CheckOnClick = $true
    $Script:GroupsCheckedListBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    $GroupsGroup.Controls.Add($Script:GroupsCheckedListBox)
    
    # Action Buttons
    $CreateUserButton = New-Object System.Windows.Forms.Button
    $CreateUserButton.Text = "👤 Create M365 User"
    $CreateUserButton.Location = New-Object System.Drawing.Point(10, 730)
    $CreateUserButton.Size = New-Object System.Drawing.Size(200, 40)
    $CreateUserButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $CreateUserButton.BackColor = [System.Drawing.Color]::LightGreen
    
    $CreateUserButton.Add_Click({
        Invoke-CreateUser
    })
    
    $ClearFormButton = New-Object System.Windows.Forms.Button
    $ClearFormButton.Text = "🗑️ Clear Form"
    $ClearFormButton.Location = New-Object System.Drawing.Point(220, 730)
    $ClearFormButton.Size = New-Object System.Drawing.Size(120, 40)
    $ClearFormButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    $ClearFormButton.Add_Click({
        Clear-UserCreationForm
    })
    
    # Add all controls to content panel
    $ContentPanel.Controls.AddRange(@(
        $UserDetailsGroup, $ManagementGroup, $GroupsGroup,
        $CreateUserButton, $ClearFormButton
    ))
    
    # Add panels to tab
    $Tab.Controls.Add($ContentPanel)
    $Tab.Controls.Add($ConnectionPanel)
    
    return $Tab
}

function New-BulkImportTab {
    $Tab = New-Object System.Windows.Forms.TabPage
    $Tab.Text = "📊 Bulk Import"
    $Tab.Padding = New-Object System.Windows.Forms.Padding(10)
    
    $InfoLabel = New-Object System.Windows.Forms.Label
    $InfoLabel.Text = "Bulk CSV import functionality - Connect to M365 to enable"
    $InfoLabel.Location = New-Object System.Drawing.Point(20, 20)
    $InfoLabel.Size = New-Object System.Drawing.Size(400, 30)
    
    $Tab.Controls.Add($InfoLabel)
    return $Tab
}

function New-TenantDataTab {
    <#
    .SYNOPSIS
        Creates comprehensive tenant data discovery tab
    #>
    
    $Tab = New-Object System.Windows.Forms.TabPage
    $Tab.Text = "🏢 Tenant Data"
    $Tab.Padding = New-Object System.Windows.Forms.Padding(10)
    
    # Main container
    $MainPanel = New-Object System.Windows.Forms.Panel
    $MainPanel.Dock = "Fill"
    
    # Tenant info panel (top)
    $TenantInfoPanel = New-Object System.Windows.Forms.Panel
    $TenantInfoPanel.Height = 200
    $TenantInfoPanel.Dock = "Top"
    
    $Script:TenantDataTextBox = New-Object System.Windows.Forms.TextBox
    $Script:TenantDataTextBox.Multiline = $true
    $Script:TenantDataTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $Script:TenantDataTextBox.Dock = "Fill"
    $Script:TenantDataTextBox.ReadOnly = $true
    $Script:TenantDataTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $Script:TenantDataTextBox.Text = "Connect to Microsoft 365 to view comprehensive tenant data..."
    
    $TenantInfoPanel.Controls.Add($Script:TenantDataTextBox)
    
    # Data lists panel (bottom)
    $DataListsPanel = New-Object System.Windows.Forms.Panel
    $DataListsPanel.Dock = "Fill"
    
    # Create tab control for data lists
    $DataTabControl = New-Object System.Windows.Forms.TabControl
    $DataTabControl.Dock = "Fill"
    
    # Users tab
    $UsersTab = New-Object System.Windows.Forms.TabPage
    $UsersTab.Text = "👥 Users"
    
    $Script:UsersListBox = New-Object System.Windows.Forms.ListBox
    $Script:UsersListBox.Dock = "Fill"
    $Script:UsersListBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $UsersTab.Controls.Add($Script:UsersListBox)
    
    # Groups tab
    $GroupsTab = New-Object System.Windows.Forms.TabPage
    $GroupsTab.Text = "🏢 Groups"
    
    $Script:GroupsListBox = New-Object System.Windows.Forms.ListBox
    $Script:GroupsListBox.Dock = "Fill"
    $Script:GroupsListBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $GroupsTab.Controls.Add($Script:GroupsListBox)
    
    # Licenses tab
    $LicensesTab = New-Object System.Windows.Forms.TabPage
    $LicensesTab.Text = "🎫 Licenses"
    
    $Script:LicensesListBox = New-Object System.Windows.Forms.ListBox
    $Script:LicensesListBox.Dock = "Fill"
    $Script:LicensesListBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $LicensesTab.Controls.Add($Script:LicensesListBox)
    
    # Domains tab
    $DomainsTab = New-Object System.Windows.Forms.TabPage
    $DomainsTab.Text = "🌐 Domains"
    
    $Script:DomainsListBox = New-Object System.Windows.Forms.ListBox
    $Script:DomainsListBox.Dock = "Fill"
    $Script:DomainsListBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $DomainsTab.Controls.Add($Script:DomainsListBox)
    
    $DataTabControl.TabPages.AddRange(@($UsersTab, $GroupsTab, $LicensesTab, $DomainsTab))
    $DataListsPanel.Controls.Add($DataTabControl)
    
    $MainPanel.Controls.AddRange(@($TenantInfoPanel, $DataListsPanel))
    $Tab.Controls.Add($MainPanel)
    
    return $Tab
}

function New-ActivityLogTab {
    $Tab = New-Object System.Windows.Forms.TabPage
    $Tab.Text = "📋 Activity Log"
    $Tab.Padding = New-Object System.Windows.Forms.Padding(10)
    
    $Script:ActivityLogTextBox = New-Object System.Windows.Forms.TextBox
    $Script:ActivityLogTextBox.Multiline = $true
    $Script:ActivityLogTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $Script:ActivityLogTextBox.Dock = "Fill"
    $Script:ActivityLogTextBox.ReadOnly = $true
    $Script:ActivityLogTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $Script:ActivityLogTextBox.Text = "$(Get-Date): M365 User Provisioning Tool - Enterprise Edition started`r`n"
    
    $Tab.Controls.Add($Script:ActivityLogTextBox)
    return $Tab
}

function Start-TenantDataDiscovery {
    <#
    .SYNOPSIS
        Performs comprehensive tenant data discovery
    #>
    
    try {
        Update-StatusLabel "🔍 Starting comprehensive tenant discovery..."
        Add-ActivityLog "Starting tenant data discovery"
        
        # Get tenant information
        if (Get-Command "Get-MgOrganization" -ErrorAction SilentlyContinue) {
            $Script:TenantInfo = Get-MgOrganization | Select-Object -First 1
        }
        
        # Get domains
        if (Get-Command "Get-MgDomain" -ErrorAction SilentlyContinue) {
            $Script:TenantDomains = Get-MgDomain | Where-Object { $_.IsVerified -eq $true }
            Update-DomainDropdown $Script:TenantDomains
        }
        
        # Get users
        Update-StatusLabel "👥 Discovering users..."
        if (Get-Command "Get-MgUser" -ErrorAction SilentlyContinue) {
            $Script:TenantUsers = Get-MgUser -All -Property "DisplayName,UserPrincipalName,JobTitle,Department" | Sort-Object DisplayName
            Update-ManagerDropdown $Script:TenantUsers
        }
        
        # Get groups
        Update-StatusLabel "🏢 Discovering groups..."
        if (Get-Command "Get-MgGroup" -ErrorAction SilentlyContinue) {
            $Script:TenantGroups = Get-MgGroup -All -Property "DisplayName,GroupTypes" | Sort-Object DisplayName
            Update-GroupsList $Script:TenantGroups
        }
        
        # Get licenses
        Update-StatusLabel "🎫 Discovering licenses..."
        if (Get-Command "Get-MgSubscribedSku" -ErrorAction SilentlyContinue) {
            $Script:TenantLicenses = Get-MgSubscribedSku
        }
        
        # Update displays
        Update-TenantDataDisplay
        $Script:TenantDataLoaded = $true
        
        Update-StatusLabel "✅ Tenant discovery completed successfully"
        Add-ActivityLog "Tenant discovery completed - $($Script:TenantUsers.Count) users, $($Script:TenantGroups.Count) groups, $($Script:TenantLicenses.Count) licenses"
        
    }
    catch {
        $ErrorMsg = "Tenant discovery failed: $($_.Exception.Message)"
        Update-StatusLabel "❌ $ErrorMsg"
        Add-ActivityLog "Tenant discovery failed: $($_.Exception.Message)"
        
        [System.Windows.Forms.MessageBox]::Show(
            "Tenant discovery encountered errors:`n`n$($_.Exception.Message)",
            "Discovery Warning",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
}

function Update-TenantDataDisplay {
    <#
    .SYNOPSIS
        Updates tenant data displays with discovered information
    #>
    
    if ($Script:TenantDataTextBox) {
        $TenantSummary = @"
MICROSOFT 365 TENANT INFORMATION
=================================
Tenant Name: $($Script:TenantInfo.DisplayName)
Tenant ID: $($Script:TenantInfo.Id)
Country: $($Script:TenantInfo.CountryLetterCode)
Discovery Date: $(Get-Date)

COMPREHENSIVE DISCOVERY RESULTS
===============================
✅ Users Discovered: $($Script:TenantUsers.Count)
✅ Groups Discovered: $($Script:TenantGroups.Count)
✅ License SKUs: $($Script:TenantLicenses.Count)
✅ Verified Domains: $($Script:TenantDomains.Count)

VERIFIED DOMAINS
================
$($Script:TenantDomains | ForEach-Object { "• $($_.Id) $(if($_.IsDefault){'(Default)'})" } | Out-String)

AVAILABLE LICENSE SKUS
======================
$($Script:TenantLicenses | ForEach-Object { "• $($_.SkuPartNumber) - $($_.ConsumedUnits)/$($_.PrepaidUnits.Enabled) used" } | Out-String)
"@
        
        $Script:TenantDataTextBox.Text = $TenantSummary
    }
    
    # Update individual data lists
    if ($Script:UsersListBox) {
        $Script:UsersListBox.Items.Clear()
        foreach ($User in $Script:TenantUsers) {
            $UserDisplay = "$($User.DisplayName) ($($User.UserPrincipalName))"
            if ($User.JobTitle) { $UserDisplay += " - $($User.JobTitle)" }
            $Script:UsersListBox.Items.Add($UserDisplay) | Out-Null
        }
    }
    
    if ($Script:GroupsListBox) {
        $Script:GroupsListBox.Items.Clear()
        foreach ($Group in $Script:TenantGroups) {
            $GroupType = if ($Group.GroupTypes -contains "Unified") { "M365" } else { "Security" }
            $Script:GroupsListBox.Items.Add("[$GroupType] $($Group.DisplayName)") | Out-Null
        }
    }
    
    if ($Script:LicensesListBox) {
        $Script:LicensesListBox.Items.Clear()
        foreach ($License in $Script:TenantLicenses) {
            $LicenseDisplay = "$($License.SkuPartNumber) - $($License.ConsumedUnits)/$($License.PrepaidUnits.Enabled) assigned"
            $Script:LicensesListBox.Items.Add($LicenseDisplay) | Out-Null
        }
    }
    
    if ($Script:DomainsListBox) {
        $Script:DomainsListBox.Items.Clear()
        foreach ($Domain in $Script:TenantDomains) {
            $DomainDisplay = "$($Domain.Id) $(if($Domain.IsDefault){'(Default)'})"
            $Script:DomainsListBox.Items.Add($DomainDisplay) | Out-Null
        }
    }
}

# Helper Functions
function Update-DisplayName {
    if ($Script:FirstNameTextBox.Text -and $Script:LastNameTextBox.Text) {
        $Script:DisplayNameTextBox.Text = "$($Script:FirstNameTextBox.Text) $($Script:LastNameTextBox.Text)"
    }
}

function Generate-SecurePassword {
    $Length = 12
    $Characters = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%^&*"
    $Password = -join ((1..$Length) | ForEach-Object { $Characters[(Get-Random -Maximum $Characters.Length)] })
    return $Password
}

function Invoke-CreateUser {
    # Validate required fields
    if (-not $Script:FirstNameTextBox.Text.Trim()) {
        [System.Windows.Forms.MessageBox]::Show("First Name is required.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    if (-not $Script:LastNameTextBox.Text.Trim()) {
        [System.Windows.Forms.MessageBox]::Show("Last Name is required.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    if (-not $Script:UsernameTextBox.Text.Trim()) {
        [System.Windows.Forms.MessageBox]::Show("Username is required.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    if (-not $Script:PasswordTextBox.Text.Trim()) {
        [System.Windows.Forms.MessageBox]::Show("Password is required.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    if ($Script:DomainDropdown.SelectedItem -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Please select a domain.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    # Get selected groups
    $SelectedGroups = @()
    for ($i = 0; $i -lt $Script:GroupsCheckedListBox.Items.Count; $i++) {
        if ($Script:GroupsCheckedListBox.GetItemChecked($i)) {
            $SelectedGroups += $Script:GroupsCheckedListBox.Items[$i]
        }
    }
    
    # Create user parameters
    $UserParams = @{
        FirstName = $Script:FirstNameTextBox.Text.Trim()
        LastName = $Script:LastNameTextBox.Text.Trim()
        Username = $Script:UsernameTextBox.Text.Trim()
        Domain = $Script:DomainDropdown.SelectedItem
        Password = $Script:PasswordTextBox.Text
        Department = $Script:DepartmentTextBox.Text.Trim()
        JobTitle = $Script:JobTitleTextBox.Text.Trim()
        Office = $Script:OfficeLocationTextBox.Text  # CHANGED: Now uses TextBox instead of SelectedItem
        Manager = $Script:ManagerDropdown.SelectedItem
        LicenseType = $Script:LicenseTypeDropdown.SelectedItem
        Groups = $SelectedGroups
    }
    
    # Call user creation function from M365.UserManagement module
    try {
        Update-StatusLabel "👤 Creating user..."
        Add-ActivityLog "Starting user creation: $($UserParams.FirstName) $($UserParams.LastName)"
        
        if (Get-Command "New-M365User" -ErrorAction SilentlyContinue) {
            $NewUser = New-M365User @UserParams
            
            if ($NewUser) {
                [System.Windows.Forms.MessageBox]::Show(
                    "User created successfully!`n`nName: $($UserParams.FirstName) $($UserParams.LastName)`nUPN: $($UserParams.Username)@$($UserParams.Domain)`nLicense Type: $($UserParams.LicenseType)",
                    "User Creation Successful",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
                
                Clear-UserCreationForm
                Update-StatusLabel "✅ User created successfully"
                Add-ActivityLog "User created successfully: $($UserParams.Username)@$($UserParams.Domain)"
            }
        } else {
            throw "New-M365User function not available. Check M365.UserManagement module."
        }
    }
    catch {
        $ErrorMsg = "Failed to create user: $($_.Exception.Message)"
        Update-StatusLabel "❌ $ErrorMsg"
        Add-ActivityLog "User creation failed: $($_.Exception.Message)"
        
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to create user:`n`n$($_.Exception.Message)",
            "User Creation Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

# UI Update Functions (Required by main script)
function Update-StatusLabel {
    param([string]$Message)
    
    if ($Script:StatusLabel) {
        $Script:StatusLabel.Text = "$(Get-Date -Format 'HH:mm:ss') - $Message"
        [System.Windows.Forms.Application]::DoEvents()
    }
    Write-Host "📊 $Message" -ForegroundColor Cyan
}

function Add-ActivityLog {
    param([string]$Message)
    
    if ($Script:ActivityLogTextBox) {
        $LogEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message"
        $Script:ActivityLogTextBox.AppendText("$LogEntry`r`n")
        $Script:ActivityLogTextBox.ScrollToCaret()
    }
}

function Update-UIAfterConnection {
    # Update connection buttons
    if ($Script:ConnectButton) {
        $Script:ConnectButton.Text = "✅ Connected to M365"
        $Script:ConnectButton.BackColor = [System.Drawing.Color]::LightGreen
        $Script:ConnectButton.Enabled = $false
    }
    
    if ($Script:RefreshDataButton) {
        $Script:RefreshDataButton.Enabled = $true
    }
    
    if ($Script:DisconnectButton) {
        $Script:DisconnectButton.Enabled = $true
    }
    
    Update-StatusLabel "✅ Connected - Starting tenant discovery..."
    Add-ActivityLog "Connected to Microsoft 365"
}

function Update-UIAfterDisconnection {
    # Reset connection buttons
    if ($Script:ConnectButton) {
        $Script:ConnectButton.Text = "🔗 Connect to M365"
        $Script:ConnectButton.BackColor = [System.Drawing.Color]::LightBlue
        $Script:ConnectButton.Enabled = $true
    }
    
    if ($Script:RefreshDataButton) {
        $Script:RefreshDataButton.Enabled = $false
    }
    
    if ($Script:DisconnectButton) {
        $Script:DisconnectButton.Enabled = $false
    }
    
    # Clear tenant data
    Clear-AllDropdowns
    $Script:TenantDataLoaded = $false
    
    Update-StatusLabel "Disconnected from Microsoft 365"
    Add-ActivityLog "Disconnected from Microsoft 365"
}

function Update-DomainDropdown {
    param([array]$Domains)
    
    if ($Script:DomainDropdown) {
        $Script:DomainDropdown.Items.Clear()
        foreach ($Domain in $Domains) {
            $Script:DomainDropdown.Items.Add($Domain.Id) | Out-Null
        }
        if ($Script:DomainDropdown.Items.Count -gt 0) {
            $Script:DomainDropdown.SelectedIndex = 0
        }
    }
}

function Update-ManagerDropdown {
    param([array]$Users)
    
    if ($Script:ManagerDropdown) {
        $Script:ManagerDropdown.Items.Clear()
        $Script:ManagerDropdown.Items.Add("(No Manager)") | Out-Null
        foreach ($User in $Users) {
            $ManagerDisplay = "$($User.DisplayName) ($($User.UserPrincipalName))"
            $Script:ManagerDropdown.Items.Add($ManagerDisplay) | Out-Null
        }
        $Script:ManagerDropdown.SelectedIndex = 0
    }
}

function Update-GroupsList {
    param([array]$Groups)
    
    if ($Script:GroupsCheckedListBox) {
        $Script:GroupsCheckedListBox.Items.Clear()
        foreach ($Group in $Groups) {
            $GroupType = if ($Group.GroupTypes -contains "Unified") { "📧" } else { "🔒" }
            $GroupDisplay = "$GroupType $($Group.DisplayName)"
            $Script:GroupsCheckedListBox.Items.Add($GroupDisplay) | Out-Null
        }
    }
}

function Update-LicenseDropdown {
    param([array]$Licenses)
    
    # This function is called by the main script but we use static license types
    # for CustomAttribute1 assignment, so no action needed
}

function Clear-AllDropdowns {
    if ($Script:DomainDropdown) { $Script:DomainDropdown.Items.Clear() }
    if ($Script:ManagerDropdown) { 
        $Script:ManagerDropdown.Items.Clear()
        $Script:ManagerDropdown.Items.Add("(No Manager)") | Out-Null
    }
    if ($Script:GroupsCheckedListBox) { $Script:GroupsCheckedListBox.Items.Clear() }
    
    # Clear data lists
    if ($Script:UsersListBox) { $Script:UsersListBox.Items.Clear() }
    if ($Script:GroupsListBox) { $Script:GroupsListBox.Items.Clear() }
    if ($Script:LicensesListBox) { $Script:LicensesListBox.Items.Clear() }
    if ($Script:DomainsListBox) { $Script:DomainsListBox.Items.Clear() }
    
    # Clear tenant data display
    if ($Script:TenantDataTextBox) {
        $Script:TenantDataTextBox.Text = "Connect to Microsoft 365 to view tenant data..."
    }
}

function Refresh-TenantDataViews {
    if ($Script:TenantDataLoaded) {
        Update-TenantDataDisplay
    }
}

function Clear-UserCreationForm {
    # Clear all form fields
    if ($Script:FirstNameTextBox) { $Script:FirstNameTextBox.Clear() }
    if ($Script:LastNameTextBox) { $Script:LastNameTextBox.Clear() }
    if ($Script:DisplayNameTextBox) { $Script:DisplayNameTextBox.Clear() }
    if ($Script:UsernameTextBox) { $Script:UsernameTextBox.Clear() }
    if ($Script:PasswordTextBox) { $Script:PasswordTextBox.Clear() }
    if ($Script:DepartmentTextBox) { $Script:DepartmentTextBox.Clear() }
    if ($Script:JobTitleTextBox) { $Script:JobTitleTextBox.Clear() }
    if ($Script:OfficePhoneTextBox) { $Script:OfficePhoneTextBox.Clear() }
    if ($Script:MobilePhoneTextBox) { $Script:MobilePhoneTextBox.Clear() }
    
    # Clear office location text box - CHANGED: Now clears TextBox instead of resetting dropdown
    if ($Script:OfficeLocationTextBox) {
        $Script:OfficeLocationTextBox.Clear()
    }
    
    # Reset dropdowns to default selections
    if ($Script:ManagerDropdown -and $Script:ManagerDropdown.Items.Count -gt 0) {
        $Script:ManagerDropdown.SelectedIndex = 0
    }
    
    if ($Script:LicenseTypeDropdown -and $Script:LicenseTypeDropdown.Items.Count -gt 0) {
        $Script:LicenseTypeDropdown.SelectedIndex = 0
    }
    
    # Uncheck all groups
    if ($Script:GroupsCheckedListBox) {
        for ($i = 0; $i -lt $Script:GroupsCheckedListBox.Items.Count; $i++) {
            $Script:GroupsCheckedListBox.SetItemChecked($i, $false)
        }
    }
    
    Add-ActivityLog "User creation form cleared"
}

# Export all 16 GUI functions expected by the main script
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