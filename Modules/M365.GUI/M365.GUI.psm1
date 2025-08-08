#Requires -Version 7.0

<#
.SYNOPSIS
    M365 GUI Module - Fixed Windows Forms Initialization
    
.DESCRIPTION
    This fixed version properly handles Windows Forms initialization by:
    1. Only initializing Windows Forms when explicitly called
    2. Using a global flag to prevent duplicate initialization
    3. Following Microsoft's recommended order of operations
    
.NOTES
    Key Fix: Windows Forms initialization moved to Start-M365ProvisioningTool function
    instead of module import time.
#>

# Global variable to track initialization state
$Script:WindowsFormsInitialized = $false

#region Windows Forms Initialization Functions
function Initialize-WindowsForms {
    <#
    .SYNOPSIS
        Properly initializes Windows Forms following Microsoft best practices
    #>
    
    if ($Script:WindowsFormsInitialized) {
        Write-Verbose "Windows Forms already initialized, skipping..."
        return $true
    }
    
    try {
        Write-Host "🔧 Initializing Windows Forms subsystem..." -ForegroundColor Cyan
        
        # Step 1: Load Windows Forms assemblies FIRST
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        Write-Verbose "✓ Windows Forms assemblies loaded"
        
        # Step 2: Enable visual styles IMMEDIATELY after assembly loading
        [System.Windows.Forms.Application]::EnableVisualStyles()
        Write-Verbose "✓ Visual styles enabled"
        
        # Step 3: Set compatible text rendering DEFAULT - CRITICAL TIMING
        # This MUST happen before ANY Windows Forms objects are created
        [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
        Write-Verbose "✓ Text rendering compatibility set"
        
        # Step 4: Mark as initialized to prevent duplicate calls
        $Script:WindowsFormsInitialized = $true
        
        Write-Host "✅ Windows Forms initialization completed successfully" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "❌ Failed to initialize Windows Forms: $($_.Exception.Message)"
        Write-Host "This typically indicates:" -ForegroundColor Yellow
        Write-Host "  1. Running on a non-Windows platform" -ForegroundColor White
        Write-Host "  2. Missing .NET Framework components" -ForegroundColor White
        Write-Host "  3. Running in a headless environment" -ForegroundColor White
        return $false
    }
}

function Test-WindowsFormsAvailability {
    <#
    .SYNOPSIS
        Tests if Windows Forms is available in the current environment
    #>
    
    try {
        # Test if we're on Windows
        if (-not $IsWindows -and $PSVersionTable.PSVersion.Major -ge 6) {
            return $false
        }
        
        # Test if assemblies can be loaded
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}
#endregion

#region Main Application Function
function Start-M365ProvisioningTool {
    <#
    .SYNOPSIS
        Main entry point for the M365 User Provisioning Tool
        
    .DESCRIPTION
        This function properly initializes Windows Forms BEFORE creating any GUI objects,
        then launches the main application form.
    #>
    
    [CmdletBinding()]
    param()
    
    try {
        Write-Host "🚀 Starting M365 User Provisioning Tool - Enterprise Edition 2025" -ForegroundColor Cyan
        Write-Host "=================================================================" -ForegroundColor Cyan
        
        # CRITICAL: Initialize Windows Forms FIRST, before any other GUI operations
        if (-not (Test-WindowsFormsAvailability)) {
            throw "Windows Forms is not available in this environment. This tool requires Windows with GUI support."
        }
        
        if (-not (Initialize-WindowsForms)) {
            throw "Failed to initialize Windows Forms subsystem."
        }
        
        # NOW it's safe to create Windows Forms objects
        Write-Host "🖥️  Creating main application form..." -ForegroundColor Green
        
        # Create and configure the main form
        $Script:MainForm = New-MainForm
        
        if (-not $Script:MainForm) {
            throw "Failed to create main application form."
        }
        
        Write-Host "📱 Launching GUI interface..." -ForegroundColor Green
        
        # Show the form and run the application message loop
        $Result = $Script:MainForm.ShowDialog()
        
        Write-Host "📊 Application closed with result: $Result" -ForegroundColor Gray
        return $Result
        
    }
    catch {
        $ErrorMessage = "🚨 Application startup failed: $($_.Exception.Message)"
        Write-Error $ErrorMessage
        
        # Try to show error dialog if Windows Forms is available
        if ($Script:WindowsFormsInitialized) {
            try {
                [System.Windows.Forms.MessageBox]::Show(
                    $ErrorMessage,
                    "M365 Provisioning Tool - Startup Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
            catch {
                # Fallback to console output
                Write-Host $ErrorMessage -ForegroundColor Red
            }
        }
        else {
            Write-Host $ErrorMessage -ForegroundColor Red
        }
        
        throw
    }
}
#endregion

#region Form Creation Functions  
function New-MainForm {
    <#
    .SYNOPSIS
        Creates the main application form with tabbed interface
        
    .DESCRIPTION
        This function creates the main form AFTER Windows Forms has been properly initialized.
        It's safe to call this only after Initialize-WindowsForms has succeeded.
    #>
    
    if (-not $Script:WindowsFormsInitialized) {
        throw "Windows Forms must be initialized before creating forms. Call Initialize-WindowsForms first."
    }
    
    try {
        Write-Verbose "Creating main application form..."
        
        # Create the main form
        $Form = New-Object System.Windows.Forms.Form
        $Form.Text = "M365 User Provisioning Tool - Enterprise Edition 2025"
        $Form.Size = New-Object System.Drawing.Size(1400, 900)
        $Form.StartPosition = "CenterScreen"
        $Form.MinimumSize = New-Object System.Drawing.Size(1200, 800)
        $Form.MaximizeBox = $true
        $Form.WindowState = "Maximized"
        
        # Set application icon
        try {
            $Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("$env:SystemRoot\System32\shell32.dll")
        }
        catch {
            Write-Verbose "Could not set application icon: $($_.Exception.Message)"
        }
        
        # Create status strip at bottom
        $Script:StatusStrip = New-Object System.Windows.Forms.StatusStrip
        $Script:StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
        $Script:StatusLabel.Text = "Ready - Not connected to Microsoft 365"
        $Script:StatusLabel.Spring = $true
        $Script:StatusLabel.TextAlign = "MiddleLeft"
        $Script:StatusStrip.Items.Add($Script:StatusLabel) | Out-Null
        
        # Create tab control
        $Script:TabControl = New-Object System.Windows.Forms.TabControl
        $Script:TabControl.Dock = "Fill"
        $Script:TabControl.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
        
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
            Write-Host "📋 Main form loaded successfully" -ForegroundColor Green
            Update-StatusLabel "Application started - Ready to connect to Microsoft 365"
        })
        
        $Form.Add_FormClosing({
            param($sender, $e)
            
            if ($Global:AppState.Connected) {
                $Result = [System.Windows.Forms.MessageBox]::Show(
                    "You are currently connected to Microsoft 365. Do you want to disconnect and exit?",
                    "Confirm Exit",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question
                )
                
                if ($Result -eq [System.Windows.Forms.DialogResult]::No) {
                    $e.Cancel = $true
                    return
                }
            }
            
            Write-Host "👋 Application shutting down..." -ForegroundColor Yellow
        })
        
        Write-Verbose "✓ Main form created successfully"
        return $Form
        
    }
    catch {
        Write-Error "Failed to create main form: $($_.Exception.Message)"
        throw
    }
}

function New-UserCreationTab {
    <#
    .SYNOPSIS
        Creates the user creation tab
    #>
    
    $Tab = New-Object System.Windows.Forms.TabPage
    $Tab.Text = "👤 Create User"
    $Tab.Name = "UserCreationTab"
    
    # Connection Panel
    $ConnectionPanel = New-Object System.Windows.Forms.Panel
    $ConnectionPanel.Height = 60
    $ConnectionPanel.Dock = "Top"
    $ConnectionPanel.BackColor = [System.Drawing.Color]::LightBlue
    
    $ConnectButton = New-Object System.Windows.Forms.Button
    $ConnectButton.Text = "🔗 Connect to M365"
    $ConnectButton.Size = New-Object System.Drawing.Size(150, 30)
    $ConnectButton.Location = New-Object System.Drawing.Point(20, 15)
    $ConnectButton.Add_Click({
        try {
            Update-StatusLabel "Connecting to Microsoft 365..."
            # Call authentication function here
            Connect-ToMicrosoftGraph
            Update-UIAfterConnection
            Update-StatusLabel "✅ Connected to Microsoft 365"
        }
        catch {
            Update-StatusLabel "❌ Connection failed: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to connect to Microsoft 365:`n$($_.Exception.Message)",
                "Connection Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    })
    
    $ConnectionPanel.Controls.Add($ConnectButton)
    $Tab.Controls.Add($ConnectionPanel)
    
    # User Creation Form Panel
    $FormPanel = New-Object System.Windows.Forms.Panel
    $FormPanel.Dock = "Fill"
    $FormPanel.AutoScroll = $true
    $FormPanel.Padding = New-Object System.Windows.Forms.Padding(20)
    
    # Add user creation form controls here
    # (Implementation details would continue...)
    
    $Tab.Controls.Add($FormPanel)
    return $Tab
}

function New-BulkImportTab {
    # Placeholder for bulk import tab
    $Tab = New-Object System.Windows.Forms.TabPage
    $Tab.Text = "📊 Bulk Import"
    $Tab.Name = "BulkImportTab"
    return $Tab
}

function New-TenantDataTab {
    # Placeholder for tenant data tab  
    $Tab = New-Object System.Windows.Forms.TabPage
    $Tab.Text = "🏢 Tenant Data"
    $Tab.Name = "TenantDataTab"
    return $Tab
}

function New-ActivityLogTab {
    # Placeholder for activity log tab
    $Tab = New-Object System.Windows.Forms.TabPage
    $Tab.Text = "📋 Activity Log"
    $Tab.Name = "ActivityLogTab"
    return $Tab
}
#endregion

#region Utility Functions
function Update-StatusLabel {
    param([string]$Message)
    
    if ($Script:StatusLabel) {
        $Script:StatusLabel.Text = "$(Get-Date -Format 'HH:mm:ss') - $Message"
        $Script:StatusLabel.GetCurrentParent().Refresh()
    }
    Write-Host "📊 $Message" -ForegroundColor Cyan
}

function Update-UIAfterConnection {
    # Enable/disable controls based on connection status
    $Global:AppState.Connected = $true
    # Implementation would continue...
}

function Update-UIAfterDisconnection {
    # Reset UI after disconnection
    $Global:AppState.Connected = $false
    # Implementation would continue...
}

function Clear-AllDropdowns {
    # Clear all dropdown controls
    # Implementation would continue...
}

function Refresh-TenantDataViews {
    # Refresh tenant data displays
    # Implementation would continue...
}

function Clear-UserCreationForm {
    # Clear user creation form
    # Implementation would continue...
}
#endregion

# Export only the main function - all other functions are internal
Export-ModuleMember -Function @(
    'Start-M365ProvisioningTool'
)