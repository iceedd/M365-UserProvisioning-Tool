@{
# Script module or binary module file associated with this manifest.
RootModule = 'M365.GUI.psm1'

# Version number of this module.
ModuleVersion = '1.0.0'

# Supported PSEditions
CompatiblePSEditions = 'Core'

# ID used to uniquely identify this module
GUID = 'b4d8e5c2-9a7f-4e3d-8c1b-6f2e9d4a7b8c'

# Author of this module
Author = 'Tom Mortiboys'

# Company or vendor of this module
CompanyName = 'M365 Project Delivery'

# Copyright statement for this module
Copyright = '(c) Tom Mortiboys. All rights reserved.'

# Description of the functionality provided by this module
Description = 'M365 GUI module providing Windows Forms interface for user provisioning tool'

# Minimum version of the PowerShell engine required by this module
PowerShellVersion = '7.0'

# Modules that must be imported into the global environment prior to importing this module
RequiredModules = @('M365.Authentication', 'M365.UserManagement')

# Functions to export from this module - ALL 16 GUI FUNCTIONS
FunctionsToExport = @(
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

# Cmdlets to export from this module
CmdletsToExport = @()

# Variables to export from this module
VariablesToExport = @()

# Aliases to export from this module
AliasesToExport = @()

# Private data to pass to the module specified in RootModule/ModuleToProcess
PrivateData = @{
    PSData = @{
        # Tags applied to this module
        Tags = @('M365', 'GUI', 'WindowsForms', 'UserProvisioning')
        
        # External dependent modules of this module
        # ExternalModuleDependencies = @()
    }
}
}