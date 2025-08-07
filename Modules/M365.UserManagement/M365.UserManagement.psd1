@{
# Script module or binary module file associated with this manifest.
RootModule = 'M365.UserManagement.psm1'

# Version number of this module.
ModuleVersion = '1.0.0'

# Supported PSEditions
CompatiblePSEditions = 'Core'

# ID used to uniquely identify this module
GUID = 'a8f9d2c1-6b4e-4f3a-9e8d-1c2b3a4f5e6d'

# Author of this module
Author = 'Tom Mortiboys'

# Company or vendor of this module
CompanyName = 'M365 Project Delivery'

# Copyright statement for this module
Copyright = '(c) Tom Mortiboys. All rights reserved.'

# Description of the functionality provided by this module
Description = 'M365 User Management module providing user creation and CSV import functionality'

# Minimum version of the PowerShell engine required by this module
PowerShellVersion = '7.0'

# Functions to export from this module
FunctionsToExport = @(
    'New-M365User',
    'Import-UsersFromCSV',
    'New-SecurePassword'
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
        Tags = @('M365', 'UserManagement', 'Azure', 'MicrosoftGraph')
        
        # External dependent modules of this module
        # ExternalModuleDependencies = @()
    }
}
}