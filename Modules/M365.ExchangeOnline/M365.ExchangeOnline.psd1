# M365.ExchangeOnline.psd1
# Module manifest for M365.ExchangeOnline module

@{
    # Script module or binary module file associated with this manifest.
    RootModule = 'M365.ExchangeOnline.psm1'

    # Version number of this module.
    ModuleVersion = '1.0.0'

    # Supported PSEditions
    CompatiblePSEditions = @('Core', 'Desktop')

    # ID used to uniquely identify this module
    GUID = 'a8b7c9d2-4e5f-6789-abc1-def234567890'

    # Author of this module
    Author = 'Tom Mortiboys'

    # Company or vendor of this module
    CompanyName = 'M365 Project Delivery'

    # Copyright statement for this module
    Copyright = '(c) Tom Mortiboys. All rights reserved.'

    # Description of the functionality provided by this module
    Description = 'Exchange Online operations module for M365 User Provisioning Tool. Provides shared mailboxes, distribution lists, and mail-enabled security group management.'

    # Minimum version of the PowerShell engine required by this module
    PowerShellVersion = '7.0'

    # Modules that must be imported into the global environment prior to importing this module
    RequiredModules = @(
        @{
            ModuleName = 'ExchangeOnlineManagement'
            ModuleVersion = '3.2.0'
        },
        @{
            ModuleName = 'Microsoft.Graph.Users'
            ModuleVersion = '2.0.0'
        },
        @{
            ModuleName = 'Microsoft.Graph.Groups'
            ModuleVersion = '2.0.0'
        }
    )

    # Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
    NestedModules = @()

    # Functions to export from this module - Exchange Online specific functions
    FunctionsToExport = @(
        'Get-ExchangeMailboxData',
        'Get-ExchangeDistributionGroupData', 
        'Get-ExchangeAcceptedDomains',
        'Add-UserToSharedMailbox',
        'Add-UserToDistributionList',
        'Add-UserToMailEnabledSecurityGroup',
        'Get-AllExchangeData',
        'Invoke-ExchangeUserProvisioning'
    )

    # Cmdlets to export from this module
    CmdletsToExport = @()

    # Variables to export from this module
    VariablesToExport = @()

    # Aliases to export from this module
    AliasesToExport = @()

    # DSC resources to export from this module
    DscResourcesToExport = @()

    # List of all modules packaged with this module
    ModuleList = @()

    # List of all files packaged with this module
    FileList = @()

    # Private data to pass to the module specified in RootModule/ModuleToProcess
    PrivateData = @{
        PSData = @{
            # Tags applied to this module. These help with module discovery in online galleries.
            Tags = @('M365', 'ExchangeOnline', 'SharedMailboxes', 'DistributionLists', 'MailEnabledSecurityGroups', 'UserProvisioning')

            # A URL to the license for this module.
            LicenseUri = ''

            # A URL to the main website for this project.
            ProjectUri = ''

            # A URL to an icon representing this module.
            IconUri = ''

            # ReleaseNotes of this module
            ReleaseNotes = @'
## Version 1.0.0
- Initial release of M365.ExchangeOnline module
- Exchange Online PowerShell V3 integration
- Shared mailbox discovery and permission management
- Distribution list management
- Mail-enabled security group management
- Graceful fallback to Graph API when Exchange Online unavailable
- Integration with M365.Authentication module
- Modern authentication support with MFA
'@

            # Prerelease string of this module
            Prerelease = ''

            # Flag to indicate whether the module requires explicit user acceptance for install/update/save
            RequireLicenseAcceptance = $false

            # External dependent modules of this module
            ExternalModuleDependencies = @('ExchangeOnlineManagement')
        }
    }

    # HelpInfo URI of this module
    HelpInfoURI = ''

    # Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
    DefaultCommandPrefix = ''
}