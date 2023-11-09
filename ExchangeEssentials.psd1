@{
    AliasesToExport        = @()
    Author                 = 'Przemyslaw Klys'
    CmdletsToExport        = @()
    CompanyName            = 'Evotec'
    CompatiblePSEditions   = @('Desktop', 'Core')
    Copyright              = '(c) 2011 - 2023 Przemyslaw Klys @ Evotec. All rights reserved.'
    Description            = 'ExchangeEssentials'
    DotNetFrameworkVersion = '4.5.2'
    FunctionsToExport      = @('Get-MyMailbox', 'Get-MyMailboxProblems', 'Invoke-ExchangeEssentials')
    GUID                   = '0dd82757-4cac-4772-9821-1c8ccaebe50d'
    ModuleVersion          = '0.9.0'
    PowerShellVersion      = '5.1'
    PrivateData            = @{
        PSData = @{
            ExternalModuleDependencies = @('ActiveDirectory')
            ProjectUri                 = 'https://github.com/EvotecIT/ExchangeEssentials'
            Tags                       = @('Windows')
        }
    }
    RequiredModules        = @(@{
            Guid          = 'b5eced50-afa4-455b-847a-d8fb64140a22'
            ModuleName    = 'ExchangeOnlineManagement'
            ModuleVersion = '3.3.0'
        }, @{
            Guid          = 'a7bdf640-f5cb-4acf-9de0-365b322d245c'
            ModuleName    = 'PSWriteHTML'
            ModuleVersion = '1.11.0'
        }, 'ActiveDirectory')
    RootModule             = 'ExchangeEssentials.psm1'
}