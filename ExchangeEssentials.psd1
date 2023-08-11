﻿@{
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
    ModuleVersion          = '0.4.0'
    PowerShellVersion      = '5.1'
    PrivateData            = @{
        PSData = @{
            Tags                       = @('Windows')
            ProjectUri                 = 'https://github.com/EvotecIT/ExchangeEssentials'
            ExternalModuleDependencies = @('ActiveDirectory')
        }
    }
    RequiredModules        = @(@{
            ModuleName    = 'ExchangeOnlineManagement'
            ModuleVersion = '3.2.0'
            Guid          = 'b5eced50-afa4-455b-847a-d8fb64140a22'
        }, @{
            ModuleName    = 'PSWriteHTML'
            ModuleVersion = '1.1.0'
            Guid          = 'a7bdf640-f5cb-4acf-9de0-365b322d245c'
        }, 'ActiveDirectory')
    RootModule             = 'ExchangeEssentials.psm1'
}