Clear-Host
Import-Module "C:\Support\GitHub\PSPublishModule\PSPublishModule.psd1" -Force

Invoke-ModuleBuild -ModuleName 'ExchangeEssentials' {
    # Usual defaults as per standard module
    $Manifest = [ordered] @{
        ModuleVersion          = '0.X.0'
        CompatiblePSEditions   = @('Desktop', 'Core')
        GUID                   = '0dd82757-4cac-4772-9821-1c8ccaebe50d'
        Author                 = 'Przemyslaw Klys'
        CompanyName            = 'Evotec'
        Copyright              = "(c) 2011 - $((Get-Date).Year) Przemyslaw Klys @ Evotec. All rights reserved."
        Description            = 'ExchangeEssentials'
        PowerShellVersion      = '5.1'
        Tags                   = @('Windows')
        IconUri                = ''
        ProjectUri             = 'https://github.com/EvotecIT/ExchangeEssentials'
        DotNetFrameworkVersion = '4.5.2'
    }
    New-ConfigurationManifest @Manifest

    New-ConfigurationModule -Type RequiredModule -Name 'ExchangeOnlineManagement' -Version Latest -Guid Auto
    New-ConfigurationModule -Type ApprovedModule -Name 'PSSharedGoods', 'PSWriteColor', 'Connectimo', 'PSUnifi', 'PSWebToolbox', 'PSMyPassword'
    New-ConfigurationModule -Type ExternalModule -Name 'ActiveDirectory'
    New-ConfigurationModuleSkip -IgnoreFunctionName @(
        'Get-Mailbox'
        'Get-CASMailbox'
        'Get-MailboxPermission'
        'Get-MailboxStatistics'
        'Get-MailContact'
        'Get-RecipientPermission'
        'Get-RemoteDomain'
        'Get-LocalMailboxStatistics'
        'Get-LocalADPermission'
        'Get-LocalCasMailbox'
        'Get-LocalMailbox'
        'Get-LocalMailboxPermission'
        'Get-LocalMailContact'
        'Get-LocalRemoteDomain'
        'Get-ExoRecipientPermission'
    ) -IgnoreModuleName 'Microsoft.PowerShell.Management', 'Microsoft.PowerShell.Utility', 'Microsoft.WSMan.Management', 'NetTCPIP'

    $ConfigurationFormat = [ordered] @{
        RemoveComments                              = $false

        PlaceOpenBraceEnable                        = $true
        PlaceOpenBraceOnSameLine                    = $true
        PlaceOpenBraceNewLineAfter                  = $true
        PlaceOpenBraceIgnoreOneLineBlock            = $false

        PlaceCloseBraceEnable                       = $true
        PlaceCloseBraceNewLineAfter                 = $true
        PlaceCloseBraceIgnoreOneLineBlock           = $false
        PlaceCloseBraceNoEmptyLineBefore            = $true

        UseConsistentIndentationEnable              = $true
        UseConsistentIndentationKind                = 'space'
        UseConsistentIndentationPipelineIndentation = 'IncreaseIndentationAfterEveryPipeline'
        UseConsistentIndentationIndentationSize     = 4

        UseConsistentWhitespaceEnable               = $true
        UseConsistentWhitespaceCheckInnerBrace      = $true
        UseConsistentWhitespaceCheckOpenBrace       = $true
        UseConsistentWhitespaceCheckOpenParen       = $true
        UseConsistentWhitespaceCheckOperator        = $true
        UseConsistentWhitespaceCheckPipe            = $true
        UseConsistentWhitespaceCheckSeparator       = $true

        AlignAssignmentStatementEnable              = $true
        AlignAssignmentStatementCheckHashtable      = $true

        UseCorrectCasingEnable                      = $true
    }
    # format PSD1 and PSM1 files when merging into a single file
    # enable formatting is not required as Configuration is provided
    New-ConfigurationFormat -ApplyTo 'OnMergePSM1', 'OnMergePSD1' -Sort None @ConfigurationFormat
    # format PSD1 and PSM1 files within the module
    # enable formatting is required to make sure that formatting is applied (with default settings)
    New-ConfigurationFormat -ApplyTo 'DefaultPSD1', 'DefaultPSM1' -EnableFormatting -Sort None
    # when creating PSD1 use special style without comments and with only required parameters
    New-ConfigurationFormat -ApplyTo 'DefaultPSD1', 'OnMergePSD1' -PSD1Style 'Minimal'
    # configuration for documentation, at the same time it enables documentation processing
    New-ConfigurationDocumentation -Enable:$false -StartClean -UpdateWhenNew -PathReadme 'Docs\Readme.md' -Path 'Docs'

    New-ConfigurationImportModule -ImportSelf

    New-ConfigurationBuild -Enable:$true -SignModule -MergeModuleOnBuild -MergeFunctionsFromApprovedModules -CertificateThumbprint '483292C9E317AA13B07BB7A96AE9D1A5ED9E7703'

    New-ConfigurationArtefact -Type Unpacked -Enable -Path "$PSScriptRoot\..\Artefacts\Unpacked" -ModulesPath "$PSScriptRoot\..\Artefacts\Unpacked\Modules" -RequiredModulesPath "$PSScriptRoot\..\Artefacts\Unpacked\Modules" -AddRequiredModules
    New-ConfigurationArtefact -Type Packed -Enable -Path "$PSScriptRoot\..\Artefacts\Packed" -ArtefactName '<ModuleName>.v<ModuleVersion>.zip'

    # global options for publishing to github/psgallery
    #New-ConfigurationPublish -Type PowerShellGallery -FilePath 'C:\Support\Important\PowerShellGalleryAPI.txt' -Enabled:$true
    #New-ConfigurationPublish -Type GitHub -FilePath 'C:\Support\Important\GitHubAPI.txt' -UserName 'EvotecIT' -Enabled:$true
}