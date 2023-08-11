﻿function New-HTMLReportExchangeEssentialsWithSplit {
    [cmdletBinding()]
    param(
        [Array] $Type,
        [switch] $Online,
        [switch] $HideHTML,
        [string] $FilePath,
        [string] $CurrentReport
    )

    # Split reports into multiple files for easier viewing
    $DateName = $(Get-Date -f yyyy-MM-dd_HHmmss)
    $FileName = [io.path]::GetFileNameWithoutExtension($FilePath)
    $DirectoryName = [io.path]::GetDirectoryName($FilePath)

    foreach ($T in $Script:ExchangeEssentialsConfiguration.Keys) {
        if ($Script:ExchangeEssentialsConfiguration[$T].Enabled -eq $true -and ((-not $CurrentReport) -or ($CurrentReport -and $CurrentReport -eq $T))) {
            $NewFileName = $FileName + '_' + $T + "_" + $DateName + '.html'
            $FilePath = [io.path]::Combine($DirectoryName, $NewFileName)

            New-HTML -Author 'Przemysław Kłys' -TitleText "ExchangeEssentials $CurrentReport Report" {
                New-HTMLTabStyle -BorderRadius 0px -TextTransform capitalize -BackgroundColorActive SlateGrey
                New-HTMLSectionStyle -BorderRadius 0px -HeaderBackGroundColor Grey -RemoveShadow
                New-HTMLPanelStyle -BorderRadius 0px
                New-HTMLTableOption -DataStore JavaScript -BoolAsString -ArrayJoinString ', ' -ArrayJoin

                New-HTMLHeader {
                    New-HTMLSection -Invisible {
                        New-HTMLSection {
                            New-HTMLText -Text "Report generated on $(Get-Date)" -Color Blue
                        } -JustifyContent flex-start -Invisible
                        New-HTMLSection {
                            New-HTMLText -Text "ExchangeEssentials - $($Script:Reporting['Version'])" -Color Blue
                        } -JustifyContent flex-end -Invisible
                    }
                }
                if ($Script:ExchangeEssentialsConfiguration[$T]['Summary']) {
                    $Script:Reporting[$T]['Summary'] = Invoke-Command -ScriptBlock $Script:ExchangeEssentialsConfiguration[$T]['Summary']
                }
                & $Script:ExchangeEssentialsConfiguration[$T]['Solution']
            } -Online:$Online.IsPresent -ShowHTML:(-not $HideHTML) -FilePath $FilePath
        }
    }
}