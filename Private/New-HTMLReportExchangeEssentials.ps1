function New-HTMLReportExchangeEssentials {
    [cmdletBinding()]
    param(
        [Array] $Type,
        [switch] $Online,
        [switch] $HideHTML,
        [string] $FilePath
    )

    New-HTML -Author 'Przemysław Kłys' -TitleText 'ExchangeEssentials Report' {
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

        if ($Type.Count -eq 1) {
            foreach ($T in $Script:ExchangeEssentialsConfiguration.Keys) {
                if ($Script:ExchangeEssentialsConfiguration[$T].Enabled -eq $true) {
                    if ($Script:ExchangeEssentialsConfiguration[$T]['Summary']) {
                        $Script:Reporting[$T]['Summary'] = Invoke-Command -ScriptBlock $Script:ExchangeEssentialsConfiguration[$T]['Summary']
                    }
                    & $Script:ExchangeEssentialsConfiguration[$T]['Solution']
                }
            }
        } else {
            foreach ($T in $Script:ExchangeEssentialsConfiguration.Keys) {
                if ($Script:ExchangeEssentialsConfiguration[$T].Enabled -eq $true) {
                    if ($Script:ExchangeEssentialsConfiguration[$T]['Summary']) {
                        $Script:Reporting[$T]['Summary'] = Invoke-Command -ScriptBlock $Script:ExchangeEssentialsConfiguration[$T]['Summary']
                    }
                    New-HTMLTab -Name $Script:ExchangeEssentialsConfiguration[$T]['Name'] {
                        & $Script:ExchangeEssentialsConfiguration[$T]['Solution']
                    }
                }
            }
        }
    } -Online:$Online.IsPresent -ShowHTML:(-not $HideHTML) -FilePath $FilePath
}