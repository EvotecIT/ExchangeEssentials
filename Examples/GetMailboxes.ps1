Import-Module .\ExchangeEssentials.psd1 -Force

#Connect-ExchangeOnline
Get-MyMailbox -IncludeStatistics -IncludeCAS | Out-HtmlView -ScrollX -DataStore JavaScript