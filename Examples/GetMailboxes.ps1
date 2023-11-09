Import-Module .\ExchangeEssentials.psd1 -Force

#Connect-ExchangeOnline

$Test = Get-MyMailbox -Verbose -IncludeMailUsers -SkipPermissions
$Test | Format-Table
$Test | Out-HtmlView -ScrollX -Filtering