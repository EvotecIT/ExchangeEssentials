Import-Module .\ExchangeEssentials.psd1 -Force

$Test = Get-MyMailbox -Verbose -IncludeMailUsers
$Test | Format-Table