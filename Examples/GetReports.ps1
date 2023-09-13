Import-Module .\ExchangeEssentials.psd1 -Force

Invoke-ExchangeEssentials -FilePath $PSScriptRoot\Reports\ExchangeEssentialsReport.html -Online -Type MailboxProblems