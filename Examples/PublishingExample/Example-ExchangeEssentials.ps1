Import-Module $PSScriptRoot\Modules\ExchangeOnlineManagement\ExchangeOnlineManagement.psd1 -Force
Import-Module $PSScriptRoot\Modules\ExchangeEssentials\ExchangeEssentials.psd1 -Force

Get-MyMailbox -Verbose