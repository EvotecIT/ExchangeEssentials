function Reset-ExchangeEssentialsStatus {
    [cmdletBinding()]
    param(

    )
    if (-not $Script:DefaultTypes) {
        $Script:DefaultTypes = foreach ($T in $Script:ExchangeEssentialsConfiguration.Keys) {
            if ($Script:ExchangeEssentialsConfiguration[$T].Enabled) {
                $T
            }
        }
    } else {
        foreach ($T in $Script:ExchangeEssentialsConfiguration.Keys) {
            if ($Script:ExchangeEssentialsConfiguration[$T]) {
                $Script:ExchangeEssentialsConfiguration[$T]['Enabled'] = $false
            }
        }
        foreach ($T in $Script:DefaultTypes) {
            if ($Script:ExchangeEssentialsConfiguration[$T]) {
                $Script:ExchangeEssentialsConfiguration[$T]['Enabled'] = $true
            }
        }
    }
}