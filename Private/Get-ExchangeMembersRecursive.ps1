function Get-ExchangeMembersRecursive {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)] $Identity,
        [Parameter(Mandatory = $true)] [ValidateSet('FullAccess', 'SendAs', 'SendOnBehalf')] [string[]] $AccessRights,
        [switch] $Local
    )
    Begin {
        $IgnoreAccounts = @(
            'Everyone'
            'NT AUTHORITY\Authenticated Users'
        )
    }
    Process {
        $MatchRegex = [Regex]::Matches($Identity, "S-\d-\d+-(\d+-|){1,14}\d+")
        if ($MatchRegex.Success) {
            # do nothing
        } elseif ($Identity -in $IgnoreAccounts) {
            # do nothing
        } else {
            if ($Local) {
                $Members = try {
                    Get-LocalDistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction Stop
                } catch {
                    try {
                        Get-LocalDynamicDistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction Stop
                    } catch {
                        Write-Warning -Message "Get-ExchangeMembersRecursive - Identity '$Identity' is not a valid group."
                    }
                }
            } else {
                $Members = try {
                    Get-DistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction Stop
                } catch {
                    try {
                        Get-DynamicDistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction Stop
                    } catch {
                        Write-Warning -Message "Get-ExchangeMembersRecursive - Identity '$Identity' is not a valid group."
                    }
                }
            }
            foreach ($Member in $Members) {
                switch ($Member) {
                    { $_.RecipientType -notlike "*Group*" } {
                        [PSCustomObject]@{
                            Identity          = $_.Identity
                            Trustee           = $_.PrimarySmtpAddress
                            AccessControlType = 'Allow'
                            AccessRights      = @($AccessRights)
                            IsInherited       = $false
                            InheritanceType   = 'None'
                        }
                    }
                    Default {
                        if ($_.primarysmtpAddress) {
                            Get-ExchangeMembersRecursive -Identity $_.primarysmtpAddress -Local:$Local.IsPresent -AccessRights $AccessRights
                        }
                    }
                }
            }
        }
    }
}
