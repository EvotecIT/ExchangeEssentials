function Get-MyMailboxMembers {
    [alias('Get-ExchangeMembersRecursive', 'Get-MyMailboxMembersRecursive')]
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory, ValueFromPipeline)] $Identity,
        [switch] $Local,
        [Parameter(DontShow)][int] $Nesting = -1,
        [Parameter(DontShow)][switch] $Nested
    )
    Begin {
        $IgnoreAccounts = @(
            'Everyone'
            'NT AUTHORITY\Authenticated Users'
            "Discovery Management"
        )
        if ($Nesting -eq -1) {
            $GroupCache = [ordered] @{}
        }
    }
    Process {
        $Nesting++
        $MatchRegex = [Regex]::Matches($Identity, "S-\d-\d+-(\d+-|){1,14}\d+")
        if ($MatchRegex.Success) {
            # do nothing
        } elseif ($Identity -in $IgnoreAccounts) {
            # do nothing
        } else {
            if (-not $GroupCache[$Identity]) {
                $GroupCache[$Identity] = $true
            } else {
                Write-Verbose -Message "Get-MyMailboxMembers - Identity '$Identity' already processed. Circular reference detected."
                return
                # Write-Color "Identity '$Identity' already processed." -Color Red
                #return
            }
            Write-Verbose -Message "Processing '$Identity'"
            if ($Local) {
                $Members = try {
                    Get-LocalDistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction Stop
                } catch {
                    try {
                        Get-LocalDynamicDistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction Stop
                    } catch {
                        Write-Warning -Message "Get-MyMailboxMembers - Identity '$Identity' is not a valid group."
                    }
                }
            } else {
                $Members = try {
                    Get-DistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction Stop
                } catch {
                    try {
                        Get-DynamicDistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction Stop
                    } catch {
                        Write-Warning -Message "Get-MyMailboxMembers - Identity '$Identity' is not a valid group."
                    }
                }
            }
            foreach ($Member in $Members) {
                switch ($Member) {
                    { $_.RecipientType -notlike "*Group*" } {
                        $_
                    }
                    Default {
                        if ($_.primarysmtpAddress) {
                            Get-MyMailboxMembers -Identity $_.primarysmtpAddress -Local:$Local.IsPresent -Nested -Nesting $Nesting
                        }
                    }
                }
            }
        }
    }
    End {
        if (-not $Nested) {
            #Write-Color -Text "Get-MyMailboxMembers - Clearing GroupCache" -Color Red
            $GroupCache = $Null
        }
    }
}