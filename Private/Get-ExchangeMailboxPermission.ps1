function Get-ExchangeMailboxPermission {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $ReversedPermissions,
        [System.Collections.IDictionary] $CacheType,
        $Mailbox,
        [switch] $Local,
        [switch] $ExpandGroupMembership
    )

    if ($Local) {
        $Permissions = Get-LocalMailboxPermission -Identity $Mailbox.Alias -ErrorAction Stop -Verbose:$false
    } else {
        $Permissions = Get-MailboxPermission -Identity $Mailbox.Alias -ErrorAction Stop -Verbose:$false
    }

    foreach ($Permission in $Permissions) {
        if ($Permission.Deny -eq $false) {
            if ($Permission.User -ne 'NT AUTHORITY\SELF') {
                #if ($CacheType[$Mailbox.Alias] -eq 'On-Premises Mailbox') {
                #    $UserSplit = $Permission.User.Split("\")
                #    $CurrentUser = $UserSplit[1]
                #} elseif ($CacheType[$Mailbox.Alias] -eq 'Online Mailbox') {
                #    $CurrentUser = $CacheNames[$Permission.User]
                #}
                #if ($CurrentUser) {
                foreach ($Right in $Permission.AccessRights) {
                    if ($Right -like 'FullAccess*') {
                        if ($Permission.User -like "*@*") {
                            if ($CacheType[$Mailbox.Alias] -eq 'On-Premises Mailbox') {
                                $UserSplit = $Permission.User.Split("\")
                                $CurrentUser = $UserSplit[1]
                            } elseif ($CacheType[$Mailbox.Alias] -eq 'Online Mailbox') {
                                $CurrentUser = $CacheNames[$Permission.User]
                            }
                            if ($CurrentUser) {
                                if (-not $ReversedPermissions[$CurrentUser]) {
                                    $ReversedPermissions[$CurrentUser] = [ordered] @{
                                        FullAccess   = [System.Collections.Generic.List[string]]::new()
                                        SendAs       = [System.Collections.Generic.List[string]]::new()
                                        SendOnBehalf = [System.Collections.Generic.List[string]]::new()
                                    }
                                }
                                $ReversedPermissions[$CurrentUser].FullAccess.Add($Mailbox.Alias)
                            }
                        } else {
                            if ($ExpandGroupMembership) {
                                $GroupMembers = Get-ExchangeMembersRecursive -Identity $Permission.User -ErrorAction SilentlyContinue -Verbose:$false -AccessRights 'FullAccess' -Local:$Local.IsPresent
                                foreach ($Member in $GroupMembers) {
                                    $CurrentUser = $CacheNames[$Member.Trustee]
                                    if ($CurrentUser) {
                                        if (-not $ReversedPermissions[$CurrentUser]) {
                                            $ReversedPermissions[$CurrentUser] = [ordered] @{
                                                FullAccess   = [System.Collections.Generic.List[string]]::new()
                                                SendAs       = [System.Collections.Generic.List[string]]::new()
                                                SendOnBehalf = [System.Collections.Generic.List[string]]::new()
                                            }
                                        }
                                        $ReversedPermissions[$CurrentUser].FullAccess.Add($Mailbox.Alias)
                                    }
                                }
                            }
                        }


                    }
                }
                #} else {
                #   Write-Warning -Message "Unable to process $($Permission.User) for $($Mailbox.Alias)"
                ##}
            }
        }
    }
}