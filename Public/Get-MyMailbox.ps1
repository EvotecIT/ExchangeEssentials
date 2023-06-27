function Get-MyMailbox {
    [CmdletBinding()]
    param(
        [switch] $IncludeStatistics,
        [switch] $IncludeCAS
    )

    $ReversedPermissions = [ordered] @{}
    $CacheMailbox = [ordered] @{}
    $CacheCasMailbox = [ordered] @{}
    $CacheNames = [ordered] @{}
    $FinalOutput = [ordered] @{}

    Write-Verbose -Message 'Get-MyMailbox - Getting Mailboxes'
    try {
        $Mailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop
    } catch {
        Write-Warning -Message "Get-MyMailbox - Unable to get Mailboxes. Error: $($_.Exception.Message.Replace("`r`n", " "))"
        return
    }
    if ($IncludeCAS) {
        Write-Verbose -Message 'Get-MyMailbox - Getting CAS Mailboxes'
        try {
            $CasMailboxes = Get-CasMailbox -ResultSize Unlimited -ErrorAction Stop
        } catch {
            Write-Warning -Message "Get-MyMailbox - Unable to get CasMailboxes. Error: $($_.Exception.Message.Replace("`r`n", " "))"
            return
        }
        foreach ($Mailbox in $CasMailboxes) {
            $CacheCasMailbox[$Mailbox.Identity] = $Mailbox
        }
    }
    foreach ($Mailbox in $Mailboxes) {
        $CacheNames[$Mailbox.UserPrincipalName] = $Mailbox.Alias
        $CacheNames[$Mailbox.Identity] = $Mailbox.Alias
        $CacheNames[$Mailbox.Alias] = $Mailbox.Alias
    }
    foreach ($Mailbox in $Mailboxes) {
        $CacheMailbox[$Mailbox.Alias] = [ordered] @{
            Mailbox = $Mailbox
        }
        try {

            $CacheMailbox[$Mailbox.Alias].MailboxPermissions = Get-MailboxPermission -Identity $Mailbox.Alias -ErrorAction Stop
        } catch {
            Write-Warning -Message "Get-MyMailbox - Unable to get MailboxPermissions for $($Mailbox.Alias). Error: $($_.Exception.Message.Replace("`r`n", " "))"
        }
        try {
            $CacheMailbox[$Mailbox.Alias].MailboxRecipientPermissions = Get-RecipientPermission -Identity $Mailbox.Alias -ErrorAction Stop
        } catch {
            Write-Warning -Message "Get-MyMailbox - Unable to get MailboxRecipientPermissions for $($Mailbox.Alias). Error: $($_.Exception.Message.Replace("`r`n", " "))"
        }

        if ($IncludeStatistics) {
            $CacheMailbox[$Mailbox.Alias]['Statistics'] = Get-MailboxStatistics -Identity $Mailbox.Alias
            if ($Mailbox.ArchiveDatabase) {
                try {
                    $Archive = Get-MailboxStatistics -Identity ($Mailbox.Guid).ToString() -Archive -Verbose:$false -ErrorAction Stop
                    $CacheMailbox[$Mailbox.Alias]['StatisticsArchive'] = $Archive
                } catch {
                    Write-Warning -Message "Get-MyMailbox - Unable to get ArchiveStatistics for $($Mailbox.Alias). Error: $($_.Exception.Message.Replace("`r`n", " "))"
                }
            }
        }
        foreach ($Permission in $CacheMailbox[$Mailbox.Alias].MailboxPermissions) {
            if ($Permission.Deny -eq $false) {
                if ($Permission.User -ne 'NT AUTHORITY\SELF') {
                    $CurrentUser = $CacheNames[$Permission.User]
                    if ($CurrentUser) {
                        if (-not $ReversedPermissions[$CurrentUser]) {
                            $ReversedPermissions[$CurrentUser] = [ordered] @{
                                FullAccess   = [System.Collections.Generic.List[string]]::new()
                                SendAs       = [System.Collections.Generic.List[string]]::new()
                                SendOnBehalf = [System.Collections.Generic.List[string]]::new()
                            }
                        }
                        if ($Permission.AccessRights -like '*FullAccess') {
                            $ReversedPermissions[$CurrentUser].FullAccess.Add($Mailbox.Alias)
                        }
                    } else {
                        Write-Warning -Message "Unable to process $($Permission.User) for $($Mailbox.Alias)"
                    }
                }
            }
        }
        foreach ($Permission in $CacheMailbox[$Mailbox.Alias].MailboxRecipientPermissions) {
            if ($Permission.AccessControlType -eq 'Allow') {
                if ($Permission.Trustee -ne 'NT AUTHORITY\SELF') {
                    $CurrentUser = $CacheNames[$Permission.Trustee]
                    if ($CurrentUser) {
                        if (-not $ReversedPermissions[$CurrentUser]) {
                            $ReversedPermissions[$CurrentUser] = [ordered] @{
                                FullAccess   = [System.Collections.Generic.List[string]]::new()
                                SendAs       = [System.Collections.Generic.List[string]]::new()
                                SendOnBehalf = [System.Collections.Generic.List[string]]::new()
                            }
                        }
                        if ($Permission.AccessRights -contains 'SendAs') {
                            $ReversedPermissions[$CurrentUser].SendAs.Add($Mailbox.Alias)
                        }
                    } else {
                        Write-Warning -Message "Unable to process $($Permission.Trustee) for $($Mailbox.Alias)"
                    }
                }
            }
        }
        foreach ($Permission in $CacheMailbox[$Mailbox.Alias].Mailbox.GrantSendOnBehalfTo) {
            $CurrentUser = $CacheNames[$Permission]
            if ($CurrentUser) {
                if (-not $ReversedPermissions[$CurrentUser]) {
                    $ReversedPermissions[$CurrentUser] = [ordered] @{
                        FullAccess   = [System.Collections.Generic.List[string]]::new()
                        SendAs       = [System.Collections.Generic.List[string]]::new()
                        SendOnBehalf = [System.Collections.Generic.List[string]]::new()
                    }
                }

                $ReversedPermissions[$CurrentUser].SendOnBehalf.Add($Mailbox.Alias)
            } else {
                Write-Warning -Message "Unable to process $($Permission) for $($Mailbox.Alias)"
            }
        }
    }
    foreach ($Alias in $ReversedPermissions.Keys) {
        $CacheMailbox[$Alias].FullAccess = $ReversedPermissions[$Alias].FullAccess
        $CacheMailbox[$Alias].SendAs = $ReversedPermissions[$Alias].SendAs
        $CacheMailbox[$Alias].SendOnBehalf = $ReversedPermissions[$Alias].SendOnBehalf
    }
    $Output = foreach ($Mailbox in $Mailboxes) {
        $User = [ordered] @{
            DisplayName                   = $Mailbox.DisplayName
            Alias                         = $Mailbox.Alias
            UserPrincipalName             = $Mailbox.UserPrincipalName
            Enabled                       = -not $Mailbox.AccountDisabled
            TypeDetails                   = $Mailbox.RecipientTypeDetails
            PrimarySmtpAddress            = $Mailbox.PrimarySmtpAddress
            SamAccountName                = $Mailbox.SamAccountName
            ExchangeUserAccountControl    = $Mailbox.ExchangeUserAccountControl
            #ForwardingSmtpAddress      = $Mailbox.ForwardingSmtpAddress
            FullAccess                    = $CacheMailbox[$Mailbox.Alias].FullAccess
            SendAs                        = $CacheMailbox[$Mailbox.Alias].SendAs
            SendOnBehalf                  = $CacheMailbox[$Mailbox.Alias].SendOnBehalf
            FullAccessCount               = $CacheMailbox[$Mailbox.Alias].FullAccess.Count
            SendAsCount                   = $CacheMailbox[$Mailbox.Alias].SendAs.Count
            SendOnBehalfCount             = $CacheMailbox[$Mailbox.Alias].SendOnBehalf.Count
            WhenCreated                   = $Mailbox.WhenCreated
            WhenMailboxCreated            = $Mailbox.WhenMailboxCreated
            HiddenFromAddressListsEnabled = $Mailbox.HiddenFromAddressListsEnabled
            #RecipientType                 = $Mailbox.RecipientType
        }
        if ($IncludeStatistics) {
            #$User['LastUserAccessTime'] = $CacheMailbox[$Mailbox.Alias].Statistics.LastUserAccessTime
            $User['LastLogonTime'] = $CacheMailbox[$Mailbox.Alias].Statistics.LastLogonTime
            $User['TotalItems'] = $CacheMailbox[$Mailbox.Alias].Statistics.ItemCount
            $User['TotalGB'] = Convert-ExchangeSize -Size $CacheMailbox[$Mailbox.Alias].Statistics.TotalItemSize -To GB
            $User['TotalArchiveItems'] = $CacheMailbox[$Mailbox.Alias].StatisticsArchive.ItemCount
            $User['TotalArchiveGB'] = Convert-ExchangeSize -Size $CacheMailbox[$Mailbox.Alias].StatisticsArchive.TotalItemSize -To GB
        }
        if ($IncludeCAS) {
            $CasProperties = @(
                'ActiveSyncEnabled'
                'OWAEnabled'
                'OWAforDevicesEnabled'
                'ECPEnabled'
                'PopEnabled'
                'PopMessageDeleteEnabled'
                'ImapEnabled'
                'MAPIEnabled'
                'MapiHttpEnabled'
                'UniversalOutlookEnabled'
                'OutlookMobileEnabled'
                'MacOutlookEnabled'
                'EwsEnabled'
                'OneWinNativeOutlookEnabled'
                'BulkMailEnabled'
                'SmtpClientAuthenticationDisabled'
            )
            if ($CacheCasMailbox[$Mailbox.Identity]) {
                foreach ($Property in $CasProperties) {
                    $User[$Property] = $CacheCasMailbox[$Mailbox.Alias].$Property
                }
            } else {
                foreach ($Property in $CasProperties) {
                    $User[$Property] = $null
                }
            }
        }
        $ConvertedUser = [PSCustomObject] $User

        $FinalOutput[$Mailbox.Alias] = $ConvertedUser
        $ConvertedUser
    }
    $Output
}