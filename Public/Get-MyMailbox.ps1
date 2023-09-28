﻿function Get-MyMailbox {
    [CmdletBinding()]
    param(
        [switch] $IncludeStatistics,
        [switch] $IncludeCAS,
        [switch] $IncludeMailUsers,
        [switch] $Local,
        [int] $LimitProcessing
    )

    $ReversedPermissions = [ordered] @{}
    $CacheMailbox = [ordered] @{}
    $CacheCasMailbox = [ordered] @{}
    $CacheNames = [ordered] @{}
    $FinalOutput = [ordered] @{}
    $CacheType = [ordered] @{}

    $CacheContacts = [ordered] @{}
    $CacheContactsLocal = [ordered] @{}
    $CacheRemoteDomains = [ordered] @{}
    $CacheRecipientPermissions = [ordered] @{}

    Write-Verbose -Message 'Get-MyMailbox - Getting Mailboxes'


    $Mailboxes = @(
        if ($Local) {
            $TimeLog = Start-TimeLog
            Write-Verbose -Message 'Get-MyMailbox - Getting Mailboxes (Local)'
            try {
                $LocalMailboxes = Get-LocalMailbox -ResultSize unlimited -ErrorAction Stop -Verbose:$false
                $EndTimeLog = Stop-TimeLog -Time $TimeLog -Option OneLiner
                Write-Verbose -Message "Get-MyMailbox - Getting Mailboxes (Local) took $($EndTimeLog)"
            } catch {
                $EndTimeLog = Stop-TimeLog -Time $TimeLog -Option OneLiner
                Write-Verbose -Message "Get-MyMailbox - Getting Mailboxes (Local) took $($EndTimeLog)"
                Write-Warning -Message "Get-MyMailbox - Unable to get Mailboxes. Error: $($_.Exception.Message.Replace("`r`n", " "))"
                return
            }
            #  $TimeLog = Start-TimeLog
            foreach ($Mailbox in $LocalMailboxes) {
                $CacheNames[$Mailbox.UserPrincipalName] = $Mailbox.Alias
                $CacheNames[$Mailbox.Identity] = $Mailbox.Alias
                $CacheNames[$Mailbox.Alias] = $Mailbox.Alias
                $CacheType[$Mailbox.Alias] = 'On-Premises Mailbox'
                $Mailbox
            }
            #  $EndTimeLog = Stop-TimeLog -Time $TimeLog -Option OneLiner
            # Write-Verbose -Message "Get-MyMailbox - Processing Mailboxes ($($LocalMailboxes.Count)) (Local) took $($EndTimeLog)"
        }
        Write-Verbose -Message 'Get-MyMailbox - Getting Mailboxes (Online)'
        $TimeLog = Start-TimeLog
        try {
            $OnlineMailboxes = Get-EXOMailbox -Properties GrantSendOnBehalfTo, ForwardingSmtpAddress, RecipientTypeDetails, SamAccountName, WhenCreated, WhenMailboxCreated, HiddenFromAddressListsEnabled, ForwardingAddress -ResultSize unlimited -ErrorAction Stop -Verbose:$false
            $EndTimeLog = Stop-TimeLog -Time $TimeLog -Option OneLiner
            Write-Verbose -Message "Get-MyMailbox - Getting Mailboxes (Online) took $($EndTimeLog) seconds"
        } catch {
            $EndTimeLog = Stop-TimeLog -Time $TimeLog -Option OneLiner
            Write-Verbose -Message "Get-MyMailbox - Getting Mailboxes (Online) took $($EndTimeLog) seconds"
            Write-Warning -Message "Get-MyMailbox - Unable to get Mailboxes. Error: $($_.Exception.Message.Replace("`r`n", " "))"
            return
        }
        #$TimeLog = Start-TimeLog
        foreach ($Mailbox in $OnlineMailboxes) {
            $CacheNames[$Mailbox.UserPrincipalName] = $Mailbox.Alias
            $CacheNames[$Mailbox.Identity] = $Mailbox.Alias
            $CacheNames[$Mailbox.Alias] = $Mailbox.Alias
            $CacheType[$Mailbox.Alias] = 'Online Mailbox'
            $Mailbox
        }


        if ($IncludeMailUsers) {
            Write-Verbose -Message 'Get-MyMailbox - Getting MailUsers (Online)'
            $TimeLog = Start-TimeLog
            try {
                $MailUsersOnline = Get-MailUser -ResultSize unlimited -ErrorAction Stop -Verbose:$false
                foreach ($M in $MailUsersOnline) {
                    $CacheNames[$M.UserPrincipalName] = $Mailbox.Alias
                    $CacheNames[$M.Identity] = $Mailbox.Alias
                    $CacheNames[$M.Alias] = $Mailbox.Alias
                    $CacheType[$M.Alias] = 'MailUser'
                    $M
                }
            } catch {
                Write-Warning -Message "Get-MyMailbox - Unable to get MailUsers. Error: $($_.Exception.Message.Replace("`r`n", " "))"
                return
            }
            $EndTimeLog = Stop-TimeLog -Time $TimeLog -Option OneLiner
            Write-Verbose -Message "Get-MyMailbox - Getting MailUsers (Online) took $($EndTimeLog) seconds"
        }

        #$EndTimeLog = Stop-TimeLog -Time $TimeLog -Option OneLiner
        #Write-Verbose -Message "Get-MyMailbox - Processing Mailboxes ($($OnlineMailboxes.Count)) (Online) took $($EndTimeLog)"
    )

    if (-not $RecipientPermissions) {
        Write-Verbose -Message 'Get-MyMailbox - Getting RecipientPermission'
        try {
            $RecipientPermissions = Get-EXORecipientPermission -ResultSize Unlimited -ErrorAction Stop -Verbose:$false
        } catch {
            Write-Warning -Message "Get-MyMailbox - Unable to get Recipient Permissions. Error: $($_.Exception.Message.Replace("`r`n", " "))"
            return
        }
    }
    if (-not $ContactsLocal) {
        if ($Local) {
            Write-Verbose -Message 'Get-MyMailbox - Getting Local Mail Contacts'
            try {
                [Array] $ContactsLocal = Get-LocalMailContact -ResultSize Unlimited -ErrorAction Stop -Verbose:$false
            } catch {
                Write-Warning -Message "Get-MyMailbox - Unable to get Local Mail Contacts. Error: $($_.Exception.Message.Replace("`r`n", " "))"
            }
        }
    }
    if (-not $Contacts) {
        Write-Verbose -Message 'Get-MyMailbox - Getting Mail Contacts'
        try {
            [Array] $Contacts = Get-MailContact -ResultSize Unlimited -ErrorAction Stop -Verbose:$false
        } catch {
            Write-Warning -Message "Get-MyMailbox - Unable to get Mail Contacts. Error: $($_.Exception.Message.Replace("`r`n", " "))"
            return
        }
    }
    if (-not $RemoteDomains) {
        try {
            $RemoteDomains = @(
                Write-Verbose -Message 'Get-MyMailbox - Getting Remote Domains'
                Get-RemoteDomain -ResultSize Unlimited -ErrorAction Stop -Verbose:$false
                if ($Local) {
                    Write-Verbose -Message 'Get-MyMailbox - Getting Local Remote Domains'
                    Get-LocalRemoteDomain -ErrorAction Stop -Verbose:$false
                }
            )
        } catch {
            Write-Warning -Message "Get-MyMailbox - Unable to get Remote Domains. Error: $($_.Exception.Message.Replace("`r`n", " "))"
            return
        }
    }
    foreach ($C in $Contacts) {
        $CacheContacts[$C.Identity] = $C
    }
    foreach ($C in $ContactsLocal) {
        $CacheContactsLocal[$C.Identity] = $C
    }
    foreach ($Domain in $RemoteDomains) {
        $CacheRemoteDomains[$Domain.DomainName] = $Domain
    }
    foreach ($RecipientPermission in $RecipientPermissions) {
        if ($RecipientPermission.Trustee -ne 'NT AUTHORITY\SELF') {
            if (-not $CacheRecipientPermissions[$RecipientPermission.Identity]) {
                $CacheRecipientPermissions[$RecipientPermission.Identity] = [System.Collections.Generic.List[PSCustomobject]]::new()
            }
            $CacheRecipientPermissions[$RecipientPermission.Identity].Add($RecipientPermission)
        }
    }
    if ($IncludeCAS) {
        Write-Verbose -Message 'Get-MyMailbox - Getting CAS Mailboxes'
        try {
            if (-not $CasMailboxes) {
                $CasMailboxes = @(
                    Write-Verbose -Message 'Get-MyMailbox - Getting CAS Mailboxes (Online)'
                    Get-CasMailbox -ResultSize unlimited -ErrorAction Stop -Verbose:$false
                    if ($Local) {
                        Write-Verbose -Message 'Get-MyMailbox - Getting CAS Mailboxes (Local)'
                        Get-LocalCasMailbox -ResultSize unlimited -ErrorAction Stop -Verbose:$false
                    }
                )
            }
        } catch {
            Write-Warning -Message "Get-MyMailbox - Unable to get CasMailboxes. Error: $($_.Exception.Message.Replace("`r`n", " "))"
            return
        }
        foreach ($Mailbox in $CasMailboxes) {
            $CacheCasMailbox[$Mailbox.Identity] = $Mailbox
        }
    }
    $Count = 0

    $FilterdMailboxes = @(
        if ($LimitProcessing) {
            $LocalMailboxes | Select-Object -First $LimitProcessing
            $OnlineMailboxes | Select-Object -First $LimitProcessing
        } else {
            $Mailboxes
        }
    )
    foreach ($Mailbox in $FilterdMailboxes) {
        $Count++
        Write-Verbose -Message "Processing Mailbox $Count/$($Mailboxes.Count) - $($Mailbox.Alias) / $($Mailbox.UserPrincipalName) / $($Mailbox.DisplayName)"
        $TimeLog = Start-TimeLog
        $CacheMailbox[$Mailbox.Alias] = [ordered] @{
            Mailbox = $Mailbox
        }
        #$TimeLogPermissions = Start-TimeLog
        if ($CacheType[$Mailbox.Alias] -eq 'On-Premises Mailbox') {
            try {
                Write-Verbose -Message "Get-MyMailbox - Getting MailboxPermissions for $($Mailbox.Alias) - Local"
                $CacheMailbox[$Mailbox.Alias].MailboxPermissions = Get-LocalMailboxPermission -Identity $Mailbox.Alias -ErrorAction Stop -Verbose:$false
            } catch {
                Write-Warning -Message "Get-MyMailbox - Unable to get MailboxPermissions for $($Mailbox.Alias). Error: $($_.Exception.Message.Replace("`r`n", " "))"
            }
        } elseif ($CacheType[$Mailbox.Alias] -eq 'Online Mailbox') {
            try {
                Write-Verbose -Message "Get-MyMailbox - Getting MailboxPermissions for $($Mailbox.Alias) - Online"
                $CacheMailbox[$Mailbox.Alias].MailboxPermissions = Get-MailboxPermission -Identity $Mailbox.Alias -ErrorAction Stop -Verbose:$false
            } catch {
                Write-Warning -Message "Get-MyMailbox - Unable to get MailboxPermissions for $($Mailbox.Alias). Error: $($_.Exception.Message.Replace("`r`n", " "))"
            }
        }
        #$TimeLogPermissionsEnd = Stop-TimeLog -Time $TimeLogPermissions
        #$TimeLogRecipient = Start-TimeLog
        if ($CacheType[$Mailbox.Alias] -eq 'On-Premises Mailbox') {
            try {
                Write-Verbose -Message "Get-MyMailbox - Getting MailboxADPermissions for $($Mailbox.Alias) - Local"
                $CacheMailbox[$Mailbox.Alias].MailboxRecipientPermissions = Get-MyMailboxSendAs -Identity $Mailbox.DistinguishedName #Get-LocalADPermission -Identity $Mailbox.Identity -ErrorAction Stop
            } catch {
                Write-Warning -Message "Get-MyMailbox - Unable to get MailboxADPermissions for $($Mailbox.Alias). Error: $($_.Exception.Message.Replace("`r`n", " "))"
            }
        } elseif ($CacheType[$Mailbox.Alias] -eq 'Online Mailbox') {
            if ($CacheRecipientPermissions[$Mailbox.Identity] -and $CacheRecipientPermissions[$Mailbox.Identity].Count -gt 0) {
                $CacheMailbox[$Mailbox.Alias].MailboxRecipientPermissions = $CacheRecipientPermissions[$Mailbox.Identity]
            }
            # try {
            #     Write-Verbose -Message "Get-MyMailbox - Getting MailboxRecipientPermissions for $($Mailbox.Alias) - Online"
            #     $CacheMailbox[$Mailbox.Alias].MailboxRecipientPermissions = Get-RecipientPermission -Identity $Mailbox.Alias -ErrorAction Stop
            # } catch {
            #     Write-Warning -Message "Get-MyMailbox - Unable to get MailboxRecipientPermissions for $($Mailbox.Alias). Error: $($_.Exception.Message.Replace("`r`n", " "))"
            # }
        }
        #$TimeLogRecipientEnd = Stop-TimeLog -Time $TimeLogRecipient
        #$TImeLogStats = Start-TimeLog
        if ($IncludeStatistics) {
            if ($CacheType[$Mailbox.Alias] -eq 'On-Premises Mailbox') {
                $CacheMailbox[$Mailbox.Alias]['Statistics'] = Get-LocalMailboxStatistics -Identity $Mailbox.Alias
            } elseif ($CacheType[$Mailbox.Alias] -eq 'Online Mailbox') {
                $CacheMailbox[$Mailbox.Alias]['Statistics'] = Get-MailboxStatistics -Identity $Mailbox.Alias
            }
            if ($Mailbox.ArchiveDatabase) {
                try {
                    if ($CacheType[$Mailbox.Alias] -eq 'On-Premises Mailbox') {
                        $Archive = Get-LocalMailboxStatistics -Identity ($Mailbox.Guid).ToString() -Archive -Verbose:$false -ErrorAction Stop
                    } elseif ($CacheType[$Mailbox.Alias] -eq 'Online Mailbox') {
                        $Archive = Get-MailboxStatistics -Identity ($Mailbox.Guid).ToString() -Archive -Verbose:$false -ErrorAction Stop
                    }
                    $CacheMailbox[$Mailbox.Alias]['StatisticsArchive'] = $Archive
                } catch {
                    Write-Warning -Message "Get-MyMailbox - Unable to get ArchiveStatistics for $($Mailbox.Alias). Error: $($_.Exception.Message.Replace("`r`n", " "))"
                }
            }
        }
        #$TimeLogStatsEnd = Stop-TimeLog -Time $TImeLogStats
        #$TimeLogProcessing = Start-TimeLog
        foreach ($Permission in $CacheMailbox[$Mailbox.Alias].MailboxPermissions) {
            if ($Permission.Deny -eq $false) {
                if ($Permission.User -ne 'NT AUTHORITY\SELF') {
                    if ($CacheType[$Mailbox.Alias] -eq 'On-Premises Mailbox') {
                        $UserSplit = $Permission.User.Split("\")
                        $CurrentUser = $UserSplit[1]
                    } elseif ($CacheType[$Mailbox.Alias] -eq 'Online Mailbox') {
                        $CurrentUser = $CacheNames[$Permission.User]
                    }
                    if ($CurrentUser) {
                        foreach ($Right in $Permission.AccessRights) {
                            if ($Right -like 'FullAccess*') {
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
                    } else {
                        # Write-Warning -Message "Unable to process $($Permission.User) for $($Mailbox.Alias)"
                    }
                }
            }
        }
        foreach ($Permission in $CacheMailbox[$Mailbox.Alias].MailboxRecipientPermissions) {
            if ($CacheType[$Mailbox.Alias] -eq 'On-Premises Mailbox') {
                if ($Permission.Deny -eq $false -and $Permission.Inherited -eq $false) {
                    if ($Permission.User -ne 'NT AUTHORITY\SELF') {
                        $UserSplit = $Permission.User.Split("\")
                        $CurrentUser = $UserSplit[1]
                        if ($CurrentUser) {
                            foreach ($Right in $Permission.AccessRights) {
                                if (($Right -like 'Send*')) {
                                    if (-not $ReversedPermissions[$CurrentUser]) {
                                        $ReversedPermissions[$CurrentUser] = [ordered] @{
                                            FullAccess   = [System.Collections.Generic.List[string]]::new()
                                            SendAs       = [System.Collections.Generic.List[string]]::new()
                                            SendOnBehalf = [System.Collections.Generic.List[string]]::new()
                                        }
                                    }
                                    $ReversedPermissions[$CurrentUser].SendAs.Add($Mailbox.Alias)
                                }
                            }
                        } else {
                            #Write-Warning -Message "Unable to process $($Permission.Trustee) for $($Mailbox.Alias)"
                        }
                    }
                }
            } else {
                if ($Permission.AccessControlType -eq 'Allow') {
                    if ($Permission.Trustee -ne 'NT AUTHORITY\SELF') {
                        $CurrentUser = $CacheNames[$Permission.Trustee]
                        if ($CurrentUser) {
                            foreach ($Right in $Permission.AccessRights) {
                                if (($Right -like 'Send*')) {
                                    if (-not $ReversedPermissions[$CurrentUser]) {
                                        $ReversedPermissions[$CurrentUser] = [ordered] @{
                                            FullAccess   = [System.Collections.Generic.List[string]]::new()
                                            SendAs       = [System.Collections.Generic.List[string]]::new()
                                            SendOnBehalf = [System.Collections.Generic.List[string]]::new()
                                        }
                                    }
                                    $ReversedPermissions[$CurrentUser].SendAs.Add($Mailbox.Alias)
                                }
                            }
                        } else {
                            #Write-Warning -Message "Unable to process $($Permission.Trustee) for $($Mailbox.Alias)"
                        }
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
                # Write-Warning -Message "Unable to process $($Permission) for $($Mailbox.Alias)"
            }
        }
        #$TimeLogProcessingEnd = Stop-TimeLog -Time $TimeLogProcessing
        $EndTimeLog = Stop-TimeLog -Time $TimeLog -Option OneLiner
        #Write-Verbose -Message "Processing Mailbox $Count/$($Mailboxes.Count) - $($Mailbox.Alias) / $($Mailbox.UserPrincipalName) / $($Mailbox.DisplayName) - [$TimeLogPermissionsEnd][$TimeLogRecipientEnd][$TimeLogStatsEnd][$TimeLogProcessingEnd]"
        Write-Verbose -Message "Processing Mailbox $Count/$($Mailboxes.Count) - $($Mailbox.Alias) / $($Mailbox.UserPrincipalName) / $($Mailbox.DisplayName) - [$EndTimeLog]"
    }
    foreach ($Alias in $ReversedPermissions.Keys) {
        if ($CacheMailbox[$Alias]) {
            $CacheMailbox[$Alias].FullAccess = $ReversedPermissions[$Alias].FullAccess
            $CacheMailbox[$Alias].SendAs = $ReversedPermissions[$Alias].SendAs
            $CacheMailbox[$Alias].SendOnBehalf = $ReversedPermissions[$Alias].SendOnBehalf
        }
    }
    $Count = 0
    foreach ($Mailbox in $FilterdMailboxes) {
        $Count++
        Write-Verbose -Message "Processing Mailbox $Count/$($FilterdMailboxes.Count) - $($Mailbox.Alias) / $($Mailbox.UserPrincipalName) / $($Mailbox.DisplayName)"
        if ($Mailbox.ForwardingAddress) {
            $Contact = $CacheContacts[$Mailbox.ForwardingAddress]
            $ContactLocal = $CacheContactsLocal[$Mailbox.ForwardingAddress]
            if ($Contact) {
                $ForwardAddress = Convert-ExchangeEmail -Emails $Contact.ExternalEmailAddress -RemovePrefix
                if ($ForwardAddress) {
                    $IsForward = $true
                } else {
                    $IsForward = $false
                }
            } elseif ($ContactLocal) {
                $ForwardAddress = Convert-ExchangeEmail -Emails $ContactLocal.ExternalEmailAddress -RemovePrefix
                if ($ForwardAddress) {
                    $IsForward = $true
                } else {
                    $IsForward = $false
                }
            } else {
                $ForwardAddress = $Mailbox.ForwardingAddress
                $IsForward = $true
            }
            $ForwardingType = 'Contact'
        } elseif ($Mailbox.ForwardingSmtpAddress) {
            $ForwardAddress = Convert-ExchangeEmail -Emails $Mailbox.ForwardingSmtpAddress -RemovePrefix
            if ($ForwardAddress) {
                $IsForward = $true
            } else {
                # this shouldn't happen
                $IsForward = 'Unknown'
            }
            $ForwardingType = 'SmtpAddress'
        } else {
            $ForwardAddress = $null
            $IsForward = $false
            $ForwardingType = 'None'
        }
        if ($ForwardAddress) {
            if ($ForwardAddress -like "*@*") {
                $SplitAddress = $ForwardAddress -split "@"
                $DomainName = $SplitAddress[1]
                $DomainName = $DomainName.Trim()
                if ($CacheRemoteDomains[$DomainName]) {
                    $ForwardingStatus = "Internal"
                } else {
                    $ForwardingStatus = "External"
                }
            } else {
                $ForwardingStatus = "Unknown"
            }
        } else {
            $ForwardingStatus = 'None'
        }


        $User = [ordered] @{
            DisplayName                   = $Mailbox.DisplayName
            Alias                         = $Mailbox.Alias
            UserPrincipalName             = $Mailbox.UserPrincipalName
            Enabled                       = -not $Mailbox.AccountDisabled
            Type                          = $CacheType[$Mailbox.Alias]
            TypeDetails                   = $Mailbox.RecipientTypeDetails
            PrimarySmtpAddress            = $Mailbox.PrimarySmtpAddress
            SamAccountName                = $Mailbox.SamAccountName
            ForwardingEnabled             = $IsForward
            ForwardingStatus              = $ForwardingStatus
            ForwardingType                = $ForwardingType
            ForwardingAddress             = $ForwardAddress

            FullAccess                    = $CacheMailbox[$Mailbox.Alias].FullAccess
            SendAs                        = $CacheMailbox[$Mailbox.Alias].SendAs
            SendOnBehalf                  = $CacheMailbox[$Mailbox.Alias].SendOnBehalf
            FullAccessCount               = $CacheMailbox[$Mailbox.Alias].FullAccess.Count
            SendAsCount                   = $CacheMailbox[$Mailbox.Alias].SendAs.Count
            SendOnBehalfCount             = $CacheMailbox[$Mailbox.Alias].SendOnBehalf.Count
            WhenCreated                   = $Mailbox.WhenCreated
            WhenMailboxCreated            = $Mailbox.WhenMailboxCreated
            HiddenFromAddressListsEnabled = $Mailbox.HiddenFromAddressListsEnabled
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
}