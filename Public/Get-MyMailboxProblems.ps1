function Get-MyMailboxProblems {
    [CmdletBinding()]
    param(
        [switch] $Local
    )
    $Problems = [ordered] @{
        Data   = [ordered] @{
            ContactsLocal   = $Null
            ContactsOnline  = $Null
            MailboxesLocal  = $Null
            MailboxesOnline = $Null
        }
        Online = [ordered] @{
            BrokenDisplayName           = [System.Collections.Generic.List[PSCustomObject]]::new()
            DuplicateAlias              = [System.Collections.Generic.List[PSCustomObject]]::new()
            NoDatabase                  = [System.Collections.Generic.List[PSCustomObject]]::new()
            ContactMissingPrimarySmtp   = [System.Collections.Generic.List[PSCustomObject]]::new()
            ContactMissingExternalEmail = [System.Collections.Generic.List[PSCustomObject]]::new()
        }
        Local  = [ordered] @{
            BrokenDisplayName           = [System.Collections.Generic.List[PSCustomObject]]::new()
            DuplicateAlias              = [System.Collections.Generic.List[PSCustomObject]]::new()
            NoDatabase                  = [System.Collections.Generic.List[PSCustomObject]]::new()
            ContactMissingPrimarySmtp   = [System.Collections.Generic.List[PSCustomObject]]::new()
            ContactMissingExternalEmail = [System.Collections.Generic.List[PSCustomObject]]::new()
        }
        Both   = [ordered] @{
            DuplicateAlias   = [System.Collections.Generic.List[PSCustomObject]]::new()
            DuplicateAccount = [System.Collections.Generic.List[PSCustomObject]]::new()
        }
    }

    $PropertiesMailbox = @(
        'Identity', 'UserPrincipalName', 'Alias', 'DisplayName', 'Database', 'PrimarySmtpAddress', 'RecipientType', 'RecipientTypeDetails', 'ExchangeUserAccountControl'
    )

    if ($Local) {
        $ContactsLocal = Get-LocalMailContact -ResultSize Unlimited
        $MailboxesLocal = Get-LocalMailbox -ResultSize Unlimited | Select-Object -Property $PropertiesMailbox
    }

    $ContactsOnline = Get-LocalMailContact -ResultSize Unlimited
    $MailboxesOnline = Get-EXOMailbox -ResultSize Unlimited -Properties $PropertiesMailbox | Select-Object -Property $PropertiesMailbox

    foreach ($C in $ContactsLocal) {
        if (-not $C.PrimarySmtpAddress) {
            $Problems.Local.ContactMissingPrimarySmtp.Add($C)
        }
        if (-not $C.ExternalEmailAddress) {
            $Problems.Local.ContactMissingExternalEmail.Add($C)
        }
    }
    foreach ($C in $ContactsOnline) {
        if (-not $C.PrimarySmtpAddress) {
            $Problems.Online.ContactMissingPrimarySmtp.Add($C)
        }
        if (-not $C.ExternalEmailAddress) {
            $Problems.Online.ContactMissingExternalEmail.Add($C)
        }
    }

    foreach ($M in $MailboxesLocal) {
        if ($M.DisplayName.StartsWith(' ') -or $M.DisplayName.EndsWith(' ')) {
            $Problems.Local.BrokenDisplayName.Add($M)
        }
        if ($Null -eq $M.Database) {
            $Problems.Local.NoDatabase.Add($M)
        }
    }
    foreach ($M in $MailboxesOnline) {
        if ($M.DisplayName.StartsWith(' ') -or $M.DisplayName.EndsWith(' ')) {
            $Problems.Online.BrokenDisplayName.Add($M)
        }
        if ($Null -eq $M.Database) {
            $Problems.Online.NoDatabase.Add($M)
        }
    }

    $CacheAll = [ordered] @{}
    $Cache = [ordered] @{}
    foreach ($M in $MailboxesOnline) {
        if (-not $Cache[$M.Alias]) {
            $Cache[$M.Alias] = $M
        } else {
            if (-not $CacheAll[$M.Alias]) {
                $CacheAll[$M.Alias] = [System.Collections.Generic.List[PSCustomObject]]::new()
                $CacheAll[$M.Alias].Add($Cache[$M.Alias])
            }
            $CacheAll[$M.Alias].Add($M)
        }
    }
    foreach ($M in $CacheAll.Values) {
        $Problems.Online.DuplicateAlias.Add($M)
    }
    # lets reset for local
    $CacheAll = [ordered] @{}
    $Cache = [ordered] @{}
    foreach ($M in $MailboxesLocal) {
        if (-not $Cache[$M.Alias]) {
            $Cache[$M.Alias] = $M
        } else {
            if (-not $CacheAll[$M.Alias]) {
                $CacheAll[$M.Alias] = [System.Collections.Generic.List[PSCustomObject]]::new()
                $CacheAll[$M.Alias].Add($Cache[$M.Alias])
            }
            $CacheAll[$M.Alias].Add($M)
        }
    }
    foreach ($M in $CacheAll.Values) {
        $Problems.Local.DuplicateAlias.Add($M)
    }
    # both?
    $CacheAll = [ordered] @{}
    $Cache = [ordered] @{}
    $CacheUPN = [ordered] @{}
    $CacheAllUPN = [ordered] @{}
    foreach ($M in $MailboxesOnline + $MailboxesLocal) {
        # Duplicate Accounts
        if (-not $CacheUPN[$M.UserPrincipalName]) {
            $CacheUPN[$M.UserPrincipalName] = $M
        } else {
            if (-not $CacheAllUPN[$M.UserPrincipalName]) {
                $CacheAllUPN[$M.UserPrincipalName] = [System.Collections.Generic.List[PSCustomObject]]::new()
                $CacheAllUPN[$M.UserPrincipalName].Add($CacheUPN[$M.UserPrincipalName])
            }
            $CacheAllUPN[$M.UserPrincipalName].Add($M)
        }
        # Duplicate Aliases
        if (-not $Cache[$M.Alias]) {
            $Cache[$M.Alias] = $M
        } else {
            if (-not $CacheAll[$M.Alias]) {
                $CacheAll[$M.Alias] = [System.Collections.Generic.List[PSCustomObject]]::new()
                $CacheAll[$M.Alias].Add($Cache[$M.Alias])
            }
            $CacheAll[$M.Alias].Add($M)
        }
    }
    foreach ($M in $CacheAll.Values) {
        $Problems.Both.DuplicateAlias.Add($M)
    }
    foreach ($M in $CacheAllUPN.Values) {
        $Problems.Both.DuplicateAccount.Add($M)
    }

    if ($Local) {
        $Problems.Data.ContactsLocal = $ContactsLocal
        $Problems.Data.MailboxesLocal = $MailboxesLocal
    }
    $Problems.Data.ContactsOnline = $ContactsOnline
    $Problems.Data.MailboxesOnline = $MailboxesOnline

    $Problems
}