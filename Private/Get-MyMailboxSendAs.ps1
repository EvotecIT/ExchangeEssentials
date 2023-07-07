function Get-MyMailboxSendAs {
    <#
    .SYNOPSIS
    Short function that returns Send-As permissions for mailbox.

    .DESCRIPTION
    Function that returns Send-As permissions for mailbox. It's replacement of Get-ADPermission cmdlet that is very slow and inefficient.

    .PARAMETER ADUser
    Active Directory user object

    .PARAMETER Identity
    DistinguishedName of mailbox

    .EXAMPLE
    Get-ADUser -Identity 'przemyslaw.klys' -Properties NtsecurityDescriptor | Get-MyMailboxSendAs

    .EXAMPLE
    $Mailbox = Get-Mailbox -Identity 'przemyslaw.klys'
    $Test = Get-MyMailboxSendAs -Identity $Mailbox.DistinguishedName
    $Test | Format-Table

    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param(
        [parameter(Mandatory, ParameterSetName = 'ADUser', ValueFromPipeline, ValueFromPipelineByPropertyName)][Microsoft.ActiveDirectory.Management.ADAccount] $ADUser,
        [parameter(Mandatory, ParameterSetName = 'Identity', ValueFromPipeline, ValueFromPipelineByPropertyName )][string] $Identity
    )
    process {
        if ($ADUser) {
            if (-not $ADUser.NtsecurityDescriptor) {
                Write-Warning -Message "Get-MyMailboxSendAs - Identity $($ADUser.SamAccountName) does not have ntSecurityDescriptor attribute. Please provide one or use Identity parameter."
                return
            }
        } else {
            if (-not $Script:ForestInformation) {
                $Script:ForestInformation = Get-WinADForestDetails
            }
            $DomainName = ConvertFrom-DistinguishedName -DistinguishedName $Identity -ToDomainCN
            $QueryServer = $Script:ForestInformation['QueryServers'][$DomainName].HostName[0]
            $ADUser = Get-ADUser -Identity $Identity -Properties ntSecurityDescriptor -Server $QueryServer
        }
        $ExtendedPermissions = foreach ($Permission in $ADUser.NtsecurityDescriptor.Access) {
            # filter out extended right and object type (Send-As)
            if ($Permission.ActiveDirectoryRights -eq 'ExtendedRight' -and $Permission.objectType -eq "ab721a54-1e2f-11d0-9819-00aa0040529b") {
                [PSCustomObject] @{
                    User         = $Permission.IdentityReference
                    Identity     = if ($ADUser.CanonicalName) { $ADUser.CanonicalName } else { $ADUser.DistinguishedName }
                    AccessRights = 'Send-As'
                    Deny         = $Permission.AccessControlType -ne 'Allow'
                    IsInherited  = $Permission.IsInherited
                }
            }
        }
        $ExtendedPermissions
    }
}