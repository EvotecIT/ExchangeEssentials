function ConvertTo-ExchangeAccessRights {
    [CmdletBinding()]
    Param (
        [parameter(Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)][PSObject] $Object,
        [string] $Identity,
        [string] $AccessRights = 'FullAccess'
    )
    Process {
        [PSCustomObject]@{
            Identity          = $Identity
            Trustee           = $Object.PrimarySmtpAddress
            AccessControlType = 'Allow'
            AccessRights      = @($AccessRights)
            IsInherited       = $false
            InheritanceType   = 'None'
        }
    }
}