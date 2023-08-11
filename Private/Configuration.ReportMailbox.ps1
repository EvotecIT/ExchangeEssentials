$Script:ReportMailbox = [ordered] @{
    Name       = 'Mailbox'
    Enabled    = $true
    Execute    = {
        Get-MyMailbox
    }
    Processing = {

    }
    Summary    = {

    }
    Variables  = @{

    }
    Solution   = {
        New-HTMLTable -DataTable $Script:Reporting['Mailbox']['Data'] -Filtering {

        }
    }
}