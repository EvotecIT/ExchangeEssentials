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

        if ($Script:Reporting['Mailbox']['Data'] -is [System.Collections.IDictionary]) {
            New-HTMLTabPanel {
                New-HTMLTable -DataTable $Script:Reporting['Mailbox']['Data'][$Domain] -Filtering {

                }
            }
        }
    }
}