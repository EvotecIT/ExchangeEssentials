$Script:ReportMailboxProblems = [ordered] @{
    Name       = 'MailboxProblems'
    Enabled    = $true
    Execute    = {
        Get-MyMailboxProblems
    }
    Processing = {

    }
    Summary    = {

    }
    Variables  = @{

    }
    Solution   = {
        if ($Script:Reporting['MailboxProblems']['Data'] -is [System.Collections.IDictionary]) {
            New-HTMLTable -DataTable $Script:Reporting['MailboxProblems']['Data'][$Domain] -Filtering {

            }
        }
    }
}