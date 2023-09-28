$Script:ReportMailboxProblems = [ordered] @{
    Name       = 'MailboxProblems'
    Enabled    = $true
    Execute    = {
        Get-MyMailboxProblems -Local
    }
    Processing = {

    }
    Summary    = {

    }
    Variables  = @{

    }
    Solution   = {
        if ($Script:Reporting['MailboxProblems']['Data'] -is [System.Collections.IDictionary]) {
            New-HTMLTabPanel -Theme forge {
                New-HTMLTab -Name "Duplicate Alias" {
                    New-HTMLSection -HeaderText 'Exchange Online - Duplicate Alias' {
                        New-HTMLTable -DataTable $Script:Reporting['MailboxProblems']['Data']['Online']['DuplicateAlias'] -Filtering
                    }
                    New-HTMLSection -HeaderText 'Exchange On-Premises - Duplicate Alias' {
                        New-HTMLTable -DataTable $Script:Reporting['MailboxProblems']['Data']['Local']['DuplicateAlias'] -Filtering
                    }
                    New-HTMLSection -HeaderText 'Exchange - Duplicate Alias (Together)' {
                        New-HTMLTable -DataTable $Script:Reporting['MailboxProblems']['Data']['Both']['DuplicateAlias'] -Filtering
                    }
                }
                New-HTMLTab -Name "Duplicate Account" {
                    New-HTMLSection -HeaderText 'Exchange - Duplicate Account' {
                        New-HTMLTable -DataTable $Script:Reporting['MailboxProblems']['Data']['Both']['DuplicateAccount'] -Filtering
                    }
                }
                New-HTMLTab -Name 'No Database' {
                    New-HTMLSection -HeaderText 'Exchange On-Premises - No Database' {
                        New-HTMLTable -DataTable $Script:Reporting['MailboxProblems']['Data']['Local']['NoDatabase'] -Filtering
                    }
                }
                New-HTMLTab -Name 'Inconsistent Data' {
                    New-HTMLSection -HeaderText 'Exchange Online - Broken DisplayName' {
                        New-HTMLTable -DataTable $Script:Reporting['MailboxProblems']['Data']['Online']['BrokenDisplayName'] -Filtering
                    }
                    New-HTMLSection -HeaderText 'Exchange On-Premises - Broken DisplayName' {
                        New-HTMLTable -DataTable $Script:Reporting['MailboxProblems']['Data']['Local']['BrokenDisplayName'] -Filtering
                    }
                    New-HTMLSection -HeaderText 'Exchange On-Premises - Missing UserPrincipalName' {
                        New-HTMLTable -DataTable $Script:Reporting['MailboxProblems']['Data']['Local']['MissingUserPrincipalName'] -Filtering
                    }
                }
                New-HTMLTab -Name 'Contact Problems' {
                    New-HTMLSection -HeaderText 'Exchange Online - Contact Missing PrimarySmtp' {
                        New-HTMLTable -DataTable $Script:Reporting['MailboxProblems']['Data']['Online']['ContactMissingPrimarySmtp'] -Filtering
                    }
                    New-HTMLSection -HeaderText 'Exchange On-Premises - Contact Missing PrimarySmtp' {
                        New-HTMLTable -DataTable $Script:Reporting['MailboxProblems']['Data']['Local']['ContactMissingPrimarySmtp'] -Filtering
                    }
                    New-HTMLSection -HeaderText 'Exchange Online - Contact Missing ExternalEmail' {
                        New-HTMLTable -DataTable $Script:Reporting['MailboxProblems']['Data']['Online']['ContactMissingExternalEmail'] -Filtering
                    }
                    New-HTMLSection -HeaderText 'Exchange On-Premises - Contact Missing ExternalEmail' {
                        New-HTMLTable -DataTable $Script:Reporting['MailboxProblems']['Data']['Local']['ContactMissingExternalEmail'] -Filtering
                    }
                }
            }
        }
    }
}