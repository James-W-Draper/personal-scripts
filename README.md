# personal-scripts

## Get all inactive mailboxes, and the aliases. Export to csv.

`Get-Mailbox -InactiveMailboxOnly -ResultSize Unlimited | Select DisplayName, PrimarySMTPAddress, DistinguishedName, ExchangeGuid, WhenSoftDeleted, @{Name="Aliases";Expression={$_.EmailAddresses -match "^smtp:" -replace "smtp:" -join "; "}} | Export-Csv -Path "C:\Temp\InactiveMailboxes.csv" -NoTypeInformation -Encoding UTF8`
