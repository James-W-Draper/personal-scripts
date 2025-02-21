# personal-scripts
## Get all inactive mailboxes, and the aliases. Export to csv.
`Get-Mailbox -InactiveMailboxOnly -ResultSize Unlimited | Select DisplayName, PrimarySMTPAddress, DistinguishedName, ExchangeGuid, WhenSoftDeleted, @{Name="Aliases";Expression={$_.EmailAddresses -match "^smtp:" -replace "smtp:" -join "; "}} | Export-Csv -Path "C:\Temp\InactiveMailboxes.csv" -NoTypeInformation -Encoding UTF8`

### Formatting guidance

[basic-writing-and-formatting-syntax](https://github.com/github/docs/blob/main/content/get-started/writing-on-github/getting-started-with-writing-and-formatting-on-github/basic-writing-and-formatting-syntax.md)
