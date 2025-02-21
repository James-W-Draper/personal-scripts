# 🔹 Essential PowerShell Commands for Exchange Online (Hybrid Environment)

A collection of **10 useful PowerShell commands** for managing Exchange Online in a **hybrid setup**. These commands help with mailbox management, delegation, forwarding, archiving, migrations, and more.

---

## 📌 1. Get a list of mailboxes and their primary email addresses
```powershell
Get-Mailbox -ResultSize Unlimited | Select DisplayName,PrimarySMTPAddress


### Get all inactive mailboxes, and the aliases. Export to csv.
`Get-Mailbox -InactiveMailboxOnly -ResultSize Unlimited | Select DisplayName, PrimarySMTPAddress, DistinguishedName, ExchangeGuid, WhenSoftDeleted, @{Name="Aliases";Expression={$_.EmailAddresses -match "^smtp:" -replace "smtp:" -join "; "}} | Export-Csv -Path "C:\Temp\InactiveMailboxes.csv" -NoTypeInformation -Encoding UTF8`

### Get a list of mailboxes and their primary email addresses
`Get-Mailbox -ResultSize Unlimited | Select DisplayName,PrimarySMTPAddress`

### Find all shared mailboxes
`Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Select DisplayName,PrimarySMTPAddress`

### Check mailbox delegation (Full Access & Send As)
`Get-Mailbox -ResultSize Unlimited | 
Select DisplayName,PrimarySMTPAddress, 
       @{Name="FullAccess";Expression={(Get-MailboxPermission $_.Identity | Where-Object {($_.AccessRights -match "FullAccess") -and ($_.User -notmatch "NT AUTHORITY\\SELF")} | Select-Object User -ExpandProperty User) -join ", "}}, 
       @{Name="SendAs";Expression={(Get-RecipientPermission $_.Identity | Where-Object {($_.AccessRights -match "SendAs")} | Select-Object Trustee -ExpandProperty Trustee) -join ", "}}
`

### Formatting guidance
[basic-writing-and-formatting-syntax](https://github.com/github/docs/blob/main/content/get-started/writing-on-github/getting-started-with-writing-and-formatting-on-github/basic-writing-and-formatting-syntax.md)
