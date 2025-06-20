# Essential PowerShell Commands for Exchange Online (Hybrid Environment)

A collection of **10 useful PowerShell commands** for managing Exchange Online in a **hybrid setup**. These commands help with mailbox management, delegation, forwarding, archiving, migrations, and more.

---

### Get a list of mailboxes and their primary email addresses
``` powershell
Get-Mailbox -ResultSize Unlimited | Select DisplayName,PrimarySMTPAddress
```

### Get all inactive mailboxes, and the aliases. Export to csv.
``` powershell
Get-Mailbox -InactiveMailboxOnly -ResultSize Unlimited | Select DisplayName, PrimarySMTPAddress, DistinguishedName, ExchangeGuid, WhenSoftDeleted, @{Name="Aliases";Expression={$_.EmailAddresses -match "^smtp:" -replace "smtp:" -join "; "}} | Export-Csv -Path "C:\Temp\InactiveMailboxes.csv" -NoTypeInformation -Encoding UTF8
```

### Get a list of mailboxes and their primary email addresses
``` powershell
Get-Mailbox -ResultSize Unlimited | Select DisplayName,PrimarySMTPAddress
```

### Find all shared mailboxes
``` powershell
Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Select DisplayName,PrimarySMTPAddress
```

### Check mailbox delegation (Full Access & Send As)
``` powershell
Get-Mailbox -ResultSize Unlimited | Select DisplayName,PrimarySMTPAddress, @{Name="FullAccess";Expression={(Get-MailboxPermission $_.Identity | Where-Object {($_.AccessRights -match "FullAccess") -and ($_.User -notmatch "NT AUTHORITY\\SELF")} | Select-Object User -ExpandProperty User) -join ", "}}, @{Name="SendAs";Expression={(Get-RecipientPermission $_.Identity | Where-Object {($_.AccessRights -match "SendAs")} | Select-Object Trustee -ExpandProperty Trustee) -join ", "}}
```

### List all mailboxes with forwarding enabled
``` powershell
Get-Mailbox -ResultSize Unlimited | Where-Object { $_.ForwardingSMTPAddress -or $_.ForwardingAddress } | Select DisplayName,PrimarySMTPAddress,ForwardingSMTPAddress,ForwardingAddress,DeliverToMailboxAndForward
```

### Get a list of all mail-enabled security groups
``` powershell
Get-Mailbox -Archive -ResultSize Unlimited | Select DisplayName,PrimarySMTPAddress,ArchiveStatus
```

### Find users with mailbox archive enabled
``` powershell
Get-Mailbox -Archive -ResultSize Unlimited | Select DisplayName,PrimarySMTPAddress,ArchiveStatus
```

### Check the last login time of a mailbox
``` powershell
Get-MailboxStatistics -Identity user@domain.com | Select DisplayName,LastLogonTime
```

### Get a list of transport rules (mail flow rules)
``` powershell
Get-TransportRule | Select Name,Mode,Priority,Comments
```

### Get the mailbox size of a user
``` powershell
Get-MailboxStatistics -Identity user@domain.com | Select DisplayName,TotalItemSize,ItemCount
```

### Find users with external email forwarding enabled
``` powershell
Get-Mailbox -ResultSize Unlimited | Where-Object { $_.ForwardingSMTPAddress -match "@" -and $_.ForwardingSMTPAddress -notmatch "yourdomain.com" } | Select DisplayName,PrimarySMTPAddress,ForwardingSMTPAddress
```


### Formatting guidance
[basic-writing-and-formatting-syntax](https://github.com/github/docs/blob/main/content/get-started/writing-on-github/getting-started-with-writing-and-formatting-on-github/basic-writing-and-formatting-syntax.md)
