# Connect to Exchange Online (Run this if not already connected)
Connect-ExchangeOnline

# Define your username
$YourUser = "draperj@enstargroup.com"

# Get all mailboxes and process permissions removal
Get-Mailbox -ResultSize Unlimited | ForEach-Object {
    $mailbox = $_.PrimarySmtpAddress
    $hasPermissions = $false

    # Remove Full Access Permissions
    $fullAccessPerm = Get-MailboxPermission -Identity $mailbox | Where-Object { $_.User -like $YourUser -and $_.AccessRights -ne "Owner" }
    if ($fullAccessPerm) {
        Write-Host "Removing Full Access from $mailbox"
        Remove-MailboxPermission -Identity $mailbox -User $YourUser -AccessRights FullAccess -Confirm:$false
        $hasPermissions = $true
    }

    # Remove SendAs Permissions
    $sendAsPerm = Get-RecipientPermission -Identity $mailbox | Where-Object { $_.Trustee -like $YourUser }
    if ($sendAsPerm) {
        Write-Host "Removing SendAs from $mailbox"
        Remove-RecipientPermission -Identity $mailbox -Trustee $YourUser -AccessRights SendAs -Confirm:$false
        $hasPermissions = $true
    }

    # Remove SendOnBehalf Permissions
    $sendOnBehalfPerm = (Get-Mailbox -Identity $mailbox).GrantSendOnBehalfTo
    if ($sendOnBehalfPerm -contains $YourUser) {
        Write-Host "Removing SendOnBehalf from $mailbox"
        Set-Mailbox -Identity $mailbox -GrantSendOnBehalfTo @{remove=$YourUser}
        $hasPermissions = $true
    }

    if (-not $hasPermissions) {
        Write-Host "No permissions found to remove on $mailbox"
    }
}

Write-Host "Permission removal process completed!"
