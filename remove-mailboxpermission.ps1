<#
.SYNOPSIS
Removes all delegated mailbox permissions (Full Access, Send As, Send on Behalf) for a specified user across all mailboxes in Exchange Online.

.DESCRIPTION
This script connects to Exchange Online and iterates through all mailboxes to remove the specified user's:
- Full Access permissions
- Send As permissions
- Send on Behalf permissions

.PARAMETER UserPrincipalName
The UPN of the user whose permissions should be removed (e.g., user@example.com).

.EXAMPLE
.\Remove-MailboxDelegations.ps1 -UserPrincipalName "user@example.com"

.NOTES
- Requires ExchangeOnlineManagement module
- Must be connected to Exchange Online or use Connect-ExchangeOnline at runtime
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName
)

# Ensure Exchange Online session is connected
if (-not (Get-PSSession | Where-Object { $_.ComputerName -like "*outlook.office365.com*" })) {
    try {
        Connect-ExchangeOnline -ErrorAction Stop
    } catch {
        Write-Error "Failed to connect to Exchange Online. $_"
        exit
    }
}

Write-Host "Removing delegated permissions for: $UserPrincipalName" -ForegroundColor Cyan

Get-Mailbox -ResultSize Unlimited | ForEach-Object {
    $mailbox = $_
    $mailboxIdentity = $mailbox.Identity
    $mailboxAddress = $mailbox.PrimarySmtpAddress
    $hasChanges = $false

    # FULL ACCESS
    try {
        $fullAccess = Get-MailboxPermission -Identity $mailboxIdentity | Where-Object {
            $_.User.ToString().ToLower() -eq $UserPrincipalName.ToLower() -and $_.AccessRights -contains "FullAccess" -and -not $_.IsInherited
        }

        if ($fullAccess) {
            Write-Host "Removing Full Access from $mailboxAddress"
            Remove-MailboxPermission -Identity $mailboxIdentity -User $UserPrincipalName -AccessRights FullAccess -Confirm:$false -ErrorAction Stop
            $hasChanges = $true
        }
    } catch {
        Write-Warning "Failed to process Full Access for $mailboxAddress - $_"
    }

    # SEND AS
    try {
        $sendAs = Get-RecipientPermission -Identity $mailboxIdentity | Where-Object {
            $_.Trustee.ToString().ToLower() -eq $UserPrincipalName.ToLower() -and $_.AccessRights -contains "SendAs"
        }

        if ($sendAs) {
            Write-Host "Removing Send As from $mailboxAddress"
            Remove-RecipientPermission -Identity $mailboxIdentity -Trustee $UserPrincipalName -AccessRights SendAs -Confirm:$false -ErrorAction Stop
            $hasChanges = $true
        }
    } catch {
        Write-Warning "Failed to process Send As for $mailboxAddress - $_"
    }

    # SEND ON BEHALF
    try {
        $sobList = ($mailbox.GrantSendOnBehalfTo | ForEach-Object { $_.ToString().ToLower() })
        if ($sobList -contains $UserPrincipalName.ToLower()) {
            Write-Host "Removing Send on Behalf from $mailboxAddress"
            Set-Mailbox -Identity $mailboxIdentity -GrantSendOnBehalfTo @{remove=$UserPrincipalName} -ErrorAction Stop
            $hasChanges = $true
        }
    } catch {
        Write-Warning "Failed to process Send on Behalf for $mailboxAddress - $_"
    }

    if (-not $hasChanges) {
        Write-Host "No delegated permissions found for $UserPrincipalName on $mailboxAddress"
    }
}

Write-Host "`nPermission removal process completed!" -ForegroundColor Green
