<#
.SYNOPSIS
    Grants Full Access permissions to specified users on all room mailboxes in Exchange Online.

.DESCRIPTION
    This script retrieves all room mailboxes in Exchange Online and grants full mailbox access
    to each specified user. This is useful for administrative or booking delegate purposes.
    The script assumes an active Exchange Online PowerShell session.

.EXAMPLE
    .\Grant-RoomAccessPermissions.ps1

.NOTES
    Ensure you are connected to Exchange Online PowerShell before running.
#>

# Define the users to grant permissions to (replace with actual email addresses).
$usersToGrant = @(
    "jordan.hunt@enstargroup.com",
    "natalie.osbourne@enstargroup.com"
)

# Get all room mailboxes in Exchange Online.
$roomMailboxes = Get-Mailbox -RecipientTypeDetails RoomMailbox

# Display all room mailboxes by their display name.
Write-Host "List of all room mailboxes by display name:" -ForegroundColor Cyan
$roomMailboxes | ForEach-Object {
    Write-Host $_.DisplayName -ForegroundColor Yellow
}

# Loop through each room mailbox to assign permissions.
foreach ($room in $roomMailboxes) {
    foreach ($user in $usersToGrant) {
        try {
            Add-MailboxPermission -Identity $room.Alias -User $user -AccessRights FullAccess -InheritanceType All -AutoMapping:$false -ErrorAction Stop
            Write-Host "Granted full access to '$($room.DisplayName)' ($($room.Alias)) for $user" -ForegroundColor Green
        }
        catch {
            Write-Warning "Failed to grant access to '$($room.DisplayName)' for $user. Error: $_"
        }
    }
}

Write-Host "\n✅ Permissions granted to all room mailboxes for specified users." -ForegroundColor Cyan
