# This script grants full access permissions to multiple users on all room mailboxes in Exchange Online.
# Ensure you are connected to Exchange Online PowerShell before running this script.

# Define the users to grant permissions to (replace with actual email addresses).
$usersToGrant = @("jordan.hunt@enstargroup.com", "natalie.osbourne@enstargroup.com")

# Get all room mailboxes in Exchange Online.
$roomMailboxes = Get-Mailbox -RecipientTypeDetails RoomMailbox

# Display all room mailboxes by their display name.
Write-Host "List of all room mailboxes by display name:"
$roomMailboxes | ForEach-Object {
    Write-Host $_.DisplayName
}

# Loop through each room mailbox to assign permissions.
foreach ($room in $roomMailboxes) {
    foreach ($user in $usersToGrant) {
        # Grant full access permissions to the user for the room mailbox.
        Add-MailboxPermission -Identity $room.Alias -User $user -AccessRights FullAccess -InheritanceType All -AutoMapping $false
        
        # Output a message indicating the permission was successfully granted.
        Write-Host "Granted full access to $($room.DisplayName) ($($room.Alias)) for $user"
    }
}

# Output a final message indicating the script has completed successfully.
Write-Host "Permissions granted to all room mailboxes for specified users."



