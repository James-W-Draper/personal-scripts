# Connect to Exchange Online
# Connect-exchangeonline

# Define the Organizational Unit (OU) paths for the search
# You can add more OUs as needed by defining additional variables
$OU1 = "OU=Inigo,OU=Group Users,DC=cwglobal,DC=local"
$OU2 = "OU=Inigo2,OU=Group Users,DC=cwglobal,DC=local"

# Retrieve all users from the specified OUs
# The Get-ADUser cmdlet fetches user accounts from Active Directory
$usersOU1 = Get-ADUser -SearchBase $OU1 -Filter * -Property UserPrincipalName
$usersOU2 = Get-ADUser -SearchBase $OU2 -Filter * -Property UserPrincipalName

# Combine the user accounts from both OUs into a single collection
$users = $usersOU1 + $usersOU2

# Identify users whose UserPrincipalName (UPN) does not end with "@enstargroup.com"
# This is for informational purposes, and can be used to list users that do not match the domain
# $nonMatchingUsers = $users | Where-Object { $_.UserPrincipalName -notlike "*@enstargroup.com" }

# (Optional) Display UPNs that do not match the domain
# This section is commented out, but can be enabled to log or display non-matching UPNs
#$nonMatchingUsers | ForEach-Object {
#    Write-Host "UserPrincipalName (non-matching): $($_.UserPrincipalName)"
#}

# Filter users whose UPN ends with "@enstargroup.com"
# Only these users will have their mailboxes converted to shared mailboxes
$filteredUsers = $users | Where-Object { $_.UserPrincipalName -like "*@enstargroup.com" }

# Display UPNs of the filtered users for verification
$filteredUsers | ForEach-Object {
    Write-Host "UserPrincipalName: $($_.UserPrincipalName)"
}

# Assuming you are already connected to Exchange Online, loop through the filtered users
# and convert their mailboxes to shared mailboxes
$filteredUsers | ForEach-Object {
    $upn = $_.UserPrincipalName
    Try {
        # Use Set-Mailbox to convert the user's mailbox to a shared mailbox
        Set-Mailbox -Identity $upn -Type Shared
        Write-Host "Successfully set mailbox type to shared for: $upn"
    } Catch {
        # If there's an error, display the error message in red
        Write-Host "Error setting mailbox for: $upn - $($_.Exception.Message)" -ForegroundColor Red
    }
}
