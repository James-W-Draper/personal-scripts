# Determine who is a manager and add to AD group, remove users who are no longer managers
# This PowerShell script retrieves users from specific Organizational Units (OUs) in Active Directory,
# extracts their managers, adds valid managers to a specific AD group, and removes any members from the 
# group who are no longer managers.

# Import the Active Directory module to use AD-related cmdlets
Import-Module ActiveDirectory

# Parameters for toggling verbose output
param (
    [switch]$Verbose
)

# Set verbose preference based on input parameter
if ($Verbose) {
    $VerbosePreference = "Continue"
} else {
    $VerbosePreference = "SilentlyContinue"
}

# Define the AD group to which managers will be added
$Group = "CN=ManagersGroup,OU=Groups,DC=yourdomain,DC=com"

# Define the Organizational Units (OUs) from which users will be collected
$OUs = @(
    "OU=Sales,DC=yourdomain,DC=com",     # Sales department OU
    "OU=Marketing,DC=yourdomain,DC=com"  # Marketing department OU
)

# Initialize an array list to store users collected from the specified OUs for better performance
$Users = [System.Collections.ArrayList]@()

# Loop through each OU and collect all users
foreach ($OU in $OUs) {
    Write-Verbose "Collecting users from $OU..."
    
    # Get all users from the current OU, including their 'Manager' property
    $OUUsers = Get-ADUser -Filter * -SearchBase $OU -Properties Manager
    
    # Add the retrieved users to the $Users array list
    $Users.AddRange($OUUsers)

    # Output the number of users found in each OU
    Write-Verbose "Found $($OUUsers.Count) users in $OU."
}

# If no users were found in the specified OUs, output a warning and terminate the script
if ($Users.Count -eq 0) {
    Write-Warning "No users were found in the specified OUs."
    return
}

# Initialize a hash set to store unique managers for better performance
$Managers = New-Object System.Collections.Generic.HashSet[Object]

# Loop through each user to collect their manager
foreach ($User in $Users) {
    if ($null -ne $User.Manager) {    # $null is now on the left side of the comparison
        Write-Verbose "Getting manager for user $($User.SamAccountName)..."

        # Retrieve the manager's details using their distinguished name (DN) stored in the 'Manager' property
        try {
            $Manager = Get-ADUser -Identity $User.Manager
            $Managers.Add($Manager)
            Write-Verbose "Manager $($Manager.SamAccountName) found for user $($User.SamAccountName)."
        } catch {
            # Output a warning if the manager cannot be retrieved
            Write-Warning "Failed to retrieve manager for user $($User.SamAccountName). Error: $_"
        }
    } else {
        # Output a message if no manager is assigned to the user
        Write-Verbose "No manager found for user $($User.SamAccountName)."
    }
}

# If no managers were found, output a warning and terminate the script
if ($Managers.Count -eq 0) {
    Write-Warning "No managers were found, exiting script."
    return
}

# Convert HashSet to array for processing
$Managers = $Managers.ToArray()

# Retrieve the current members of the AD group
$GroupMembers = Get-ADGroupMember -Identity $Group | Select-Object SamAccountName

# Remove users in bulk if they're no longer managers
$MembersToRemove = $GroupMembers | Where-Object { $Managers.SamAccountName -notcontains $_.SamAccountName }
if ($MembersToRemove.Count -gt 0) {
    Write-Verbose "Removing users no longer valid managers from group $Group..."
    try {
        Remove-ADGroupMember -Identity $Group -Members $MembersToRemove.SamAccountName -Confirm:$false
        Write-Verbose "Users removed successfully."
    } catch {
        Write-Warning "Failed to remove users from group $Group. Error: $_"
    }
}

# Add managers in bulk who are not already members
$ManagersToAdd = $Managers | Where-Object { $GroupMembers.SamAccountName -notcontains $_.SamAccountName }
if ($ManagersToAdd.Count -gt 0) {
    Write-Verbose "Adding new managers to group $Group..."
    try {
        Add-ADGroupMember -Identity $Group -Members $ManagersToAdd.SamAccountName
        Write-Verbose "Managers added successfully."
    } catch {
        Write-Warning "Failed to add managers to group $Group. Error: $_"
    }
}

# Output a final message to indicate the script has completed successfully
Write-Verbose "Script execution completed."
