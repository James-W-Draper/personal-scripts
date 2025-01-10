# Import the AzureAD module
# This module provides cmdlets to interact with Azure Active Directory, allowing us to retrieve and manage directory roles and members.
Import-Module AzureAD

# Connect to Azure AD
# Prompts the administrator for credentials to authenticate and establish a session with Azure AD.
Connect-AzureAD

# Get all roles in Azure AD, treating all roles as potentially privileged
# This retrieves a list of all available directory roles within Azure AD. Each role represents a specific set of permissions or privileges.
$allRoles = Get-AzureADDirectoryRole

# Initialize an empty array to store user-role information
# This array will hold the details of each user with their associated role(s).
$privilegedAccounts = @()

# Loop through each role found in Azure AD
# For each role, we retrieve the list of users assigned to it and log their details.
foreach ($role in $allRoles) {

    # Get all members assigned to the current role
    # For each role, this command fetches all users (members) assigned to that role.
    $roleMembers = Get-AzureADDirectoryRoleMember -ObjectId $role.ObjectId

    # Loop through each member assigned to the role
    # This allows us to create a record for each user with the role they are assigned to.
    foreach ($member in $roleMembers) {
        
        # Add each user-role association to the array
        # We create a custom object for each user, storing their DisplayName, UserPrincipalName, and assigned Role.
        $privilegedAccounts += [pscustomobject]@{
            DisplayName       = $member.DisplayName       # User's display name, helpful for easy identification
            UserPrincipalName = $member.UserPrincipalName # The user's principal name, often their login email address
            Role              = $role.DisplayName         # The name of the role assigned to the user
        }
    }
}

# Define the file path where the report will be saved
# This specifies the location and file name for the CSV output, which will contain all privileged accounts and their roles.
$outputFile = "C:\scripts\AllPrivilegedAccountsReport.csv"

# Export the array of privileged accounts to a CSV file
# -Path specifies the file location
# -NoTypeInformation prevents extra type info from appearing in the file, making it cleaner
# -Encoding UTF8 ensures compatibility with most text and spreadsheet applications
$privilegedAccounts | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8

# Output a message to the console to confirm that the report was generated
# This message provides feedback to the user, indicating the script has completed and where to find the output file.
Write-Output "All Privileged Accounts Report generated: $outputFile"
