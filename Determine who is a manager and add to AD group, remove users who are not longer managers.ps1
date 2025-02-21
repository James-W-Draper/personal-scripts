# Determine who is a manager and add them to an AD group, removing those who are no longer managers
# This PowerShell script dynamically retrieves users from specified OUs, identifies managers, 
# updates a designated AD group, and removes outdated members.

# Import the Active Directory module
Import-Module ActiveDirectory

# Parameters for script customization
param (
    [string]$Group = "CN=ManagersGroup,OU=Groups,DC=example,DC=com",
    [array]$OUs = @(
        "OU=Employees,DC=example,DC=com",
        "OU=Contractors,DC=example,DC=com"
    ),
    [switch]$Verbose
)

# Set verbose preference
if ($Verbose) {
    $VerbosePreference = "Continue"
} else {
    $VerbosePreference = "SilentlyContinue"
}

# Validate Group existence
try {
    Get-ADGroup -Identity $Group -ErrorAction Stop | Out-Null
} catch {
    Write-Warning "The specified AD group '$Group' does not exist. Please verify the group name."
    exit
}

# Initialize a list to store users
$Users = [System.Collections.ArrayList]@()

# Retrieve users from specified OUs
foreach ($OU in $OUs) {
    Write-Verbose "Fetching users from $OU..."
    try {
        $OUUsers = Get-ADUser -Filter * -SearchBase $OU -Properties Manager
        $Users.AddRange($OUUsers)
        Write-Verbose "Found $($OUUsers.Count) users in $OU."
    } catch {
        Write-Warning "Failed to retrieve users from $OU. Error: $_"
    }
}

# Exit if no users were found
if ($Users.Count -eq 0) {
    Write-Warning "No users were found in the specified OUs. Exiting script."
    exit
}

# Collect unique managers
$Managers = New-Object System.Collections.Generic.HashSet[Object]

foreach ($User in $Users) {
    if ($User.Manager) {
        try {
            $Manager = Get-ADUser -Identity $User.Manager
            $Managers.Add($Manager) | Out-Null
            Write-Verbose "Manager $($Manager.SamAccountName) identified for $($User.SamAccountName)."
        } catch {
            Write-Warning "Unable to retrieve manager for $($User.SamAccountName). Error: $_"
        }
    } else {
        Write-Verbose "$($User.SamAccountName) has no assigned manager."
    }
}

# Exit if no managers were found
if ($Managers.Count -eq 0) {
    Write-Warning "No managers identified. Exiting script."
    exit
}

# Convert HashSet to array for processing
$Managers = @($Managers)

# Retrieve current members of the AD group
$GroupMembers = Get-ADGroupMember -Identity $Group | Select-Object -ExpandProperty SamAccountName

# Determine which users should be removed
$MembersToRemove = $GroupMembers | Where-Object { $_ -notin $Managers.SamAccountName }

if ($MembersToRemove.Count -gt 0) {
    Write-Verbose "Removing non-managers from $Group..."
    try {
        Remove-ADGroupMember -Identity $Group -Members $MembersToRemove -Confirm:$false
        Write-Verbose "Successfully removed outdated members."
    } catch {
        Write-Warning "Error removing users from $Group. Error: $_"
    }
}

# Determine which managers should be added
$ManagersToAdd = $Managers | Where-Object { $_.SamAccountName -notin $GroupMembers }

if ($ManagersToAdd.Count -gt 0) {
    Write-Verbose "Adding new managers to $Group..."
    try {
        Add-ADGroupMember -Identity $Group -Members $ManagersToAdd.SamAccountName
        Write-Verbose "Successfully added new managers."
    } catch {
        Write-Warning "Error adding managers to $Group. Error: $_"
    }
}

Write-Verbose "Script execution completed."
