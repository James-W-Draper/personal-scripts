<#
.SYNOPSIS
    Updates an AD group to contain only users who are currently listed as managers in specified OUs.

.DESCRIPTION
    This PowerShell script retrieves all users from specified Organizational Units (OUs) in Active Directory,
    identifies their managers, adds the managers to a specific AD group, and removes any users from the group
    who are no longer managers. Supports verbose output for detailed execution logs.

.EXAMPLE
    .\Update-ADGroupWithManagers.ps1

.EXAMPLE
    .\Update-ADGroupWithManagers.ps1 -Verbose
#>

Import-Module ActiveDirectory

param (
    [switch]$Verbose
)

if ($Verbose) {
    $VerbosePreference = "Continue"
} else {
    $VerbosePreference = "SilentlyContinue"
}

# Define target AD group and OUs
$Group = "CN=ManagersGroup,OU=Groups,DC=yourdomain,DC=com"
$OUs = @(
    "OU=Sales,DC=yourdomain,DC=com",
    "OU=Marketing,DC=yourdomain,DC=com"
)

$Users = [System.Collections.ArrayList]@()

foreach ($OU in $OUs) {
    Write-Verbose "Collecting users from $OU..."
    $OUUsers = Get-ADUser -Filter * -SearchBase $OU -Properties Manager
    $Users.AddRange($OUUsers)
    Write-Verbose "Found $($OUUsers.Count) users in $OU."
}

if ($Users.Count -eq 0) {
    Write-Warning "No users were found in the specified OUs."
    return
}

$Managers = New-Object System.Collections.Generic.HashSet[Object]

foreach ($User in $Users) {
    if ($null -ne $User.Manager) {
        try {
            $Manager = Get-ADUser -Identity $User.Manager
            $Managers.Add($Manager) | Out-Null
            Write-Verbose "Manager $($Manager.SamAccountName) found for user $($User.SamAccountName)."
        } catch {
            Write-Warning "Failed to retrieve manager for user $($User.SamAccountName). Error: $_"
        }
    } else {
        Write-Verbose "No manager found for user $($User.SamAccountName)."
    }
}

if ($Managers.Count -eq 0) {
    Write-Warning "No managers were found, exiting script."
    return
}

$Managers = $Managers.ToArray()
$GroupMembers = Get-ADGroupMember -Identity $Group | Select-Object -ExpandProperty SamAccountName

$MembersToRemove = $GroupMembers | Where-Object { $Managers.SamAccountName -notcontains $_ }
if ($MembersToRemove.Count -gt 0) {
    Write-Verbose "Removing users no longer valid managers from group $Group..."
    try {
        Remove-ADGroupMember -Identity $Group -Members $MembersToRemove -Confirm:$false
        Write-Verbose "Users removed successfully."
    } catch {
        Write-Warning "Failed to remove users from group $Group. Error: $_"
    }
}

$ManagersToAdd = $Managers | Where-Object { $GroupMembers -notcontains $_.SamAccountName }
if ($ManagersToAdd.Count -gt 0) {
    Write-Verbose "Adding new managers to group $Group..."
    try {
        Add-ADGroupMember -Identity $Group -Members $ManagersToAdd.SamAccountName
        Write-Verbose "Managers added successfully."
    } catch {
        Write-Warning "Failed to add managers to group $Group. Error: $_"
    }
}

Write-Verbose "Script execution completed."
