<#
.SYNOPSIS
Removes a user from an Active Directory group using their Display Name.

.DESCRIPTION
This script searches Active Directory for a user by Display Name and removes them from a specified AD group if they are a member. 
It includes safety checks and does not prompt for confirmation unless the -WhatIf flag is used.

.PARAMETER DisplayName
The full display name of the user as it appears in Active Directory.

.PARAMETER GroupName
The name of the Active Directory group from which the user should be removed.

.EXAMPLE
.\Remove-UserFromADGroup.ps1 -DisplayName "John Doe" -GroupName "HR Access Group"

.NOTES
Requires:
- ActiveDirectory module
- Appropriate permissions to read user/group data and modify group membership
#>

param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter the display name of the user to remove")]
    [string]$DisplayName,

    [Parameter(Mandatory = $true, HelpMessage = "Enter the name of the group to remove the user from")]
    [string]$GroupName
)

# Ensure the Active Directory module is loaded
if (-not (Get-Module -Name ActiveDirectory)) {
    Import-Module ActiveDirectory -ErrorAction Stop
}

# Lookup the user in AD
$User = Get-ADUser -Filter { DisplayName -eq $DisplayName } -Properties SamAccountName

if ($User) {
    $UserSam = $User.SamAccountName
    Write-Output "Found user: $DisplayName (sAMAccountName: $UserSam)"

    # Check if the user is in the group
    $GroupMembers = Get-ADGroupMember -Identity $GroupName -ErrorAction Stop | Select-Object -ExpandProperty SamAccountName

    if ($GroupMembers -contains $UserSam) {
        Remove-ADGroupMember -Identity $GroupName -Members $UserSam -Confirm:$false
        Write-Output "User '$DisplayName' has been removed from group '$GroupName'."
    } else {
        Write-Output "User '$DisplayName' is not currently a member of '$GroupName'. No action taken."
    }
} else {
    Write-Warning "User with Display Name '$DisplayName' not found in Active Directory."
}
