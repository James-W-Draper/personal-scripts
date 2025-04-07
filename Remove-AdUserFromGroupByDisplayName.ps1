<#
.SYNOPSIS
Removes a user from an Active Directory group using their Display Name.

.DESCRIPTION
This script looks up a user in Active Directory by their Display Name and removes them
from a specified AD group if they are a current member.

.PARAMETER DisplayName
The full display name of the user as it appears in Active Directory.

.PARAMETER GroupName
The name of the AD group from which the user should be removed.

.NOTES
- Requires the ActiveDirectory module
- Must be run with sufficient privileges to modify group membership
#>

# === CONFIGURATION ===
$DisplayName = "John Doe"      # <-- Replace with the actual Display Name
$GroupName   = "Group Name"    # <-- Replace with the target group name

# === LOOKUP USER ===
$User = Get-ADUser -Filter { DisplayName -eq $DisplayName } -Properties SamAccountName

if ($User) {
    $UserSam = $User.SamAccountName
    Write-Output "Found user: $DisplayName (sAMAccountName: $UserSam)"

    # === CHECK GROUP MEMBERSHIP ===
    $GroupMembers = Get-ADGroupMember -Identity $GroupName | Select-Object -ExpandProperty SamAccountName

    if ($GroupMembers -contains $UserSam) {
        Remove-ADGroupMember -Identity $GroupName -Members $UserSam -Confirm:$false
        Write-Output "User '$DisplayName' has been removed from group '$GroupName'."
    } else {
        Write-Output "User '$DisplayName' is not currently a member of '$GroupName'. No action taken."
    }
} else {
    Write-Warning "User with Display Name '$DisplayName' not found in Active Directory."
}
