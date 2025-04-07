<#
.SYNOPSIS
Adds a user to an Active Directory group using their Display Name.

.DESCRIPTION
This script searches Active Directory for a user by their Display Name and adds them
to a specified AD group if they're not already a member.

.PARAMETER DisplayName
The full display name of the user as it appears in Active Directory.

.PARAMETER GroupName
The name of the AD group to which the user should be added.

.NOTES
- Requires the ActiveDirectory module
- Must be run with appropriate privileges to modify group membership
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
        Write-Output "User '$DisplayName' is already a member of '$GroupName'."
    } else {
        Add-ADGroupMember -Identity $GroupName -Members $UserSam
        Write-Output "User '$DisplayName' has been added to group '$GroupName'."
    }
} else {
    Write-Warning "User with Display Name '$DisplayName' not found in Active Directory."
}
