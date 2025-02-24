$DisplayName = "John Doe"  # Replace with the actual Display Name
$GroupName = "Group Name" # Replace with the groupname

# Retrieve the user's AD account using Display Name
$User = Get-ADUser -Filter {DisplayName -eq $DisplayName} -Properties SamAccountName

if ($User) {
    # Retrieve the members of the group and check if the user is already a member
    $GroupMembers = Get-ADGroupMember -Identity $GroupName | Select-Object -ExpandProperty SamAccountName

    if ($GroupMembers -contains $User.SamAccountName) {
        Write-Output "User '$DisplayName' is already a member of '$GroupName'."
    } else {
        # Add the user to the AD group
        Add-ADGroupMember -Identity $GroupName -Members $User.SamAccountName
        Write-Output "User '$DisplayName' has been added to group '$GroupName'."
    }
} else {
    Write-Output "User with Display Name '$DisplayName' not found in Active Directory."
}
