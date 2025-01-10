# Define the group name
$groupName = "LIC - Copilot"

# Get all members of the group
$groupMembers = Get-ADGroupMember -Identity $groupName

# Filter to only include user accounts and count them
$userCount = $groupMembers | Where-Object { $_.objectClass -eq 'user' } | Measure-Object | Select-Object -ExpandProperty Count

# Output the result
Write-Host "Number of users in the group '$groupName': $userCount"
