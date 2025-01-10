# Connect to Microsoft Teams, list all shared channels on all teams
# and list all the email address of all members of each shared channel who are external of the tenant
# Export to csv a table of all the teams, each shared channel and it's members
# The list of email addresses should be in one column, separated by a comma

# This is version 2 of the script, which only exports guests from the results

# Install the Microsoft Teams PowerShell module if it's not already installed
if (-not (Get-Module -Name MicrosoftTeams)) {
    Install-Module -Name MicrosoftTeams -Force
}


# Connect to Microsoft Teams
Connect-MicrosoftTeams

# Get all teams
$teams = Get-Team

# Create an empty array to store the results
$sharedChannels = @()

# Loop through each team
foreach ($team in $teams) {
    # Get all shared channels for the team
    $channels = Get-TeamChannel -GroupId $team.GroupId -MembershipType "Private"

    # Loop through each shared channel
    foreach ($channel in $channels) {
        # Get all members of the shared channel
        $members = Get-TeamChannelUser -GroupId $team.GroupId -DisplayName $channel.DisplayName

        # Create an empty array to store the external members
        $externalMembers = @()

        # Loop through each member
        foreach ($member in $members) {
            # Check if the member is external
            if ($member.UserType -eq "Guest") {
                # Add the member to the external members array
                $externalMembers += $member
            }
        }

        # Create a new object with the team name, shared channel name and external members
        $sharedChannel = New-Object -TypeName PSObject -Property @{
            TeamName = $team.DisplayName
            SharedChannelName = $channel.DisplayName
            ExternalMembers = $externalMembers
        }

        # Add the shared channel to the results array
        $sharedChannels += $sharedChannel
    }
}

# Export the results to csv
$sharedChannels | Export-Csv -Path "C:\Temp\SharedChannels.csv" -NoTypeInformation