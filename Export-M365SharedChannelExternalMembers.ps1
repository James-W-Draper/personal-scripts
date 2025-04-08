<#
.SYNOPSIS
Exports a report of all Microsoft Teams shared channels and their external (guest) members.

.DESCRIPTION
This script connects to Microsoft Teams, lists all shared channels across all Teams, and retrieves the email addresses of external (guest) users in those shared channels.
It outputs the results to a CSV file, where each row contains the team name, shared channel name, and a comma-separated list of guest email addresses.

.PARAMETER OutputPath
Optional. The full path where the CSV report will be saved. Defaults to "C:\Temp\SharedChannels.csv".

.EXAMPLE
.\Export-M365SharedChannelExternalMembers.ps1 -OutputPath "D:\Reports\ExternalGuests.csv"

.NOTES
Requires the Microsoft Teams PowerShell module.
Must be run by a Teams Administrator or user with appropriate permissions.
#>

[CmdletBinding()]
param (
    [string]$OutputPath = "C:\Temp\SharedChannels.csv"
)

# Ensure MicrosoftTeams module is available
if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) {
    try {
        Install-Module -Name MicrosoftTeams -Force -Scope CurrentUser -ErrorAction Stop
    } catch {
        Write-Error "Failed to install MicrosoftTeams module: $_"
        return
    }
}

# Connect to Microsoft Teams
try {
    Connect-MicrosoftTeams -ErrorAction Stop
} catch {
    Write-Error "Failed to connect to Microsoft Teams: $_"
    return
}

# Retrieve all Teams
$teams = Get-Team

# Prepare results
$sharedChannelResults = @()

foreach ($team in $teams) {
    # Get shared channels (membership type = Private Shared)
    $channels = Get-TeamChannel -GroupId $team.GroupId | Where-Object { $_.MembershipType -eq "Shared" }

    foreach ($channel in $channels) {
        # Get members of the shared channel
        $channelMembers = Get-TeamChannelUser -GroupId $team.GroupId -DisplayName $channel.DisplayName

        # Filter external users (Guests)
        $guestEmails = $channelMembers |
            Where-Object { $_.UserType -eq "Guest" } |
            Select-Object -ExpandProperty User

        # Build CSV-friendly output
        $sharedChannelResults += [PSCustomObject]@{
            TeamName          = $team.DisplayName
            SharedChannelName = $channel.DisplayName
            GuestEmails       = ($guestEmails -join ", ")
        }
    }
}

# Export to CSV
$sharedChannelResults | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Export complete. File saved to: $OutputPath"
