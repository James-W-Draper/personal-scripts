<#
.SYNOPSIS
Exports Microsoft Teams shared channels and their external (guest) members to an Excel file.

.DESCRIPTION
Connects to Microsoft Teams, finds all shared channels across all teams, and retrieves guest members' display names and email addresses.
Outputs the results to an Excel file using the ImportExcel PowerShell module.

.PARAMETER OutputPath
Optional. Full path to the output Excel file. Defaults to "C:\Temp\SharedChannels.xlsx".

.EXAMPLE
.\Export-M365SharedChannelsGuestMembersToExcel.ps1 -OutputPath "D:\Reports\SharedGuests.xlsx"

.NOTES
- Requires the MicrosoftTeams and ImportExcel modules.
- Run with sufficient Teams admin permissions.
#>

[CmdletBinding()]
param (
    [string]$OutputPath = "C:\Temp\SharedChannels.xlsx"
)

# Ensure output directory exists
$OutputDirectory = Split-Path -Path $OutputPath -Parent
if (-not (Test-Path $OutputDirectory)) {
    New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
}

# Install MicrosoftTeams if not already installed
if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) {
    Install-Module -Name MicrosoftTeams -Scope CurrentUser -Force
}
Import-Module MicrosoftTeams

# Install ImportExcel if not already installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel

# Connect to Microsoft Teams
try {
    Connect-MicrosoftTeams -ErrorAction Stop
} catch {
    Write-Error "❌ Failed to connect to Microsoft Teams: $_"
    return
}

# Retrieve all Teams
$teams = Get-Team

# Prepare result array
$sharedChannelGuests = @()

foreach ($team in $teams) {
    # Get shared channels
    $sharedChannels = Get-TeamChannel -GroupId $team.GroupId | Where-Object { $_.MembershipType -eq "Shared" }

    foreach ($channel in $sharedChannels) {
        # Get members of the shared channel
        $members = Get-TeamChannelUser -GroupId $team.GroupId -DisplayName $channel.DisplayName

        # Filter and format guest users
        $guestUsers = $members | Where-Object { $_.UserType -eq "Guest" }

        foreach ($guest in $guestUsers) {
            $sharedChannelGuests += [PSCustomObject]@{
                TeamName          = $team.DisplayName
                SharedChannelName = $channel.DisplayName
                GuestDisplayName  = $guest.Name
                GuestEmail        = $guest.User
            }
        }
    }
}

# Export to Excel
$sharedChannelGuests | Export-Excel -Path $OutputPath -WorksheetName "GuestMembers" -BoldTopRow -AutoSize -AutoFilter -FreezeTopRow
Write-Host "✅ Excel report saved to: $OutputPath"
