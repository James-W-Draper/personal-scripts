<#
.SYNOPSIS
    Audits external (guest) user access across Microsoft Teams, SharePoint Online, and Microsoft 365 Groups.

.DESCRIPTION
    This script connects to Microsoft Graph and Microsoft Teams to identify guest users and where they have
    access:
        - Teams membership
        - Microsoft 365 Group membership
        - (Optional future extension) SharePoint site sharing links or direct access

    It can optionally filter guest users by inactivity (no sign-in for X days), and exports results to Excel
    with one sheet per resource type.

.PARAMETER InactiveDays
    Optional. Only include guest users whose last sign-in was more than X days ago.

.PARAMETER OutputPath
    Optional. Directory to save the Excel file. Defaults to "C:\Scripts".

.EXAMPLE
    .\Audit-ExternalUserAccess.ps1 -InactiveDays 90

.EXAMPLE
    .\Audit-ExternalUserAccess.ps1 -OutputPath "D:\Reports"

.NOTES
    Requires:
    - Microsoft.Graph module (Teams, Users, Groups)
    - MicrosoftTeams module
    - ImportExcel module
    - Microsoft Graph permissions: Directory.Read.All, Group.Read.All, Reports.Read.All
#>

[CmdletBinding()]
param (
    [int]$InactiveDays,
    [string]$OutputPath = "C:\Scripts"
)

# Validate modules
foreach ($module in @("Microsoft.Graph", "MicrosoftTeams", "ImportExcel")) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Install-Module $module -Scope CurrentUser -Force
    }
}

Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Groups
Import-Module Microsoft.Graph.Reports
Import-Module MicrosoftTeams
Import-Module ImportExcel

# Connect to Graph and Teams
Connect-MgGraph -Scopes "Directory.Read.All", "Group.Read.All", "User.Read.All", "Reports.Read.All"
Select-MgProfile -Name beta
Connect-MicrosoftTeams

$cutoff = if ($PSBoundParameters.ContainsKey('InactiveDays')) {
    (Get-Date).AddDays(-$InactiveDays)
} else {
    $null
}

$guests = Get-MgUser -Filter "userType eq 'Guest'" -All -Property Id, DisplayName, Mail, SignInActivity, UserPrincipalName

if ($cutoff) {
    $guests = $guests | Where-Object {
        $lastSignIn = $_.SignInActivity.LastSignInDateTime
        -not $lastSignIn -or ($lastSignIn -lt $cutoff)
    }
}

# Audit Teams access
$teamsData = @()
foreach ($user in $guests) {
    try {
        $teams = Get-TeamUser -User $user.UserPrincipalName -ErrorAction SilentlyContinue
        foreach ($team in $teams) {
            $teamsData += [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                DisplayName       = $user.DisplayName
                LastSignIn        = $user.SignInActivity.LastSignInDateTime
                TeamName          = $team.DisplayName
                Role              = $team.Role
            }
        }
    } catch {}
}

# Audit Microsoft 365 Group membership
$groupData = @()
$allGroups = Get-MgGroup -Filter "groupTypes/any(c:c eq 'Unified')" -All -Property Id, DisplayName

foreach ($group in $allGroups) {
    try {
        $members = Get-MgGroupMember -GroupId $group.Id -All | Where-Object { $_.AdditionalProperties.userType -eq 'Guest' }
        foreach ($member in $members) {
            $user = $guests | Where-Object { $_.Id -eq $member.Id }
            if ($user) {
                $groupData += [PSCustomObject]@{
                    GroupName         = $group.DisplayName
                    UserPrincipalName = $user.UserPrincipalName
                    DisplayName       = $user.DisplayName
                    LastSignIn        = $user.SignInActivity.LastSignInDateTime
                }
            }
        }
    } catch {}
}

# Prepare Excel output
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFile = Join-Path $OutputPath "ExternalUserAccess_$timestamp.xlsx"

if ($teamsData.Count -gt 0) {
    $teamsData | Export-Excel -Path $outputFile -WorksheetName "TeamsAccess" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
}

if ($groupData.Count -gt 0) {
    $groupData | Export-Excel -Path $outputFile -WorksheetName "M365GroupAccess" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
}

Write-Host "Report saved to: $outputFile" -ForegroundColor Green

Disconnect-MgGraph
Disconnect-MicrosoftTeams
