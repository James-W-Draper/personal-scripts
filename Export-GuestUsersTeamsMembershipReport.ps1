<#
.SYNOPSIS
Exports a detailed report of Microsoft 365 guest users and their Microsoft Teams memberships.

.DESCRIPTION
This script connects to Microsoft Graph and Microsoft Teams, retrieves all guest users from Azure AD,
and exports their identity, creation, sign-in details, and Teams membership into an Excel report.

Includes:
- Guest user metadata: UPN, Display Name, Created Date, Sign-In Timestamps
- Microsoft Teams memberships
- Optional export path configuration

Requires:
- Microsoft Graph PowerShell SDK
- MicrosoftTeams module
- Exchange Online module (if extended further)
- ImportExcel module

.NOTES
Author: James Draper
#>

param (
    [string]$OutputPath = "C:\Scripts",
    [string]$FilePrefix = "GuestUserTeamsReport"
)

# Connect to Microsoft 365 services
Connect-ExchangeOnline -ErrorAction SilentlyContinue
Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All"
Select-MgProfile beta
Connect-MicrosoftTeams -ErrorAction SilentlyContinue

# Retrieve all guest users and Teams membership data
$GuestUsersWithTeams = Get-MgUser -Filter "userType eq 'Guest'" -ConsistencyLevel eventual -All -Property @(
    'UserPrincipalName',
    'SignInActivity',
    'CreatedDateTime',
    'DisplayName',
    'Mail',
    'SignInSessionsValidFromDateTime',
    'RefreshTokensValidFromDateTime',
    'id'
) | Select-Object @(
    'UserPrincipalName',
    'CreatedDateTime',
    'DisplayName',
    'Mail',
    'SignInSessionsValidFromDateTime',
    'RefreshTokensValidFromDateTime',
    'id',
    @{Name = 'LastSignInDateTime'; Expression = { [datetime]$_.SignInActivity.LastSignInDateTime }}
) | ForEach-Object {
    $user = $_
    try {
        $teamsMemberships = Get-TeamUser -User $user.UserPrincipalName -ErrorAction Stop
        foreach ($membership in $teamsMemberships) {
            [PSCustomObject]@{
                UserPrincipalName              = $user.UserPrincipalName
                CreatedDateTime                = $user.CreatedDateTime
                DisplayName                    = $user.DisplayName
                Mail                           = $user.Mail
                SignInSessionsValidFromDateTime = $user.SignInSessionsValidFromDateTime
                RefreshTokensValidFromDateTime  = $user.RefreshTokensValidFromDateTime
                LastSignInDateTime             = $user.LastSignInDateTime
                TeamDisplayName                = $membership.DisplayName
            }
        }
    } catch {
        [PSCustomObject]@{
            UserPrincipalName              = $user.UserPrincipalName
            CreatedDateTime                = $user.CreatedDateTime
            DisplayName                    = $user.DisplayName
            Mail                           = $user.Mail
            SignInSessionsValidFromDateTime = $user.SignInSessionsValidFromDateTime
            RefreshTokensValidFromDateTime  = $user.RefreshTokensValidFromDateTime
            LastSignInDateTime             = $user.LastSignInDateTime
            TeamDisplayName                = "<No Teams Membership>"
        }
    }
}

# Export settings for Excel
$ExcelParams = @{
    BoldTopRow   = $true
    AutoSize     = $true
    AutoFilter   = $true
    FreezeTopRow = $true
}

# Format file name
$Timestamp = Get-Date -Format "yyyyMMddTHHmmss"
$ExcelFilePath = Join-Path -Path $OutputPath -ChildPath "${FilePrefix}_${Timestamp}.xlsx"

# Export the report
$GuestUsersWithTeams | Sort-Object UserPrincipalName | Export-Excel -Path $ExcelFilePath -WorksheetName "GuestUsers" @ExcelParams

Write-Host "Guest user Teams membership report exported to: $ExcelFilePath"
