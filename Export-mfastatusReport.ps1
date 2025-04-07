<#
.SYNOPSIS
    Export a report of licensed Azure AD users with and without MFA enabled using the MSOnline module.

.DESCRIPTION
    This script connects to Microsoft Online using the MSOnline module, retrieves all enabled and licensed users,
    separates them into MFA and non-MFA groups, and exports both reports to CSV files including metadata like
    display name, UPN, department, and timestamps for account creation and logins.

.EXAMPLE
    .\Export-MFAStatusReport.ps1

.NOTES
    Requires MSOnline module and admin credentials.
    Output saved to C:\scripts\NonMfaUsersReport.csv and C:\scripts\MfaUsersReport.csv
#>

# Import the MSOnline module
Import-Module MSOnline

# Connect to Microsoft Online
Connect-MsolService

# Retrieve all licensed and enabled users
$allUsers = Get-MsolUser -All | Where-Object { $_.IsLicensed -eq $true -and $_.BlockCredential -eq $false }

# --- Non-MFA Users ---
$nonMfaUsers = $allUsers | Where-Object { $_.StrongAuthenticationMethods.Count -eq 0 }
$outputFile = "C:\scripts\NonMfaUsersReport.csv"
$nonMfaUsers | Select-Object DisplayName, UserPrincipalName, Department, LastDirSyncTime, LastLogonTime, WhenCreated | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
Write-Output "Non-MFA Report generated: $outputFile"

# --- MFA-Enabled Users ---
$mfaUsers = $allUsers | Where-Object { $_.StrongAuthenticationMethods.Count -gt 0 }
$mfaDetails = foreach ($user in $mfaUsers) {
    foreach ($method in $user.StrongAuthenticationMethods) {
        [pscustomobject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            Department        = $user.Department
            LastDirSyncTime   = $user.LastDirSyncTime
            LastLogonTime     = $user.LastLogonTime
            WhenCreated       = $user.WhenCreated
            MFA_Method        = $method.MethodType
        }
    }
}
$mfaOutputFile = "C:\scripts\MfaUsersReport.csv"
$mfaDetails | Export-Csv -Path $mfaOutputFile -NoTypeInformation -Encoding UTF8
Write-Output "MFA Users Report generated: $mfaOutputFile"
