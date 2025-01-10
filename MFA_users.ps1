# Import the MSOnline module 
# This module provides cmdlets for Microsoft Online Services, allowing us to manage users and settings in Azure AD/Microsoft 365.
Import-Module MSOnline

# Connect to Microsoft Online
# Prompts for admin credentials to establish a session with Microsoft Online.
Connect-MsolService

# Get all enabled users in the directory
# Retrieves a list of enabled users in Azure AD, including both licensed and unlicensed users.
$allUsers = Get-MsolUser -All | Where-Object { $_.IsLicensed -eq $true -and $_.BlockCredential -eq $false }

# Filter users without MFA enabled
# This filters out users who have no multi-factor authentication (MFA) methods set up.
# Users without any StrongAuthenticationMethods are considered to not have MFA enabled.
$nonMfaUsers = $allUsers | Where-Object { $_.StrongAuthenticationMethods.Count -eq 0 }

# Output the non-MFA users to a CSV file
# Creates a CSV file with details of users who do not have MFA enabled, including their last logon time and account creation date.
$outputFile = "C:\scripts\NonMfaUsersReport.csv"
$nonMfaUsers | Select-Object DisplayName, UserPrincipalName, Department, LastDirSyncTime, LastLogonTime, WhenCreated | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8

# Output a message confirming the non-MFA report was generated
Write-Output "Report generated: $outputFile"

# Filter users with MFA enabled
# This filters for users who have one or more MFA methods configured (StrongAuthenticationMethods.Count > 0).
$mfaUsers = $allUsers | Where-Object { $_.StrongAuthenticationMethods.Count -gt 0 }

# Extract MFA methods for each user and create a detailed list
# For each user with MFA enabled, this block gathers specific MFA methods in use and stores each as a custom object.
$mfaDetails = foreach ($user in $mfaUsers) {
    foreach ($method in $user.StrongAuthenticationMethods) {
        [pscustomobject]@{
            DisplayName         = $user.DisplayName       # User's display name for easy identification
            UserPrincipalName   = $user.UserPrincipalName # The user's principal name, often their login email
            Department          = $user.Department        # Department or organizational unit, if available
            LastDirSyncTime     = $user.LastDirSyncTime   # Time of last directory sync, helpful for synced environments
            LastLogonTime       = $user.LastLogonTime     # User's last logon time
            WhenCreated         = $user.WhenCreated       # Account creation date
            MFA_Method          = $method.MethodType      # Specific MFA method enabled for the user (e.g., Phone, Authenticator app)
        }
    }
}

# Export MFA-enabled users with details to a CSV file
# Creates a CSV file with details of users who have MFA enabled, along with their specific MFA methods and account creation date.
$mfaOutputFile = "C:\scripts\MfaUsersReport.csv"
$mfaDetails | Export-Csv -Path $mfaOutputFile -NoTypeInformation -Encoding UTF8

# Output a message confirming the MFA report was generated
Write-Output "MFA Users Report generated: $mfaOutputFile"
