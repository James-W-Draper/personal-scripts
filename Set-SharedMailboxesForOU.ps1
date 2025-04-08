<#
.SYNOPSIS
Converts mailboxes of users in a specified Active Directory OU to shared mailboxes if they belong to a target domain.

.DESCRIPTION
This script connects to Exchange Online and retrieves user accounts from a specified Active Directory Organizational Unit (OU).
It filters accounts whose UserPrincipalName ends with "@enstargroup.com" and converts their mailboxes to shared type
if not already set. The script provides real-time progress and a final summary.

.PARAMETER OU
The distinguished name of the Active Directory Organizational Unit containing user accounts.

.EXAMPLE
.\Set-SharedMailboxesForOU.ps1 -OU "OU=Finance,DC=example,DC=com"

.NOTES
- Requires the ActiveDirectory and ExchangeOnlineManagement modules
- Requires sufficient Exchange Online and AD permissions
- Created by: [Your Name Here]
- Version: 1.0
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$OU
)

# Constants
$TargetDomain = "@enstargroup.com"

# Ensure the required module is available
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    try {
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -ErrorAction Stop
    } catch {
        Write-Host "Failed to install ExchangeOnlineManagement module: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}
Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue

# Check if already connected to EXO
function Test-ExchangeOnlineConnection {
    try {
        return (Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" }).Count -gt 0
    } catch {
        return $false
    }
}

# Connect to Exchange Online if needed
if (-not (Test-ExchangeOnlineConnection)) {
    try {
        Connect-ExchangeOnline -ErrorAction Stop
        Write-Host "Connected to Exchange Online successfully."
    } catch {
        Write-Host "Failed to connect to Exchange Online: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "Already connected to Exchange Online."
}

# Get users from specified OU
try {
    $users = Get-ADUser -SearchBase $OU -Filter * -Property UserPrincipalName
    Write-Host "Retrieved $($users.Count) users from OU: $OU"
} catch {
    Write-Host "Failed to retrieve users: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Filter by UPN domain
$filteredUsers = $users | Where-Object { $_.UserPrincipalName -like "*$TargetDomain" }
if (-not $filteredUsers) {
    Write-Host "No users found with UPN ending in '$TargetDomain'" -ForegroundColor Yellow
    exit 0
}

Write-Host "Total users to process: $($filteredUsers.Count)"

# Process users
$convertedCount = 0
$index = 1

foreach ($user in $filteredUsers) {
    $upn = $user.UserPrincipalName
    try {
        $mailbox = Get-Mailbox -Identity $upn -ErrorAction Stop
        if ($mailbox.RecipientTypeDetails -ne 'SharedMailbox') {
            Set-Mailbox -Identity $upn -Type Shared -ErrorAction Stop
            Write-Host "[$index/$($filteredUsers.Count)] Converted: $upn to Shared Mailbox"
            $convertedCount++
        } else {
            Write-Host "[$index/$($filteredUsers.Count)] Already shared: $upn"
        }
    } catch {
        Write-Host "[$index/$($filteredUsers.Count)] Error processing $upn - $($_.Exception.Message)" -ForegroundColor Red
    }
    $index++
}

Write-Host "`nCompleted: $convertedCount out of $($filteredUsers.Count) mailboxes converted to Shared."

# Disconnect cleanly
Disconnect-ExchangeOnline -Confirm:$false
