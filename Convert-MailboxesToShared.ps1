<#
.SYNOPSIS
Converts mailboxes of users in specified Active Directory OUs to shared mailboxes.

.DESCRIPTION
This script retrieves user accounts from one or more specified Organizational Units (OUs) in Active Directory,
filters those whose UserPrincipalName (UPN) ends with a specified domain suffix, and converts their Exchange Online
mailboxes to shared mailboxes.

.PARAMETER OUPaths
An array of distinguished names representing the OUs to search.

.PARAMETER DomainSuffix
The UPN suffix used to filter accounts (e.g., "@yourcompany.com").

.PARAMETER ShowExcluded
Switch to display UPNs that do not match the specified domain suffix.

.EXAMPLE
.\Convert-MailboxesToShared.ps1 -OUPaths "OU=Dept1,DC=domain,DC=local","OU=Dept2,DC=domain,DC=local" -DomainSuffix "@yourcompany.com"

.NOTES
- Requires ActiveDirectory and ExchangeOnlineManagement modules
- Must be run with sufficient permissions in both AD and Exchange Online
#>

param (
    [Parameter(Mandatory = $true)]
    [string[]]$OUPaths,

    [Parameter(Mandatory = $true)]
    [string]$DomainSuffix,

    [Parameter()]
    [switch]$ShowExcluded
)

# Ensure required modules are loaded
Import-Module ActiveDirectory -ErrorAction Stop
Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue

# Check connection to Exchange Online
if (-not (Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" })) {
    try {
        Connect-ExchangeOnline -ErrorAction Stop
        Write-Host "Connected to Exchange Online." -ForegroundColor Cyan
    } catch {
        Write-Host "Failed to connect to Exchange Online: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

# Collect all users from the provided OUs
$allUsers = foreach ($OU in $OUPaths) {
    Get-ADUser -SearchBase $OU -Filter * -Property UserPrincipalName
}

Write-Host "Total users retrieved from specified OUs: $($allUsers.Count)" -ForegroundColor Cyan

# Separate matching and non-matching users
$matchingUsers    = $allUsers | Where-Object { $_.UserPrincipalName -like "*$DomainSuffix" }
$nonMatchingUsers = $allUsers | Where-Object { $_.UserPrincipalName -notlike "*$DomainSuffix" }

# Optional: Show users who don't match the domain suffix
if ($ShowExcluded) {
    Write-Host "`nUPNs excluded (do not match '$DomainSuffix'):" -ForegroundColor Yellow
    $nonMatchingUsers | ForEach-Object {
        Write-Host $_.UserPrincipalName
    }
}

# Summary
Write-Host "`nTotal users matching '$DomainSuffix': $($matchingUsers.Count)" -ForegroundColor Green

# Convert mailboxes to shared
$convertedCount = 0
$index = 1

foreach ($user in $matchingUsers) {
    $upn = $user.UserPrincipalName
    try {
        $mailbox = Get-Mailbox -Identity $upn -ErrorAction Stop
        if ($mailbox.RecipientTypeDetails -ne "SharedMailbox") {
            Set-Mailbox -Identity $upn -Type Shared -ErrorAction Stop
            Write-Host "[$index/$($matchingUsers.Count)] Converted to shared: $upn"
            $convertedCount++
        } else {
            Write-Host "[$index/$($matchingUsers.Count)] Already shared: $upn"
        }
    } catch {
        Write-Host "[$index/$($matchingUsers.Count)] Failed for $upn - $($_.Exception.Message)" -ForegroundColor Red
    }
    $index++
}

Write-Host "`nConversion complete. $convertedCount out of $($matchingUsers.Count) mailboxes converted." -ForegroundColor Cyan

# Disconnect EXO session
Disconnect-ExchangeOnline -Confirm:$false
