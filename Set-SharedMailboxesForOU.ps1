<#
.SYNOPSIS
    This script connects to Exchange Online, retrieves user accounts from a specified 
    Organizational Unit (OU) in Active Directory, and sets the mailboxes of users whose 
    User Principal Name (UPN) ends with '@enstargroup.com' to shared mailboxes.

.DESCRIPTION
    - The script first checks if a session to Exchange Online is already active. If not, it connects to Exchange Online.
    - It then retrieves all user accounts from the specified OU in Active Directory.
    - The script filters these users based on their UPN to include only those ending with 
      '@enstargroup.com'.
    - It checks if each user's mailbox is already set to a shared mailbox. If not, the 
      mailbox is converted to a shared mailbox.
    - The script provides progress feedback on how many accounts have been discovered and 
      how many mailboxes have been actioned.

.PARAMETER OU
    The Organizational Unit (OU) path in Active Directory where user accounts are located.
    .\Set-SharedMailboxesForOU.ps1 -OU "OU=Sales,DC=example,DC=com"

.OUTPUTS
    Progress messages and a summary of how many mailboxes were converted to shared.

.EXAMPLE
    .\Set-SharedMailboxesForOU.ps1
#>

param (
    [Parameter(Mandatory)]
    [string]$OU = "OU=CoreSpec,OU=Group Users,DC=cwglobal,DC=local"
)

# Import the ExchangeOnlineManagement module if not already imported
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    try {
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -ErrorAction Stop
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
    } catch {
        Write-Host "Failed to install or import ExchangeOnlineManagement module: $($_.Exception.Message)" -ForegroundColor Red
        Exit
    }
} else {
    Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
}

# Function to check if already connected to Exchange Online
function Test-ExchangeOnlineConnection {
    try {
        # Check if any existing PSSession to Exchange Online exists
        $exoSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" }
        if ($exoSession) {
            Write-Host "Already connected to Exchange Online."
            return $true
        } else {
            return $false
        }
    } catch {
        return $false
    }
}

# Connect to EXO only if not already connected
if (-not (Test-ExchangeOnlineConnection)) {
    try {
        Connect-ExchangeOnline -ErrorAction Stop
        Write-Host "Connected to Exchange Online successfully."
    } catch {
        Write-Host "Failed to connect to Exchange Online: $($_.Exception.Message)" -ForegroundColor Red
        Exit
    }
}

# Get all users in the specified OU
try {
    $users = Get-ADUser -SearchBase $OU -Filter * -Property UserPrincipalName -ErrorAction Stop
    $totalUsers = $users.Count
    Write-Host "Users retrieved successfully from OU: $OU. Total users found: $totalUsers"
} catch {
    Write-Host "Error retrieving users: $($_.Exception.Message)" -ForegroundColor Red
    Exit
}

# Filter users whose UPN ends with "@enstargroup.com"
$filteredUsers = $users | Where-Object { $_.UserPrincipalName -like "*@enstargroup.com" }
$totalFilteredUsers = $filteredUsers.Count

# If no users are found, display a message and exit
if (-not $totalFilteredUsers) {
    Write-Host "No users found with UPN ending in '@enstargroup.com'." -ForegroundColor Yellow
    Exit
}

# Display the total number of users being processed
Write-Host "Total users to process: $totalFilteredUsers"

# Initialize action counter and progress index
$actionedCount = 0
$progressIndex = 1

# Loop through the filtered users and set their mailbox to shared if not already shared
$filteredUsers | ForEach-Object {
    $upn = $_.UserPrincipalName
    try {
        # Check mailbox type
        $mailbox = Get-Mailbox -Identity $upn -ErrorAction Stop
        if ($mailbox.RecipientTypeDetails -eq "SharedMailbox") {
            Write-Host "[$progressIndex/$totalFilteredUsers] Mailbox for $upn is already set to shared."
        } else {
            # Set mailbox type to shared if not already shared
            Set-Mailbox -Identity $upn -Type Shared -ErrorAction Stop
            Write-Host "[$progressIndex/$totalFilteredUsers] Successfully set mailbox type to shared for: $upn."
            $actionedCount++
        }
    } catch {
        Write-Host "[$progressIndex/$totalFilteredUsers] Error processing mailbox for: $upn - $($_.Exception.Message)" -ForegroundColor Red
    }

    # Update progress
    $progressIndex++
}

# Display completion message
Write-Host "Processing complete. $actionedCount out of $totalFilteredUsers mailboxes were changed to shared."

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
