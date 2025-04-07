<#
.SYNOPSIS
Synchronises an AD group with users who are currently managers based on direct reports.

.DESCRIPTION
This script scans specified OUs for users with a defined `Manager` attribute, then adds those
managers to a given AD group and removes any existing group members who are no longer managers.

.PARAMETER Group
Distinguished name (DN) of the AD group to update.

.PARAMETER OUs
An array of distinguished names (DNs) of OUs from which to pull users.

.PARAMETER Verbose
Optional switch to output detailed processing information.

.NOTES
- Requires the ActiveDirectory module
- Must be run with sufficient permissions to read and write group membership
#>

param (
    [string]$Group = "CN=ManagersGroup,OU=Groups,DC=example,DC=com",
    [array]$OUs = @(
        "OU=Employees,DC=example,DC=com",
        "OU=Contractors,DC=example,DC=com"
    ),
    [switch]$Verbose
)

Import-Module ActiveDirectory

if ($Verbose) { $VerbosePreference = "Continue" } else { $VerbosePreference = "SilentlyContinue" }

# Validate AD group existence
try {
    Get-ADGroup -Identity $Group -ErrorAction Stop | Out-Null
} catch {
    Write-Warning "The specified AD group '$Group' does not exist. Please verify the group name."
    exit
}

# Fetch users from specified OUs
$Users = foreach ($OU in $OUs) {
    Write-Verbose "Fetching users from $OU..."
    try {
        Get-ADUser -Filter * -SearchBase $OU -Properties Manager
    } catch {
        Write-Warning "Failed to retrieve users from $OU. Error: $_"
    }
}

if (-not $Users) {
    Write-Warning "No users found in the specified OUs. Exiting script."
    exit
}

# Identify unique managers
$Managers = @()
$ManagerHashes = @{}

foreach ($User in $Users) {
    if ($User.Manager) {
        try {
            $Manager = Get-ADUser -Identity $User.Manager
            if (-not $ManagerHashes.ContainsKey($Manager.SamAccountName)) {
                $ManagerHashes[$
