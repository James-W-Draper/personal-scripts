<#
.SYNOPSIS
    Retrieves a list of active Active Directory users whose accounts have expired.

.DESCRIPTION
    This script queries Active Directory for user accounts that are still enabled
    but have a set AccountExpirationDate that is in the past. It filters out any
    accounts that do not have an expiration date and displays the user's name,
    SamAccountName, and expiration date in a readable format.

.PARAMETER ExpiryThresholdDays
    Optional. Filters results to only include accounts that expired more than a certain
    number of days ago. For example, specify 30 to get users whose accounts expired over a month ago.

.PARAMETER ExportCsvPath
    Optional. If provided, the results will also be exported to the specified CSV path.

.EXAMPLE
    .\Get-ExpiredADUsers.ps1

.EXAMPLE
    .\Get-ExpiredADUsers.ps1 -ExpiryThresholdDays 30

.EXAMPLE
    .\Get-ExpiredADUsers.ps1 -ExportCsvPath "C:\Reports\ExpiredUsers.csv"

.NOTES
    Requires ActiveDirectory module.
    Must be run with appropriate privileges to query user accounts in AD.
#>

[CmdletBinding()]
param (
    [int]$ExpiryThresholdDays,
    [string]$ExportCsvPath
)

# Import the Active Directory module if not already loaded
if (-not (Get-Module -Name ActiveDirectory)) {
    Import-Module ActiveDirectory -ErrorAction Stop
}

# Calculate threshold date if filtering by number of days
$cutoffDate = if ($PSBoundParameters.ContainsKey('ExpiryThresholdDays')) {
    (Get-Date).AddDays(-$ExpiryThresholdDays)
} else {
    Get-Date
}

# Get enabled users with an expired AccountExpirationDate
$expiredUsers = Get-ADUser -Filter {
    Enabled -eq $true -and AccountExpirationDate -lt $cutoffDate
} -Properties Name, SamAccountName, AccountExpirationDate

# Filter out users with no expiration date set
$filteredUsers = $expiredUsers | Where-Object { $_.AccountExpirationDate -ne $null }

# Select and format output
$report = $filteredUsers | Select-Object Name, SamAccountName, @{
    Name       = "AccountExpirationDate"
    Expression = { $_.AccountExpirationDate.ToString("yyyy-MM-dd") }
} | Sort-Object AccountExpirationDate

# Output to screen
$report

# Optionally export to CSV
if ($ExportCsvPath) {
    try {
        $report | Export-Csv -Path $ExportCsvPath -NoTypeInformation -Encoding UTF8
        Write-Host "Report exported to: $ExportCsvPath" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to export CSV: $_"
    }
}
