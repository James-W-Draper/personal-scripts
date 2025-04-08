<#
.SYNOPSIS
    Exports all mailbox proxy addresses from Exchange Online to a timestamped CSV or Excel file.

.DESCRIPTION
    This script connects to Exchange Online, retrieves all user mailboxes, extracts their primary and proxy SMTP addresses,
    and exports them to a file. You can filter the export by domain and choose between CSV or Excel output.

.PARAMETER OutputPath
    The directory where the export file should be saved. Defaults to C:\Scripts if not specified.

.PARAMETER DomainFilter
    (Optional) A domain string to filter users. Only mailboxes with UPNs ending in this domain will be included.

.PARAMETER OutputExcel
    (Switch) If set, the output will be written to an Excel file (.xlsx) using ImportExcel. Otherwise, a CSV is created.

.EXAMPLE
    .\Export-ProxyAddresses.ps1 -OutputPath "C:\Exports" -DomainFilter "@contoso.com" -OutputExcel
#>

param(
    [string]$OutputPath = "C:\Scripts",
    [string]$DomainFilter,
    [switch]$OutputExcel
)

# Ensure ExchangeOnlineManagement is imported
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
}
Import-Module ExchangeOnlineManagement

# Ensure ImportExcel if needed
if ($OutputExcel -and -not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module ImportExcel -Scope CurrentUser -Force
}
if ($OutputExcel) {
    Import-Module ImportExcel
}

# Connect to Exchange Online
Connect-ExchangeOnline -ErrorAction Stop

# Get all user mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited

# Prepare array to store results
$proxyAddressesInfo = @()

foreach ($mailbox in $mailboxes) {
    if ($DomainFilter -and ($mailbox.UserPrincipalName -notlike "*${DomainFilter}")) {
        continue
    }

    $userProxyAddresses = @($mailbox.PrimarySmtpAddress.ToString())

    foreach ($proxy in $mailbox.EmailAddresses) {
        if ($proxy.PrefixString -eq "smtp") {
            $userProxyAddresses += $proxy.SmtpAddress
        }
    }

    $proxyAddressesInfo += [pscustomobject]@{
        DisplayName     = $mailbox.DisplayName
        UserPrincipalName = $mailbox.UserPrincipalName
        ProxyAddresses  = ($userProxyAddresses -join "; ")
    }
}

# Generate timestamped file name
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$fileNameBase = "ProxyAddresses_$timestamp"
$outputFile = Join-Path -Path $OutputPath -ChildPath ($fileNameBase + $(if ($OutputExcel) { ".xlsx" } else { ".csv" }))

# Export
if ($OutputExcel) {
    $proxyAddressesInfo | Export-Excel -Path $outputFile -WorksheetName "ProxyAddresses" -AutoSize
} else {
    $proxyAddressesInfo | Export-Csv -Path $outputFile -NoTypeInformation
}

# Disconnect
Disconnect-ExchangeOnline -Confirm:$false

Write-Host "Export completed. Output saved to: $outputFile" -ForegroundColor Green
