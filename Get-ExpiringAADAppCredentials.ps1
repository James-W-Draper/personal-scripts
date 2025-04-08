<#
.SYNOPSIS
Exports a report of Azure AD application secrets and certificates that are expired or expiring soon.

.DESCRIPTION
This script connects to Azure Active Directory using the Az module, retrieves all app registrations, and checks their associated client secrets and certificates.
It then exports a report to Excel using ImportExcel, highlighting credentials that are expired or will expire within 30 days.

.PARAMETER OutputPath
Optional. Full path for the output Excel file. Defaults to "C:\scripts\ExpiringAADAppCredentials_<date>.xlsx".

.EXAMPLE
.\Get-ExpiringAADAppCredentials.ps1 -OutputPath "D:\Reports\AADAppCreds_20250408.xlsx"

.NOTES
Requires:
- Az.Accounts module
- Az.Resources module
- ImportExcel module
#>

[CmdletBinding()]
param (
    [string]$OutputPath
)

# Set default output path with today's date if not provided
if (-not $OutputPath) {
    $DateStamp = Get-Date -Format "yyyyMMdd"
    $OutputPath = "C:\scripts\ExpiringAADAppCredentials_$DateStamp.xlsx"
}

# Ensure output directory exists
$OutputDir = Split-Path -Path $OutputPath -Parent
if (-not (Test-Path $OutputDir)) {
    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
}

# Import required modules
foreach ($module in @('Az.Accounts', 'Az.Resources', 'ImportExcel')) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Install-Module -Name $module -Scope CurrentUser -Force
    }
    Import-Module $module -ErrorAction Stop
}

# Connect to Azure
try {
    Connect-AzAccount -ErrorAction Stop
} catch {
    Write-Error "❌ Failed to connect to Azure: $_"
    return
}

# Set the window for flagging upcoming expirations
$Now = Get-Date
$Soon = $Now.AddDays(30)

$CredentialReport = @()

# Get all App Registrations
$Apps = Get-AzADApplication

foreach ($App in $Apps) {
    # Check client secrets
    foreach ($Secret in $App.PasswordCredentials) {
        $CredentialReport += [PSCustomObject]@{
            AppDisplayName   = $App.DisplayName
            AppId            = $App.AppId
            CredentialType   = 'Client Secret'
            StartDate        = $Secret.StartDateTime
            EndDate          = $Secret.EndDateTime
            Expired          = $Secret.EndDateTime -lt $Now
            ExpiringSoon     = ($Secret.EndDateTime -ge $Now) -and ($Secret.EndDateTime -le $Soon)
        }
    }

    # Check certificates
    foreach ($Cert in $App.KeyCredentials) {
        $CredentialReport += [PSCustomObject]@{
            AppDisplayName   = $App.DisplayName
            AppId            = $App.AppId
            CredentialType   = 'Certificate'
            StartDate        = $Cert.StartDateTime
            EndDate          = $Cert.EndDateTime
            Expired          = $Cert.EndDateTime -lt $Now
            ExpiringSoon     = ($Cert.EndDateTime -ge $Now) -and ($Cert.EndDateTime -le $Soon)
        }
    }
}

# Export to Excel
$CredentialReport | Export-Excel -Path $OutputPath -WorksheetName "AppCreds" -BoldTopRow -AutoSize -AutoFilter -FreezeTopRow

Write-Host "✅ Report exported to $OutputPath" -ForegroundColor Green
