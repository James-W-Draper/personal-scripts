<#
.SYNOPSIS
Connects to Microsoft Graph using certificate-based authentication (CBA).

.DESCRIPTION
This script uses an application (client ID), tenant ID, and a certificate thumbprint
to authenticate to Microsoft Graph without user interaction (ideal for automation).

.NOTES
- Requires: Microsoft.Graph module
- Application must be registered in Entra ID (Azure AD) with API permissions
- Certificate must be installed in the CurrentUser\My or LocalMachine\My store
#>

# === CONFIGURATION ===
$ClientId              = "00000000-0000-0000-0000-000000000000"  # <-- Replace with your App Registration Client ID
$TenantId              = "00000000-0000-0000-0000-000000000000"  # <-- Replace with your Tenant ID
$CertificateThumbprint = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"  # <-- Replace with certificate thumbprint

# === CONNECT TO GRAPH ===
try {
    Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint
    Write-Host "✅ Connected to Microsoft Graph successfully." -ForegroundColor Green
} catch {
    Write-Error "❌ Failed to connect to Microsoft Graph: $($_.Exception.Message)"
}
