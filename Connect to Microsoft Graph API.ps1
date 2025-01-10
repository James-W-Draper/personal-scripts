# Connect to Microsoft Graph API
# https://o365info.com/connect-microsoft-graph-powershell/

# Configuration
$ClientId = "a10b4f3b-a29c-4742-b5ad-b26c304a1011"
$TenantId = "dbda57bd-564a-4ae2-b756-24442e84ba38"
$CertificateThumbprint = "EF3F27F50744E2D20E49AD8C14DA24FD3634052C"

# Connect to Microsoft Graph with CBA
Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint