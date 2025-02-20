# PowerShell Script to Manage SharePoint Online 'Sites.Selected' Permissions
# This script interacts with Microsoft Graph API to assign and manage SharePoint site permissions for a registered Azure AD application.
#
# Prerequisites:
# - Azure AD application with 'Sites.Selected' permission granted.
# - Microsoft Graph PowerShell SDK installed.
# - Proper role-based access control (RBAC) configured for the user executing this script.

# Define variables
$AppId = "<YOUR-APPLICATION-ID>"  # Replace with the App Registration's Client ID
$PermissionRole = "write"  # Options: "read", "write", or "fullControl"
$SiteName = "<YOUR-SITE-NAME>"  # Replace with your SharePoint Online site name

###  Azure AD Application Details (DO NOT STORE SECRETS IN PLAIN TEXT) ###
$TenantId = "<YOUR-TENANT-ID>"  # Replace with your Azure AD Tenant ID
$ClientId = "<YOUR-CLIENT-ID>"  # Replace with your Client ID
# Note: ClientSecret should be securely stored, e.g., in Azure Key Vault or environment variables

# Function to retrieve an access token for Microsoft Graph API
function Get_Mg_Access_Token ($ClientId, $TenantId) {
    # Prompt user for Client Secret securely
    $ClientSecret = Read-Host -Prompt "Enter Client Secret" -AsSecureString
    
    # Create credential object
    $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $ClientSecret
    
    # Authenticate with Azure using the provided Service Principal credentials
    Connect-AzAccount -ServicePrincipal -TenantId $TenantId -Credential $Credential
    
    # Retrieve Microsoft Graph access token
    $AzAccessToken = Get-AzAccessToken -ResourceTypeName MSGraph
    return $AzAccessToken
} # End of function Get_Mg_Access_Token

# Connect to Microsoft Graph API using the retrieved access token
Connect-MgGraph -AccessToken (Get_Mg_Access_Token $ClientId $TenantId).Token -NoWelcome

# Retrieve the target SharePoint site
$Site = Get-MgSite -Search $SiteName
if (-not $Site) {
    Write-Host "Error: SharePoint site '$SiteName' not found." -ForegroundColor Red
    exit 1
}

# Define permission parameters
$params = @{ 
    roles = @($PermissionRole)  # Assigning role-based permissions
    grantedToV2 = @{ application = @{ id = $AppId } }  # Grant access to the specified application
}

# Apply site permissions
Write-Host "Assigning '$PermissionRole' permissions to App ID: $AppId for site: $($Site.DisplayName)" -ForegroundColor Cyan
New-MgSitePermission -SiteId $Site.Id -BodyParameter $params

# Retrieve and display assigned site permissions
Write-Host "Current permissions for site: $($Site.DisplayName)" -ForegroundColor Green
(Get-MgSitePermission -SiteId $Site.Id) | Select-Object @{n="PermissionID"; e={ $_.Id } }, 
                                                          @{n="AppName"; e={ $_.GrantedToIdentitiesV2.Application.DisplayName } }, 
                                                          @{n="AppID"; e={ $_.GrantedToIdentitiesV2.Application.Id } }, 
                                                          Roles

# Optional: Remove assigned permissions (Use with caution!)
if ($Confirm -eq $true) {
    Write-Host "Removing assigned permissions..." -ForegroundColor Yellow
    $PermissionId = (Get-MgSitePermission -SiteId $Site.Id).Id
    Remove-MgSitePermission -SiteId $Site.Id -PermissionId $PermissionId
    Write-Host "Permissions removed successfully." -ForegroundColor Green
}

# Script complete
Write-Host "Script execution completed." -ForegroundColor Magenta
