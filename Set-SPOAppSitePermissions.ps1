<#
.SYNOPSIS
Assigns or removes 'Sites.Selected' permissions for a registered Azure AD application on a specified SharePoint Online site.

.DESCRIPTION
This script uses the Microsoft Graph PowerShell SDK to manage site-specific permissions for an Azure AD application that has the 'Sites.Selected' permission.
It securely prompts for a client secret and connects to Microsoft Graph to assign (or optionally remove) application-level permissions on a given SharePoint site.

.PARAMETER AppId
The Client ID of the Azure AD application to assign permissions to.

.PARAMETER SiteName
The name of the SharePoint Online site (e.g., "Project X").

.PARAMETER PermissionRole
The level of access to assign. Valid values: "read", "write", or "fullControl". Defaults to "write".

.PARAMETER TenantId
The Azure AD Tenant ID.

.PARAMETER ClientId
The Client ID used for authenticating with Microsoft Graph.

.PARAMETER RemovePermissions
Optional. If specified, removes existing permissions for the app on the site.

.EXAMPLE
.\Set-SPOAppSitePermissions.ps1 -AppId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -SiteName "Finance" -PermissionRole "read" -TenantId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" -ClientId "zzzzzzzz-zzzz-zzzz-zzzz-zzzzzzzzzzzz"

.NOTES
- Requires the Microsoft Graph PowerShell SDK and Az.Accounts module.
- Do not store client secrets in plain text. Use secure prompts or secret stores.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]$AppId,

    [Parameter(Mandatory)]
    [string]$SiteName,

    [Parameter()]
    [ValidateSet("read", "write", "fullControl")]
    [string]$PermissionRole = "write",

    [Parameter(Mandatory)]
    [string]$TenantId,

    [Parameter(Mandatory)]
    [string]$ClientId,

    [switch]$RemovePermissions
)

function Get-MgAccessToken {
    param (
        [string]$ClientId,
        [string]$TenantId
    )

    $ClientSecret = Read-Host -Prompt "Enter Client Secret" -AsSecureString
    $Credential = New-Object -TypeName PSCredential -ArgumentList $ClientId, $ClientSecret

    Connect-AzAccount -ServicePrincipal -TenantId $TenantId -Credential $Credential -ErrorAction Stop
    $Token = Get-AzAccessToken -ResourceTypeName MSGraph
    return $Token
}

# Connect to Microsoft Graph
$AccessToken = (Get-MgAccessToken -ClientId $ClientId -TenantId $TenantId).Token
Connect-MgGraph -AccessToken $AccessToken -NoWelcome

# Get SharePoint site by name
$Site = Get-MgSite -Search $SiteName
if (-not $Site) {
    Write-Error "Site '$SiteName' not found."
    exit 1
}

# Remove existing permissions if requested
if ($RemovePermissions.IsPresent) {
    Write-Host "Removing assigned permissions..." -ForegroundColor Yellow
    $Permissions = Get-MgSitePermission -SiteId $Site.Id
    foreach ($Perm in $Permissions) {
        Remove-MgSitePermission -SiteId $Site.Id -PermissionId $Perm.Id -ErrorAction SilentlyContinue
    }
    Write-Host "Permissions removed successfully." -ForegroundColor Green
    return
}

# Build permissions object
$Body = @{
    roles = @($PermissionRole)
    grantedToV2 = @{
        application = @{
            id = $AppId
        }
    }
}

Write-Host "Assigning '$PermissionRole' access to app $AppId on site '$($
