# SharePoint Online Site Selected Permission

# Varieables
$AppId = "076af198-2122-4066-89e3-0343f06e13d9"
$PermissionRole = "write"  # Change to "read", "write", or "fullControl"
$SiteName = "DOCOsoftSIT"
$SiteName = "DOCOSIT2"
###  Azure AD Application Details - IT ONLY  ###
$TenantId = "dbda57bd-564a-4ae2-b756-24442e84ba38"
$ClientId = "a10b4f3b-a29c-4742-b5ad-b26c304a1011"
# ClientSecret = "9MO8Q~d1w-wMH6DKzsbDFyhkBbcVm2WY79nU6awe"  # Store securely in production

function Get_Mg_Access_Token ($ClientId, $TenantId) {
    $ClientSecret = Read-Host -Prompt "Input Client Secret" -AsSecureString
    $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $ClientSecret
    Connect-AzAccount -ServicePrincipal -TenantId $TenantId -Credential $Credential
    $AzAccessToken = Get-AzAccessToken -AsSecureString -ResourceTypeName MSGraph
    return $AzAccessToken
} # Function get_token
Connect-MgGraph -AccessToken (Get_Mg_Access_Token $ClientId $TenantId).token -NoWelcome

# Get Site
$Site = Get-MgSite -Search $SiteName

# Set Site Permission
$params = @{ roles = @($PermissionRole); grantedToV2 = @{ application = @{ id = $AppId } } }
New-MgSitePermission -SiteId $Site.Id -BodyParameter $params

# Get Site Permissions
Write-Host "Site - $($Site.DisplayName)" -BackgroundColor Green -ForegroundColor Black -NoNewline
(Get-MgSitePermission -SiteId $Site.Id) | Select-Object @{n = "PermissionID"; e = { $_.Id } }, @{n = "AppName"; e = { $_.GrantedToIdentitiesV2.Application.DisplayName } }, @{n = "AppID"; e = { $_.GrantedToIdentitiesV2.Application.Id } }, Roles
#>

# Remove Permission
$PermissionId = (Get-MgSitePermission -SiteId $Site.Id).Id
Remove-MgSitePermission -SiteId $Site.Id -PermissionId $PermissionId
#>

























<#
# Application to Assign Permissions To
#$AppDisplayName = "EIP - Docosoft SIT 1"
# SharePoint Site Details
#$SharePointDomain = "enstargroup.sharepoint.com"
#>
