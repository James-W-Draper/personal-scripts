<#
.SYNOPSIS
    Creates a new app registration in Microsoft Entra ID (Azure AD), with optional service principal, client secret, and delegated API permissions.

.DESCRIPTION
    This script connects to Microsoft Graph using device code authentication, creates a new app registration
    with the provided display name and sign-in audience, creates a corresponding service principal 
    if applicable, optionally assigns common Microsoft Graph delegated API permissions, and optionally 
    creates a client secret for the app. The session can be left open for further work if desired.

.PARAMETER AppName
    The display name (friendly name) for the new app registration.

.PARAMETER SignInAudience
    Specifies who can sign in to the application. Options include:
    - AzureADMyOrg
    - AzureADMultipleOrgs
    - AzureADandPersonalMicrosoftAccount (default)
    - PersonalMicrosoftAccount

.PARAMETER AssignPermissions
    If specified, assigns common Microsoft Graph delegated API permissions (e.g., User.Read, Mail.Read).

.PARAMETER CreateClientSecret
    If specified, creates a client secret for the app and outputs the value.

.PARAMETER StayConnected
    If specified, the connection to Microsoft Graph remains open after the script finishes.

.EXAMPLE
    .\New-AppRegistrationWithSP.ps1 -AppName "My Custom App"

.EXAMPLE
    .\New-AppRegistrationWithSP.ps1 -AppName "Internal API" -SignInAudience "AzureADMyOrg" -AssignPermissions -CreateClientSecret -StayConnected
#>

param (
    [Parameter(Mandatory = $true, HelpMessage = "The friendly name of the app registration")]
    [string]$AppName,

    [Parameter(Mandatory = $false, HelpMessage = "The sign in audience for the app")]
    [ValidateSet("AzureADMyOrg", "AzureADMultipleOrgs", "AzureADandPersonalMicrosoftAccount", "PersonalMicrosoftAccount")]
    [string]$SignInAudience = "AzureADandPersonalMicrosoftAccount",

    [Parameter(Mandatory = $false, HelpMessage = "Assign Microsoft Graph delegated API permissions")]
    [switch]$AssignPermissions,

    [Parameter(Mandatory = $false, HelpMessage = "Create a client secret and display it")]
    [switch]$CreateClientSecret,

    [Parameter(Mandatory = $false, HelpMessage = "Leave the Microsoft Graph session open after the script finishes")]
    [switch]$StayConnected
)

$authTenant = switch ($SignInAudience) {
    "AzureADMyOrg" { "tenantId" }
    "AzureADMultipleOrgs" { "organizations" }
    "AzureADandPersonalMicrosoftAccount" { "common" }
    "PersonalMicrosoftAccount" { "consumers" }
    default { "invalid" }
}

if ($authTenant -eq "invalid") {
    Write-Host -ForegroundColor Red "❌ Invalid sign-in audience specified: $SignInAudience"
    exit 1
}

Connect-MgGraph -Scopes "Application.ReadWrite.All", "DelegatedPermissionGrant.ReadWrite.All", "AppRoleAssignment.ReadWrite.All", "User.Read" -UseDeviceAuthentication -ErrorAction Stop

$context = Get-MgContext -ErrorAction Stop
if ($authTenant -eq "tenantId") {
    $authTenant = $context.TenantId
}

$appRegistration = New-MgApplication -DisplayName $AppName -SignInAudience $SignInAudience `
    -IsFallbackPublicClient -ErrorAction Stop

Write-Host -ForegroundColor Cyan "✅ App registration created with App ID:" $appRegistration.AppId

if ($SignInAudience -ne "PersonalMicrosoftAccount") {
    New-MgServicePrincipal -AppId $appRegistration.AppId -ErrorAction SilentlyContinue -ErrorVariable SPError | Out-Null
    if ($SPError) {
        Write-Host -ForegroundColor Red "⚠️ Failed to create service principal for the app."
        Write-Host -ForegroundColor Red $SPError
        exit 1
    }
    Write-Host -ForegroundColor Cyan "✅ Service principal created successfully."
}

if ($AssignPermissions) {
    $permissions = @(
        @{ Id = "311a71cc-e848-46a1-bdf8-97ff7156d8e6"; Type = "Scope" },  # User.Read
        @{ Id = "570282fd-fa5c-430d-a7fd-fc8dc98a9dca"; Type = "Scope" }   # Mail.Read
    )
    Write-Host -ForegroundColor Cyan "Assigning Microsoft Graph delegated permissions..."
    Update-MgApplication -ApplicationId $appRegistration.Id -RequiredResourceAccess @(
        @{ ResourceAppId = "00000003-0000-0000-c000-000000000000"; ResourceAccess = $permissions }
    )
    Write-Host -ForegroundColor Green "✅ Delegated API permissions assigned."
}

if ($CreateClientSecret) {
    Write-Host -ForegroundColor Cyan "Creating client secret..."
    $secret = Add-MgApplicationPassword -ApplicationId $appRegistration.Id -PasswordCredential @{ DisplayName = "AppSecret" }
    Write-Host -ForegroundColor Green "✅ Client secret created."
    Write-Host -ForegroundColor Cyan "Client Secret Value: " -NoNewline
    Write-Host -ForegroundColor Yellow $secret.SecretText
}

Write-Host
Write-Host -ForegroundColor Green "🎉 SUCCESS"
Write-Host -ForegroundColor Cyan "Client ID: " -NoNewline
Write-Host -ForegroundColor Yellow $appRegistration.AppId
Write-Host -ForegroundColor Cyan "Auth Tenant: " -NoNewline
Write-Host -ForegroundColor Yellow $authTenant

if (-not $StayConnected) {
    Disconnect-MgGraph
    Write-Host -ForegroundColor Gray "Disconnected from Microsoft Graph."
} else {
    Write-Host
    Write-Host -ForegroundColor Yellow "⚠️ The connection to Microsoft Graph is still active."
    Write-Host -ForegroundColor Yellow "You can disconnect manually using: Disconnect-MgGraph"
}
