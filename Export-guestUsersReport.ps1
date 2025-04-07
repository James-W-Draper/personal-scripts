<#
=============================================================================================
.SYNOPSIS
    Export Office 365 guest user report with optional filters and group membership via Microsoft Graph.

.DESCRIPTION
    Connects to Microsoft Graph using delegated or certificate-based auth, retrieves all guest users,
    and outputs a CSV containing their name, UPN, company, email, account age, invitation state,
    creation details, and group memberships. Optionally filters stale or recently created accounts.

.PARAMETER StaleGuests
    Only show guests older than this many days.

.PARAMETER RecentlyCreatedGuests
    Only show guests newer than this many days.

.PARAMETER TenantId
    Your Azure AD tenant ID for app-based certificate login.

.PARAMETER ClientId
    App registration client ID for certificate-based auth.

.PARAMETER CertificateThumbprint
    Certificate thumbprint to authenticate the app.

.EXAMPLE
    .\Export-GuestUsersReport.ps1 -StaleGuests 90

.EXAMPLE
    .\Export-GuestUsersReport.ps1 -RecentlyCreatedGuests 7

.EXAMPLE
    .\Export-GuestUsersReport.ps1 -TenantId "<tenant-id>" -ClientId "<app-id>" -CertificateThumbprint "<thumbprint>"
=============================================================================================
#>
param (
    [Parameter(Mandatory = $false)]
    [int]$StaleGuests,
    [int]$RecentlyCreatedGuests,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

# Ensure Microsoft.Graph module is installed
if (-not (Get-Module Microsoft.Graph -ListAvailable)) {
    Write-Host "Microsoft Graph module is not installed." -ForegroundColor Yellow
    $confirm = Read-Host "Do you want to install it now? [Y/N]"
    if ($confirm -match '^[Yy]$') {
        Install-Module Microsoft.Graph -Scope CurrentUser -Force
    } else {
        Write-Host "Module is required. Exiting." -ForegroundColor Red
        exit
    }
}

# Connect to Microsoft Graph
if ($TenantId -and $ClientId -and $CertificateThumbprint) {
    Connect-MgGraph -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -ErrorAction Stop | Out-Null
} else {
    Connect-MgGraph -Scopes "Directory.Read.All" -ErrorAction Stop | Out-Null
}

Write-Host "Connected to Microsoft Graph successfully." -ForegroundColor Green
Set-MgProfile -Name beta

$guestUsers = Get-MgUser -All -Filter "UserType eq 'Guest'" -ExpandProperty MemberOf
$results = @()
$guestCount = 0
$printedGuests = 0

$ExportCSV = "GuestUserReport_$((Get-Date -Format 'yyyyMMdd_HHmm')).csv"
Write-Host "\nExporting report to $ExportCSV..."

foreach ($user in $guestUsers) {
    $guestCount++
    $displayName = $user.DisplayName
    $age = (New-TimeSpan -Start $user.CreatedDateTime).Days

    if ($StaleGuests -and $age -lt $StaleGuests) { continue }
    if ($RecentlyCreatedGuests -and $age -gt $RecentlyCreatedGuests) { continue }

    $groupMemberships = @($user.MemberOf.AdditionalProperties.displayName) -join ','
    if (-not $groupMemberships) { $groupMemberships = '-' }

    $company = if ($user.CompanyName) { $user.CompanyName } else { '-' }

    $results += [pscustomobject]@{
        DisplayName       = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        Company           = $company
        EmailAddress      = $user.Mail
        CreationTime      = $user.CreatedDateTime
        "AccountAge(days)" = $age
        CreationType      = $user.CreationType
        InvitationAccepted = $user.ExternalUserState
        GroupMembership   = $groupMemberships
    }
    $printedGuests++
    Write-Progress -Activity "Processing Guests" -Status "Currently Processing: $displayName" -PercentComplete (($guestCount / $guestUsers.Count) * 100)
}

$results | Export-Csv -Path $ExportCSV -NoTypeInformation -Encoding UTF8
Disconnect-MgGraph | Out-Null

if (Test-Path $ExportCSV) {
    Write-Host "\n✅ Report saved to $ExportCSV with $printedGuests guest(s)." -ForegroundColor Green
    $prompt = New-Object -ComObject wscript.shell
    $openFile = $prompt.popup("Do you want to open the report now?", 0, "Open Report", 4)
    if ($openFile -eq 6) {
        Invoke-Item $ExportCSV
    }
} else {
    Write-Host "No guests matched the specified criteria." -ForegroundColor Yellow
}
