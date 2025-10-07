<#
.SYNOPSIS
Export SharePoint Online sites with owners and Teams linkage.


.DESCRIPTION
- Enumerates all SPO site collections (excludes OneDrive unless -IncludeOneDrive).
- Group-connected sites: pulls M365 Group owners via Graph, falling back to UPN when Mail is empty.
- Classic sites: reads SharePoint "Owners" group via PnP, falling back to UPN/LoginName when Email is empty.
- Flags if site has a Microsoft Teams team.
- Caches PnP connections per site to avoid repeated interactive prompts.
- Outputs CSV by default; can write .xlsx with -Excel if ImportExcel is available.


.REQUIREMENTS
Modules:
  - Microsoft.Online.SharePoint.PowerShell
  - Microsoft.Graph
  - PnP.PowerShell
  - (Optional) ImportExcel


Graph delegated scopes required:
  - Group.Read.All
  - Directory.Read.All
  - Team.ReadBasic.All
#>


<#
.EXAMPLE
.\Export-SPO-Sites-Owners.ps1 -Tenant "contoso"


Lists all SharePoint Online sites in the "contoso" tenant (excluding OneDrive).  
Exports results to SPO-Sites-Owners.csv in the current folder.


.EXAMPLE
.\Export-SPO-Sites-Owners.ps1 -Tenant "contoso" -OutputPath "C:\Reports\Sites.csv"


Same as above, but writes the CSV to C:\Reports\Sites.csv.


.EXAMPLE
.\Export-SPO-Sites-Owners.ps1 -Tenant "contoso" -IncludeOneDrive


Includes OneDrive for Business personal sites (SPSPERS#0).  
Useful if you want a full tenant inventory including personal storage.


.EXAMPLE
.\Export-SPO-Sites-Owners.ps1 -Tenant "contoso" -Excel


Exports to Excel (.xlsx) instead of CSV (requires the ImportExcel module).  
File is created as SPO-Sites-Owners.xlsx in the current folder by default.


.EXAMPLE
.\Export-SPO-Sites-Owners.ps1 -Tenant "contoso" -AdminUrl "https://contoso-admin.sharepoint.de"


Specifies a custom SharePoint admin URL (needed for some sovereign clouds, e.g. Germany, GCC High, China).
#>



param(
  [Parameter(Mandatory = $true)]
  [string]$Tenant,     
  <#
    The tenant short name (no spaces). 
    Example: "contoso" -> resolves to https://contoso-admin.sharepoint.com
  #>


  [Parameter(Mandatory = $false)]
  [string]$AdminUrl,   
  <#
    Optional explicit SharePoint admin URL. 
    If not provided, it is built from $Tenant (https://<tenant>-admin.sharepoint.com).
    Use this if your tenant has a custom admin domain.
  #>


  [Parameter(Mandatory = $false)]
  [string]$OutputPath = ".\SPO-Sites-Owners.csv",
  <#
    Path for output file.
    Defaults to a CSV in the current folder. 
    If -Excel is supplied and ImportExcel module is present, 
    this path will be used with .xlsx extension.
  #>


  [switch]$IncludeOneDrive,
  <#
    By default, OneDrive personal sites (SPSPERS#0) are excluded.
    Add this switch to include them.
  #>


  [switch]$Excel
  <#
    If specified, output will be exported as Excel (.xlsx) 
    instead of CSV (requires ImportExcel module).
  #>
)



# ---------- Helpers ----------
function Ensure-Module {
  param([string]$Name, [string]$MinVersion = "0.0.0")
  if (-not (Get-Module -ListAvailable -Name $Name)) {
    Write-Host "Installing module $Name..." -ForegroundColor Yellow
    Install-Module $Name -Scope CurrentUser -Force -AllowClobber -MinimumVersion $MinVersion -ErrorAction Stop
  }
  Import-Module $Name -ErrorAction Stop
}


function Resolve-Display {
  param(
    [string]$DisplayName,
    [string]$Mail,
    [string]$UserPrincipalName,
    [string]$LoginName
  )
  $addr = $null
  if ($Mail) { $addr = $Mail }
  elseif ($UserPrincipalName) { $addr = $UserPrincipalName }
  elseif ($LoginName) {
    # Try to extract UPN from claims or legacy login names
    if ($LoginName -match 'i:0#\.f\|membership\|(.+)$') { $addr = $Matches[1] }
    elseif ($LoginName -match 'i:0#.w\|[^\]+\(.+)$') { $addr = $Matches[1] }
    else { $addr = $LoginName }
  }
  if ($DisplayName) {
    if ($addr) { return ('{0} <{1}>' -f $DisplayName, $addr) }
    else { return $DisplayName }
  } else {
    return ($addr ?? "")
  }
}


# Cache PnP connections per site URL to avoid repeated prompts
$script:PnPConnections = @{}
function Get-PnPConnectionCached {
  param([string]$SiteUrl)
  if ($script:PnPConnections.ContainsKey($SiteUrl)) {
    return $script:PnPConnections[$SiteUrl]
  }
  # -Interactive reuses token cache, so you should only be prompted once per run
  $conn = Connect-PnPOnline -Url $SiteUrl -Interactive -ReturnConnection -ErrorAction Stop
  $script:PnPConnections[$SiteUrl] = $conn
  return $conn
}


function Get-GroupOwnersText {
  param([Guid]$GroupId)
  try {
    $owners = Get-MgGroupOwner -GroupId $GroupId -All -ErrorAction Stop
    if (-not $owners) { return "" }
    $owners |
      ForEach-Object {
        switch ($_. '@odata.type') {
          "#microsoft.graph.user" {
            Resolve-Display -DisplayName $_.DisplayName -Mail $_.Mail -UserPrincipalName $_.UserPrincipalName
          }
          "#microsoft.graph.servicePrincipal" {
            "[App] $($_.DisplayName)"
          }
          default {
            $_.DisplayName
          }
        }
      } |
      Where-Object { $_ } |
      -join "; "
  } catch {
    Write-Warning "Failed to get owners for Group $GroupId - $($_.Exception.Message)"
    return ""
  }
}


function Test-HasTeam {
  param([Guid]$GroupId)
  try {
    $null = Get-MgGroupTeam -GroupId $GroupId -ErrorAction Stop
    return $true
  } catch { return $false }
}


function Get-ClassicOwnersText {
  param([string]$SiteUrl)
  try {
    $conn = Get-PnPConnectionCached -SiteUrl $SiteUrl
    $web = Get-PnPWeb -Includes AssociatedOwnerGroup -Connection $conn
    if (-not $web.AssociatedOwnerGroup) { return "" }
    $grpId = $web.AssociatedOwnerGroup.Id
    $members = Get-PnPGroupMember -Identity $grpId -Connection $conn -ErrorAction Stop
    if (-not $members) { return "" }
    $members |
      ForEach-Object {
        Resolve-Display -DisplayName $_.Title -Mail $_.Email -UserPrincipalName $_.UserPrincipalName -LoginName $_.LoginName
      } |
      Where-Object { $_ } |
      -join "; "
  } catch {
    Write-Warning "Failed to get classic owners for $SiteUrl - $($_.Exception.Message)"
    return ""
  }
}


# ---------- Bootstrapping ----------
Ensure-Module -Name Microsoft.Online.SharePoint.PowerShell
Ensure-Module -Name Microsoft.Graph
Ensure-Module -Name PnP.PowerShell
if ($Excel) { Ensure-Module -Name ImportExcel }


# Sanitize Tenant and compute AdminUrl if missing
$Tenant = ($Tenant -replace '\s', '')
if (-not $Tenant) { throw "Tenant cannot be empty after sanitisation." }
if (-not $AdminUrl) { $AdminUrl = "https://$Tenant-admin.sharepoint.com" }


Write-Host "Connecting to SharePoint Admin: $AdminUrl" -ForegroundColor Cyan
Connect-SPOService -Url $AdminUrl


Write-Host "Connecting to Microsoft Graph (interactive)..." -ForegroundColor Cyan
$scopes = @("Group.Read.All","Directory.Read.All","Team.ReadBasic.All")
Connect-MgGraph -Scopes $scopes | Out-Null
Select-MgProfile -Name "v1.0"


# ---------- Enumerate sites ----------
Write-Host "Retrieving SharePoint sites..." -ForegroundColor Cyan
$allSites = Get-SPOSite -Limit All -Detailed


if (-not $IncludeOneDrive) {
  $allSites = $allSites | Where-Object {
    $_.Template -ne 'SPSPERS#0' -and $_.Url -notmatch '-my\.sharepoint\.com'
  }
}


# ---------- Process ----------
$results = New-Object System.Collections.Generic.List[Object]
$counter = 0


foreach ($site in $allSites) {
  $counter++
  Write-Host ("[{0}/{1}] {2}" -f $counter, $allSites.Count, $site.Url)


  $groupId = $null
  $isGroupConnected = $false
  $hasTeam = $false
  $ownersText = ""


  if ($site.PSObject.Properties.Name -contains "GroupId" -and $site.GroupId) {
    $groupId = [Guid]$site.GroupId
    $isGroupConnected = $true
  } elseif ($site.Template -like "GROUP#0*") {
    $isGroupConnected = $true
    # Best-effort GroupId discovery - optional and harmless if it fails
    try {
      $mgSite = Get-MgSite -SiteId $site.Url -ErrorAction Stop
      $drive = Get-MgSiteDrive -SiteId $mgSite.Id -ErrorAction Stop
      if ($drive.Owner -and $drive.Owner.Group -and $drive.Owner.Group.Id) {
        $groupId = [Guid]$drive.Owner.Group.Id
      }
    } catch { }
  }


  if ($isGroupConnected -and $groupId) {
    $ownersText = Get-GroupOwnersText -GroupId $groupId
    $hasTeam = Test-HasTeam -GroupId $groupId
  } elseif (-not $isGroupConnected) {
    $ownersText = Get-ClassicOwnersText -SiteUrl $site.Url
  }


  $results.Add([pscustomobject]@{
    Title            = $site.Title
    Url              = $site.Url
    Template         = $site.Template
    StorageQuotaMB   = $site.StorageQuota
    StorageUsageMB   = $site.StorageUsageCurrent
    PrimaryAdmin     = $site.Owner
    SecondaryAdmins  = ($site.AdditionalAdministrators -join "; ")
    IsGroupConnected = $isGroupConnected
    GroupId          = if ($groupId) { $groupId } else { $null }
    HasTeamsTeam     = $hasTeam
    Owners           = $ownersText
  }) | Out-Null
}


# ---------- Export ----------
if ($Excel -and (Get-Module -ListAvailable -Name ImportExcel)) {
  $xlsx = [System.IO.Path]::ChangeExtension($OutputPath, ".xlsx")
  $results | Export-Excel -Path $xlsx -WorksheetName "Sites" -AutoSize -FreezeTopRow -AutoFilter
  Write-Host "Exported to $xlsx" -ForegroundColor Green
} else {
  $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
  Write-Host "Exported to $OutputPath" -ForegroundColor Green
}
