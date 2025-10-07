<#
.SYNOPSIS
Export SharePoint Online sites with owners, Teams linkage, and usage analytics (non-intrusive: no site-level admin required).

.DESCRIPTION
- Enumerates all SPO site collections (excludes OneDrive unless -IncludeOneDrive).
- Group-connected sites: pulls M365 Group owners via Graph, falling back to UPN when Mail is empty.
- Flags if site has a Microsoft Teams team.
- Outputs last activity date via Graph (Reports.Read.All required).
- No PnP.PowerShell usage, no site-level customization or list enumeration.
- Outputs CSV by default; can write .xlsx with -Excel if ImportExcel is available.

.REQUIREMENTS
Modules:
  - Microsoft.Online.SharePoint.PowerShell
  - Microsoft.Graph
  - (Optional) ImportExcel

Graph delegated scopes required:
  - Group.Read.All
  - Directory.Read.All
  - Team.ReadBasic.All
  - Reports.Read.All (for usage)
#>

param(
  [Parameter(Mandatory = $true)]
  [string]$Tenant,     

  [Parameter(Mandatory = $false)]
  [string]$AdminUrl,   

  [Parameter(Mandatory = $false)]
  [string]$OutputPath = ".\SPO-Sites-Owners-lite.csv",

  [switch]$IncludeOneDrive,

  [switch]$Excel
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
    [string]$UserPrincipalName
  )
  $addr = $null
  if ($Mail) { $addr = $Mail }
  elseif ($UserPrincipalName) { $addr = $UserPrincipalName }
  if ($DisplayName) {
    if ($addr) { return ('{0} <{1}>' -f $DisplayName, $addr) }
    else { return $DisplayName }
  } else {
    return ($addr ?? "")
  }
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

function Get-SiteUsageDetails {
  # Returns a hashtable: URL => @{LastActivityDate=..., IsActive=...}
  Write-Host "Retrieving site usage analytics from Graph..." -ForegroundColor Cyan
  $usage = @{}
  try {
    $bytes = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D30')" -OutputType Byte
    $csv = [System.Text.Encoding]::UTF8.GetString($bytes)
    $lines = $csv -split "`n"
    $header = $lines[0].Trim().Split(",")
    $urlIdx = $header.IndexOf("Site URL")
    $lastActivityIdx = $header.IndexOf("Last Activity Date")
    for ($i = 1; $i -lt $lines.Length; $i++) {
      $row = $lines[$i].Trim().Split(",")
      if ($row.Count -lt [Math]::Max($urlIdx,$lastActivityIdx)+1) { continue }
      $url = $row[$urlIdx]
      $lastDate = $row[$lastActivityIdx]
      $parsedDate = $null
      if ($lastDate -and [datetime]::TryParse($lastDate, [ref]$parsedDate)) {
        $isActive = ($parsedDate -gt (Get-Date).AddDays(-90))
        $usage[$url] = @{LastActivityDate=$parsedDate; IsActive=$isActive}
      }
    }
  } catch {
    Write-Warning "Failed to fetch usage analytics: $($_.Exception.Message)"
  }
  return $usage
}

# ------------------- Bootstrapping ----------------------
Ensure-Module -Name Microsoft.Online.SharePoint.PowerShell
Ensure-Module -Name Microsoft.Graph
if ($Excel) { Ensure-Module -Name ImportExcel }

# Sanitize Tenant and compute AdminUrl if missing
$Tenant = ($Tenant -replace '\s', '')
if (-not $Tenant) { throw "Tenant cannot be empty after sanitisation." }
if (-not $AdminUrl) { $AdminUrl = "https://$Tenant-admin.sharepoint.com" }

Write-Host "Connecting to SharePoint Admin: $AdminUrl" -ForegroundColor Cyan
Connect-SPOService -Url $AdminUrl

Write-Host "Connecting to Microsoft Graph (interactive)..." -ForegroundColor Cyan
$scopes = @("Group.Read.All","Directory.Read.All","Team.ReadBasic.All","Reports.Read.All")
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

# ----------- Usage Analytics -----------
$siteUsage = Get-SiteUsageDetails

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
  }

  # Usage Analytics
  $lastActivity = $null
  $isActive = $null
  if ($siteUsage.ContainsKey($site.Url)) {
    $lastActivity = $siteUsage[$site.Url].LastActivityDate
    $isActive = $siteUsage[$site.Url].IsActive
  }

  $results.Add([pscustomobject]@{
    Title               = $site.Title
    Url                 = $site.Url
    Template            = $site.Template
    StorageQuotaMB      = $site.StorageQuota
    StorageUsageMB      = $site.StorageUsageCurrent
    PrimaryAdmin        = $site.Owner
    SecondaryAdmins     = ($site.AdditionalAdministrators -join "; ")
    IsGroupConnected    = $isGroupConnected
    GroupId             = if ($groupId) { $groupId } else { $null }
    HasTeamsTeam        = $hasTeam
    Owners              = $ownersText
    LastActivityDate    = $lastActivity
    IsActiveSite        = $isActive
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
