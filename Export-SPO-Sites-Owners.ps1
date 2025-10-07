<#
.SYNOPSIS
Export SharePoint Online sites with owners, usage analytics, content overlap, dependency map, and migration readiness.


.DESCRIPTION
- Enumerates all SPO site collections (excludes OneDrive unless -IncludeOneDrive).
- Group-connected sites: pulls M365 Group owners via Graph, falling back to UPN when Mail is empty.
- Classic sites: reads SharePoint "Owners" group via PnP, falling back to UPN/LoginName when Email is empty.
- Flags if site has a Microsoft Teams team.
- Adds usage analytics (last activity, is active).
- Compares document libraries/lists for overlap.
- Maps inter-site dependencies.
- Scores migration readiness based on customizations and modern compatibility.
- Caches PnP connections per site to avoid repeated interactive prompts.
- Outputs CSV by default; can write .xlsx with -Excel if ImportExcel is available.


.REQUIREMENTS
Modules:
  - Microsoft.Online.SharePoint.PowerShell
  - Microsoft.Graph
  - PnP.PowerShell
  - (Optional) ImportExcel

M365 Admin roles required:
  - SharePoint Admin

Site Collection Admin required:
  - This is a bit painful, because the only way to accomplish this is with Sharepoint admin, map yourself as a collection admin to every site (or everysite you need), run the script, and then remove yourself afterwards. This gets captured in audit!
  - The way to do this is to run this script:
    
    # Connect to SharePoint Online
    Connect-SPOService -Url https://<tenant>-admin.sharepoint.com
    # Get all sites
    $sites = Get-SPOSite -Limit All
    # Add yourself as Site Collection Admin to each
    foreach ($site in $sites) {
      Set-SPOUser -Site $site.Url -LoginName "<your-upn>@<yourdomain>.com" -IsSiteCollectionAdmin $true
    }

  - Then run the script where you need site collection
  - Then remove yourself as a site collection admin using this script:
    foreach ($site in $sites) {
      Set-SPOUser -Site $site.Url -LoginName "<your-upn>@<yourdomain>.com" -IsSiteCollectionAdmin $false
    }

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
  [string]$OutputPath = ".\SPO-Sites-Owners.csv",

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

# ----------------- New Feature Helpers -------------------

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

function Get-SiteListsSummary {
  param([string]$SiteUrl)
  # Returns array of @{Title=...; BaseTemplate=...; ItemCount=...}
  $summ = @()
  try {
    $conn = Get-PnPConnectionCached -SiteUrl $SiteUrl
    $lists = Get-PnPList -Connection $conn
    foreach ($l in $lists) {
      $summ += @{
        Title = $l.Title
        BaseTemplate = $l.BaseTemplate
        ItemCount = $l.ItemCount
      }
    }
  } catch {
    # ignore individual failures
  }
  return $summ
}

function Get-SiteDependencies {
  param([string]$SiteUrl,[string[]]$AllSiteUrls)
  $deps = @()
  try {
    $conn = Get-PnPConnectionCached -SiteUrl $SiteUrl
    # Scan modern site pages for links to other SPO sites
    $pages = @()
    try { 
      $pages = Get-PnPListItem -List "Site Pages" -Connection $conn -ErrorAction Stop
    } catch { }
    foreach ($p in $pages) {
      $content = $null
      if ($p.FieldValues.ContainsKey("CanvasContent1")) {
        $content = $p.FieldValues["CanvasContent1"]
      } elseif ($p.FieldValues.ContainsKey("WikiField")) {
        $content = $p.FieldValues["WikiField"]
      }
      if ($null -ne $content) {
        foreach ($siteUrl in $AllSiteUrls) {
          if ($siteUrl -ne $SiteUrl -and $content -match [regex]::Escape($siteUrl)) {
            $deps += $siteUrl
          }
        }
      }
    }
    # Check for lookup fields referencing lists in other sites (not fully robust)
    $lists = Get-PnPList -Connection $conn
    foreach ($list in $lists) {
      $fields = Get-PnPField -List $list.Id -Connection $conn
      foreach ($f in $fields) {
        if ($f.TypeAsString -eq "Lookup" -and $f.SchemaXml -match 'WebId="(.+?)"') {
          $refWebId = $matches[1]
          # Could map WebId to URL if needed
        }
      }
    }
  } catch { }
  return ($deps | Select-Object -Unique)
}

function Get-MigrationReadiness {
  param([string]$SiteUrl)
  $score = 100
  $notes = @()
  try {
    $conn = Get-PnPConnectionCached -SiteUrl $SiteUrl
    $web = Get-PnPWeb -Connection $conn -Includes MasterUrl, CustomMasterUrl, WebTemplate, EnableCustomQuickLaunch, EnableMinimalDownload
    if ($web.MasterUrl -notmatch "/_catalogs/masterpage/seattle.master") {
      $score -= 30
      $notes += "Custom master page"
    }
    if ($web.CustomMasterUrl -and $web.CustomMasterUrl -ne $web.MasterUrl) {
      $score -= 10
      $notes += "Custom custom master"
    }
    if ($web.WebTemplate -eq "STS" -and $web.EnableCustomQuickLaunch) {
      $score -= 10
      $notes += "Custom QuickLaunch"
    }
    if ($web.EnableMinimalDownload -eq $false) {
      $score -= 5
      $notes += "Minimal Download disabled"
    }
    # Custom scripts
    $siteProps = Get-PnPTenantSite -Url $SiteUrl -Connection $conn
    if ($siteProps.DenyAddAndCustomizePages -eq 0) {
      $score -= 20
      $notes += "Custom scripts enabled"
    }
    # Sandbox solutions
    $sandbox = Get-PnPSolution -Connection $conn -ErrorAction SilentlyContinue
    if ($sandbox) {
      $score -= 15
      $notes += "Sandbox solution(s) present"
    }
  } catch { $notes += "Could not evaluate all migration factors" }
  return @{Score=($score -lt 0 ? 0 : $score); Notes=($notes -join "; ")}
}

# ------------------- Bootstrapping ----------------------
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

$allSiteUrls = $allSites | ForEach-Object { $_.Url }

# ----------- Usage Analytics -----------
$siteUsage = Get-SiteUsageDetails

# ----------- Content Overlap: collect library/list names for all sites -----------
$siteListsMap = @{}
foreach ($site in $allSites) {
  $siteListsMap[$site.Url] = Get-SiteListsSummary -SiteUrl $site.Url
}
# Build overlap: for each list/library title, find other sites with same title
$listTitleToSites = @{}
foreach ($siteUrl in $siteListsMap.Keys) {
  foreach ($list in $siteListsMap[$siteUrl]) {
    $title = $list.Title
    if (-not $listTitleToSites.ContainsKey($title)) { $listTitleToSites[$title] = @() }
    $listTitleToSites[$title] += $siteUrl
  }
}
# Now for each site, build a list of overlapping libraries/lists
$siteContentOverlap = @{}
foreach ($siteUrl in $siteListsMap.Keys) {
  $overlap = @()
  foreach ($list in $siteListsMap[$siteUrl]) {
    $otherSites = $listTitleToSites[$list.Title] | Where-Object { $_ -ne $siteUrl }
    if ($otherSites.Count -gt 0) {
      $overlap += "$($list.Title):" + ($otherSites -join ",")
    }
  }
  $siteContentOverlap[$siteUrl] = $overlap -join " | "
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

  # Usage Analytics
  $lastActivity = $null
  $isActive = $null
  if ($siteUsage.ContainsKey($site.Url)) {
    $lastActivity = $siteUsage[$site.Url].LastActivityDate
    $isActive = $siteUsage[$site.Url].IsActive
  }

  # Content overlap
  $contentOverlap = $siteContentOverlap[$site.Url]

  # Dependency mapping
  $dependencies = (Get-SiteDependencies -SiteUrl $site.Url -AllSiteUrls $allSiteUrls) -join "; "

  # Migration readiness
  $migration = Get-MigrationReadiness -SiteUrl $site.Url

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
    ContentOverlap      = $contentOverlap
    SiteDependencies    = $dependencies
    MigrationReadinessScore = $migration.Score
    MigrationNotes      = $migration.Notes
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
