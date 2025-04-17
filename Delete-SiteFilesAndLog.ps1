<#
.SYNOPSIS
Deletes specific files from SharePoint/OneDrive using Microsoft Graph, logs the results, and handles retention hold exclusions.

.DESCRIPTION
This script is for compliance use cases where files must be deleted from OneDrive or SharePoint
even when under a retention policy. It:
1. Accepts a CSV with full file URLs
2. Excludes affected sites from Purview retention policies
3. Waits for the exclusion to take effect
4. Deletes the listed files
5. Logs all results to a CSV
6. Re-applies retention to the excluded sites

.PARAMETER CsvPath
The path to a CSV with a "Combined" column containing full URLs of files to delete.

.PARAMETER LogPath
Optional. If not provided, it defaults to the same folder as the CSV, with the filename:
report_<CsvFileName>_<yyyymmdd>.csv

.PARAMETER AdminUrl
Included for backwards compatibility but not used in this version.

.EXAMPLE
.\Delete-SiteFilesAndLog.ps1 -CsvPath "C:\scripts\deleteme.csv" -AdminUrl "https://contoso-admin.sharepoint.com"
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$CsvPath,

    [Parameter()]
    [string]$LogPath,

    [Parameter(Mandatory = $true)]
    [string]$AdminUrl # No longer used
)

# Auto-generate log path if not provided
if (-not $LogPath) {
    $csvFileName = [System.IO.Path]::GetFileNameWithoutExtension($CsvPath)
    $csvDirectory = [System.IO.Path]::GetDirectoryName($CsvPath)
    $date = Get-Date -Format "yyyyMMdd"
    $LogPath = Join-Path $csvDirectory ("report_${csvFileName}_${date}.csv")
}

# Connect to Microsoft services
Connect-MgGraph -Scopes "Sites.ReadWrite.All"
Connect-IPPSSession

# Load CSV
$csv = Import-Csv -Path $CsvPath

# Auto-generate Combined column if missing
if (-not $csv[0].PSObject.Properties.Name -contains "Combined") {
    foreach ($entry in $csv) {
        $entry | Add-Member -MemberType NoteProperty -Name Combined -Value ("$($entry.'File Server')$($entry.Path)")
    }
}

$fileUrls = $csv.Combined

# Extract unique site URLs
$siteUrls = $fileUrls | ForEach-Object {
    try {
        $uri = [uri]$_
        if ($uri.AbsolutePath -match "^/(personal|sites|teams)/[^/]+") {
            "$($uri.Scheme)://$($uri.Host)$($Matches[0])"
        } else {
            "$($uri.Scheme)://$($uri.Host)"
        }
    } catch {
        Write-Warning "Invalid URL: $_"
        $null
    }
} | Where-Object { $_ } | Sort-Object -Unique

# Define Purview retention policy GUIDs
$retentionPolicies = @(
    "145b42b0-1cd4-44a2-8527-c0ea879ba9dc" # Replace this with your actual policy ID(s)
)

# STEP 1: Apply exclusions
foreach ($policyId in $retentionPolicies) {
    try {
        $policy = Get-RetentionCompliancePolicy -Identity $policyId
        $existingSites = $policy.ExcludedSharePointSiteUrls
        $combinedSites = ($existingSites + $siteUrls) | Sort-Object -Unique

        Write-Host "Excluding sites from policy '$($policy.Name)'..."
        Set-RetentionCompliancePolicy -Identity $policyId -ExcludedSharePointSiteUrls $combinedSites
    } catch {
        Write-Warning "Failed to update exclusions for policy ${policyId}: $($_.Exception.Message)"
    }
}

# STEP 2: Wait for changes to propagate
Write-Host "`nWaiting 30 minutes for exclusions to propagate..."
Start-Sleep -Seconds 1800

# STEP 3: Delete files and log results
$log = @()

foreach ($fileUrl in $fileUrls) {
    if (-not $fileUrl) { continue }

    Write-Host "`nDeleting: $fileUrl"

    try {
        $uri = [uri]$fileUrl
        $hostname = $uri.Host
        $sitePath = $uri.AbsolutePath -replace "^(/[^/]+/[^/]+)/.*", '$1'
        $siteUrl = "$($uri.Scheme)://$hostname$sitePath"

        $siteId = (Get-MgSite -Hostname $hostname -ErrorAction Stop).Id
        $drive = Get-MgSiteDrive -SiteId $siteId
        $driveId = $drive.Id

        $filePath = $uri.AbsolutePath -replace "^/$sitePath/", ""
        $driveItem = Get-MgDriveItemByPath -DriveId $driveId -Path "/$filePath" -ErrorAction Stop

        Remove-MgDriveItem -DriveId $driveId -ItemId $driveItem.Id -Confirm:$false

        $log += [pscustomobject]@{
            SiteURL   = $siteUrl
            FilePath  = $filePath
            FileURL   = $fileUrl
            Status    = "Deleted"
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
    } catch {
        $log += [pscustomobject]@{
            SiteURL   = $siteUrl
            FilePath  = "Unknown"
            FileURL   = $fileUrl
            Status    = "Failed - $($_.Exception.Message)"
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
    }
}

# STEP 4: Export deletion log
$Utf8WithBom = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllLines($LogPath, ($log | ConvertTo-Csv -NoTypeInformation), $Utf8WithBom)

Write-Host "`nDeletion complete. Log saved to: $LogPath"

# STEP 5: Re-apply retention
foreach ($policyId in $retentionPolicies) {
    try {
        $policy = Get-RetentionCompliancePolicy -Identity $policyId
        $existingSites = $policy.ExcludedSharePointSiteUrls
        $updatedSites = $existingSites | Where-Object { $siteUrls -notcontains $_ }

        Write
