<#
.SYNOPSIS
Gets folder identifiers for Exchange mailboxes or SharePoint/OneDrive sites for use in Microsoft Purview Content Searches.

.PARAMETER Target
An email address (Exchange) or a SharePoint/OneDrive URL.

.EXAMPLE
.\Get-ContentSearchFolderTargets.ps1 -Target "user@contoso.com"
.\Get-ContentSearchFolderTargets.ps1 -Target "https://contoso.sharepoint.com/sites/mysite"
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$Target
)

$today = Get-Date -Format "yyyyMMdd"
$outputDir = "C:\Scripts"
if (-not (Test-Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
}

# -------------------------
# Exchange mailbox
# -------------------------
if ($Target -match "@") {
    $emailAddress = $Target
    $safeName = $emailAddress -replace "[^a-zA-Z0-9]", "_"
    $csvPath = Join-Path $outputDir "ExchangeFolders_${safeName}_${today}.csv"

    Write-Host "`n🔐 Connecting to Exchange Online..." -ForegroundColor Cyan
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -ShowBanner:$false -CommandName Get-MailboxFolderStatistics

    $folderQueries = @()
    $folderStatistics = Get-MailboxFolderStatistics -Identity $emailAddress -ResultSize Unlimited

    foreach ($folder in $folderStatistics) {
        $folderId = $folder.FolderId
        $folderPath = $folder.FolderPath

        $encoding = [System.Text.Encoding]::GetEncoding("us-ascii")
        $nibbler = $encoding.GetBytes("0123456789ABCDEF")
        $folderIdBytes = [Convert]::FromBase64String($folderId)
        $indexIdBytes = New-Object byte[] 48
        $indexIdIdx = 0

        $folderIdBytes | Select-Object -Skip 23 -First 24 | ForEach-Object {
            $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -shr 4]
            $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -band 0xF]
        }

        $folderQuery = "folderid:$($encoding.GetString($indexIdBytes))"
        $folderQueries += [PSCustomObject]@{
            FolderPath  = $folderPath
            FolderQuery = $folderQuery
        }
    }

    Write-Host "`n📂 Exporting Exchange folder list to:`n$csvPath" -ForegroundColor Green
    $folderQueries | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
}

# -------------------------
# SharePoint / OneDrive
# -------------------------
elseif ($Target -match "^http") {
    $siteUrl = $Target
    $safeSite = ($siteUrl -split "/")[-1] -replace "[^a-zA-Z0-9]", "_"
    $csvPath = Join-Path $outputDir "SharePointFolders_${safeSite}_${today}.csv"

    $searchName = "SPFoldersSearch"
    $searchActionName = "${searchName}_Preview"
    $documentLinks = @()

    Write-Host "`n🔐 Connecting to Compliance Center..." -ForegroundColor Cyan
    Import-Module ExchangeOnlineManagement
    Connect-IPPSSession

    Remove-ComplianceSearch -Identity $searchName -Confirm:$false -ErrorAction SilentlyContinue

    Write-Host "📡 Starting folder discovery search for site: $siteUrl"
    $search = New-ComplianceSearch -Name $searchName -ContentMatchQuery 'contenttype:folder OR contentclass:STS_Web' -SharePointLocation $siteUrl
    Start-ComplianceSearch -Identity $searchName

    do {
        Start-Sleep -Seconds 5
        $search = Get-ComplianceSearch -Identity $searchName
        Write-Host "⏳ Waiting for Compliance Search to complete..."
    } while ($search.Status -ne 'Completed')

    if ($search.Items -gt 0) {
        $searchAction = New-ComplianceSearchAction -SearchName $searchName -Preview

        do {
            Start-Sleep -Seconds 5
            $searchAction = Get-ComplianceSearchAction -Identity $searchActionName
            Write-Host "⏳ Waiting for Preview Action to complete..."
        } while ($searchAction.Status -ne 'Completed')

        $results = $searchAction.Results
        $matches = Select-String -InputObject $results -Pattern "Data Link:.+[,}]" -AllMatches

        foreach ($match in $matches.Matches) {
            $url = $match.Value -replace "Data Link: " -replace "," -replace "}"
            $documentLinks += [PSCustomObject]@{
                DocumentLink = "documentlink:`"$url`""
            }
        }

        Write-Host "`n📁 Exporting SharePoint/OneDrive folder list to:`n$csvPath" -ForegroundColor Green
        $documentLinks | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    } else {
        Write-Warning "No folders found for: $siteUrl"
    }

    Remove-ComplianceSearch -Identity $searchName -Confirm:$false -ErrorAction SilentlyContinue
}

# -------------------------
# Invalid input
# -------------------------
else {
    Write-Error "❌ '$Target' is not a valid email address or site URL."
}
