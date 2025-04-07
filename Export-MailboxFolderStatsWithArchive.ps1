<#
.SYNOPSIS
Retrieves folder statistics (including archive) for a specified mailbox in Exchange Online.

.DESCRIPTION
This script connects to Exchange Online, retrieves all folders (primary and archive) for a mailbox,
and exports folder paths, item counts, unread counts, folder types, and folder sizes (in MB) to a CSV file.

.NOTES
- Requires ExchangeOnlineManagement module
- Run as a user with the appropriate permissions (e.g. View-Only Recipients)
#>

# === CONFIGURATION ===
$Mailbox = "user@domain.com"  # <-- Replace with the target mailbox
$csvPath = "C:\Scripts\MailboxFolders_WithPrimary_Archive.csv"

# === CONNECT TO EXCHANGE ONLINE ===
Write-Host "`nConnecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline

Write-Host "`nRetrieving mailbox folder statistics for: $Mailbox" -ForegroundColor Cyan

# === PRIMARY MAILBOX FOLDERS ===
Write-Host "Fetching primary mailbox folders..." -ForegroundColor Yellow
$primaryFolders = Get-MailboxFolderStatistics -Identity $Mailbox | 
    Select-Object @{Name="Mailbox"; Expression={$Mailbox}},
                  @{Name="MailboxType"; Expression={"Primary"}},
                  FolderPath, Name, FolderType,
                  ItemsInFolder, ItemsInFolderAndSubfolders,
                  UnreadCount, FolderSize,
                  @{Name="FolderSizeMB"; Expression={[math]::Round($_.FolderSize.Value.ToMB(), 2)}}

Write-Host "✔ Retrieved $($primaryFolders.Count) primary folders." -ForegroundColor Green

# === ARCHIVE MAILBOX FOLDERS ===
Write-Host "Checking for archive mailbox..." -ForegroundColor Yellow
try {
    $archiveFolders = Get-MailboxFolderStatistics -Identity $Mailbox -Archive -ErrorAction Stop | 
        Select-Object @{Name="Mailbox"; Expression={$Mailbox}},
                      @{Name="MailboxType"; Expression={"Archive"}},
                      FolderPath, Name, FolderType,
                      ItemsInFolder, ItemsInFolderAndSubfolders,
                      UnreadCount, FolderSize,
                      @{Name="FolderSizeMB"; Expression={[math]::Round($_.FolderSize.Value.ToMB(), 2)}}

    Write-Host "✔ Retrieved $($archiveFolders.Count) archive folders." -ForegroundColor Green
} catch {
    Write-Warning "No archive mailbox found or access denied. Continuing without archive data."
    $archiveFolders = @()
}

# === COMBINE RESULTS ===
$allFolders = $primaryFolders + $archiveFolders

# === DISPLAY RESULTS ===
Write-Host "`nDisplaying folder details..." -ForegroundColor Yellow
$allFolders | Format-Table Mailbox, MailboxType, FolderPath, Name, FolderType, ItemsInFolder, UnreadCount, FolderSizeMB -AutoSize

# === EXPORT TO CSV ===
$allFolders | Export-Csv -Path $csvPath -NoTypeInformation
Write-Host "`n📄 Results exported to: $csvPath" -ForegroundColor Magenta

# === DISCONNECT ===
# Uncomment the line below to disconnect at the end of the script
# Disconnect-ExchangeOnline -Confirm:$false
# Write-Host "`nDisconnected from Exchange Online." -ForegroundColor Red
