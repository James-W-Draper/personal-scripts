# Exchange Online PowerShell Script
# Retrieves all folders (including archive) from a specified mailbox,
# listing folder paths, item counts, unread counts, folder types, and sizes.
# Exports results to a CSV file.

# -----------------------------
# Connect to Exchange Online
# -----------------------------
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline -UserPrincipalName admin@yourdomain.com

# -----------------------------
# Define the target mailbox
# -----------------------------
$Mailbox = "user@yourdomain.com"
Write-Host "Retrieving mailbox folder statistics for: $Mailbox" -ForegroundColor Cyan

# -----------------------------
# Retrieve primary mailbox folders
# -----------------------------
Write-Host "Fetching folder details for primary mailbox..." -ForegroundColor Yellow
$primaryFolders = Get-MailboxFolderStatistics -Identity $Mailbox | 
    Select-Object @{Name="Mailbox"; Expression={$Mailbox}},   # Display the mailbox
                  @{Name="MailboxType"; Expression={"Primary"}}, # Indicate it's a primary mailbox
                  FolderPath, Name, FolderType, # Folder details
                  ItemsInFolder, ItemsInFolderAndSubfolders,  # Item counts
                  UnreadCount, FolderSize, # Unread and folder size
                  @{Name="FolderSizeMB"; Expression={[math]::Round($_.FolderSize.Value.ToMB(), 2)}} # Convert size to MB

Write-Host "Primary mailbox folders retrieved: $($primaryFolders.Count)" -ForegroundColor Green

# -----------------------------
# Retrieve archive mailbox folders
# -----------------------------
Write-Host "Fetching folder details for archive mailbox (if available)..." -ForegroundColor Yellow
$archiveFolders = Get-MailboxFolderStatistics -Identity $Mailbox -Archive | 
    Select-Object @{Name="Mailbox"; Expression={$Mailbox}},   # Display the mailbox
                  @{Name="MailboxType"; Expression={"Archive"}}, # Indicate it's an archive mailbox
                  FolderPath, Name, FolderType, # Folder details
                  ItemsInFolder, ItemsInFolderAndSubfolders,  # Item counts
                  UnreadCount, FolderSize, # Unread and folder size
                  @{Name="FolderSizeMB"; Expression={[math]::Round($_.FolderSize.Value.ToMB(), 2)}} # Convert size to MB

Write-Host "Archive mailbox folders retrieved: $($archiveFolders.Count)" -ForegroundColor Green

# -----------------------------
# Combine Results
# -----------------------------
$allFolders = $primaryFolders + $archiveFolders

# -----------------------------
# Output Results in Table Format
# -----------------------------
Write-Host "Displaying folder details..." -ForegroundColor Yellow
$allFolders | Format-Table Mailbox, MailboxType, FolderPath, Name, FolderType, ItemsInFolder, UnreadCount, FolderSizeMB -AutoSize

# -----------------------------
# Export Results to CSV
# -----------------------------
$csvPath = "C:\MailboxFolders_WithPrimary_Archive_Verbose.csv"
$allFolders | Export-Csv -Path $csvPath -NoTypeInformation
Write-Host "Results exported to: $csvPath" -ForegroundColor Magenta

# -----------------------------
# Disconnect from Exchange Online
# -----------------------------
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Disconnected from Exchange Online." -ForegroundColor Red
