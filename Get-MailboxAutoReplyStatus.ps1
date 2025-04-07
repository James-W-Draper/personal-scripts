<#
.SYNOPSIS
Checks the Automatic Replies (Out of Office) status for all user mailboxes in Exchange Online.

.DESCRIPTION
This script connects to Exchange Online, retrieves all user mailboxes,
checks their Automatic Replies (OOF) status, and exports the result to an Excel file.

.NOTES
- Requires the ExchangeOnlineManagement and ImportExcel modules
- Must be run with appropriate permissions to access mailboxes
#>

# === MODULES & CONNECTION ===
# Import-Module ExchangeOnlineManagement
# Import-Module ImportExcel

Connect-ExchangeOnline

# === INITIALISE RESULTS ===
$results = @()
$mailboxes = Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited

# === LOOP THROUGH MAILBOXES ===
foreach ($mailbox in $mailboxes) {
    $oofConfig = Get-MailboxAutoReplyConfiguration -Identity $mailbox.PrimarySmtpAddress

    $results += [PSCustomObject]@{
        DisplayName      = $mailbox.DisplayName
        Mailbox          = $mailbox.PrimarySmtpAddress
        AutomaticReplies = if ($oofConfig.AutomaticRepliesEnabled) { "Enabled" } else { "Disabled" }
        StartTime        = $oofConfig.StartTime
        EndTime          = $oofConfig.EndTime
        # ExternalMessage = $oofConfig.ExternalMessage  # Uncomment if needed
        # InternalMessage = $oofConfig.InternalMessage  # Uncomment if needed
    }
}

# === EXPORT RESULTS ===
$excelFilePath = "C:\Scripts\AutoReplyStatus.xlsx"
$results | Export-Excel -Path $excelFilePath -AutoSize

# === CLEANUP ===
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "`n✅ Auto-reply status exported successfully to:`n$excelFilePath"
