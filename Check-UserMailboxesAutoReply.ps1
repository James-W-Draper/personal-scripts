
# Connect to Exchange Online
# Import-Module ExchangeOnlineManagement
# Import-Module ImportExcel
Connect-ExchangeOnline

# Create an empty array to store the results
$results = @()

# Get all user mailboxes
$mailboxes = Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited

# Loop through each user mailbox to check for Automatic Replies (Out of Office)
foreach ($mailbox in $mailboxes) {
    # Use PrimarySmtpAddress instead of Alias for uniqueness
    $oofConfig = Get-MailboxAutoReplyConfiguration -Identity $mailbox.PrimarySmtpAddress
    
    # Create a hashtable for each mailbox to store relevant information, including DisplayName
    $result = [PSCustomObject]@{
        DisplayName      = $mailbox.DisplayName
        Mailbox          = $mailbox.PrimarySmtpAddress
        AutomaticReplies = if ($oofConfig.AutomaticRepliesEnabled -eq $true) { "Enabled" } else { "Disabled" }
     #   ExternalMessage  = $oofConfig.ExternalMessage
     #   InternalMessage  = $oofConfig.InternalMessage
        StartTime        = $oofConfig.StartTime
        EndTime          = $oofConfig.EndTime
    }

    # Add the result to the array
    $results += $result
}

# Export the results to an Excel file
$excelFilePath = "C:\scripts\AutoReplyStatus.xlsx"
$results | Export-Excel -Path $excelFilePath -AutoSize

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false

Write-Host "Results exported to Excel file successfully."