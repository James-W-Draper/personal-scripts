# Connect to Exchange Online
Connect-ExchangeOnline -ShowBanner:$false

# Set the default date range to the last 90 days
$StartDate = (Get-Date).AddDays(-90).Date
$EndDate = Get-Date

# Define the output CSV file with a timestamp in ddMMyyyy format
$OutputCSV = ".\NonOwnerMailboxAudit_$((Get-Date -Format ddMMyyyy_HHmm)).csv"

# Define the audit operations to search for
$Operations = 'ApplyRecord','Copy','Create','FolderBind','HardDelete','MessageBind','Move','MoveToDeletedItems','RecordDelete','SendAs','SendOnBehalf','SoftDelete','Update','UpdateCalendarDelegation','UpdateFolderPermissions','UpdateInboxRules'

# Generate a unique session ID based on the current date and purpose
$SessionId = "NonOwnerMailboxAudit_$((Get-Date -Format ddMMyyyy_HHmm))"

# Retrieve audit log entries for the last 90 days with maximum result size
$Results = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations $Operations -SessionId $SessionId -SessionCommand ReturnLargeSet -ResultSize 50000

foreach ($Result in $Results) {
    # Convert the audit data from JSON format
    $Audit = $Result.AuditData | ConvertFrom-Json

    # Skip owner and external access records
    if ($Audit.LogonType -eq 0 -or $Audit.ExternalAccess) { continue }

    # Identify non-owner access records, including SendAs and SendOnBehalf operations
    if ($Audit.LogonUserSId -ne $Audit.MailboxOwnerSid -or (($Audit.Operation -eq "SendAs" -or $Audit.Operation -eq "SendOnBehalf") -and $Audit.UserType -eq 0)) {
        # Determine the accessed mailbox based on the operation
        $AccessedMB = if ($Audit.Operation -eq "SendAs") { $Audit.SendAsUserSMTP } elseif ($Audit.Operation -eq "SendOnBehalf") { $Audit.SendOnBehalfOfUserSmtp } else { $Audit.MailboxOwnerUPN }
        $AccessedBy = $Audit.UserId

        # Skip if the accessed mailbox is the same as the user
        if ($AccessedMB -eq $AccessedBy) { continue }

        # Export the audit data to CSV
        [PSCustomObject]@{
            'Access Time' = (Get-Date($Audit.CreationTime)).ToLocalTime()
            'Accessed by' = $AccessedBy
            'Performed Operation' = $Audit.Operation
            'Accessed Mailbox' = $AccessedMB
            'Logon Type' = switch ($Audit.LogonType) { 1 {"Administrator"} 2 {"Delegated"} default {"Microsoft Datacenter"} }
            'Result Status' = $Audit.ResultStatus
            'External Access' = $Audit.ExternalAccess
            'More Info' = $Result.AuditData
        } | Export-Csv $OutputCSV -NoTypeInformation -Append
    }
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
