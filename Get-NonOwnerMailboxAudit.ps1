# Connect to Exchange Online
Connect-ExchangeOnline -ShowBanner:$false

# Set the default date range to the last 90 days
$StartDate = (Get-Date).AddDays(-90).Date
$EndDate = Get-Date

# Define the output CSV file with a timestamp in ddMMyyyy format
$OutputCSV = ".\NonOwnerMailboxAudit_$((Get-Date -Format ddMMyyyy_HHmm)).csv"

# Define the audit operations to search for
$Operations = 'ApplyRecord','Copy','Create','FolderBind','HardDelete','MessageBind','Move','MoveToDeletedItems','RecordDelete','SendAs','SendOnBehalf','SoftDelete','Update','UpdateCalendarDelegation','UpdateFolderPermissions','UpdateInboxRules'

# Calculate total 12-hour periods for progress tracking
$TotalChunks = (($EndDate - $StartDate).Days + 1) * 2
$CurrentChunk = 1

# Process audit log in 12-hour chunks to handle result size limitations
for ($Date = $StartDate; $Date -le $EndDate; $Date = $Date.AddDays(1)) {
    foreach ($Half in @(0,12)) {
        $ChunkStart = $Date.AddHours($Half)
        $ChunkEnd = $ChunkStart.AddHours(12).AddSeconds(-1)
        $SessionId = "NonOwnerMailboxAudit_$($ChunkStart.ToString('ddMMyyyy_HHmm'))"
        
        Write-Progress -Activity "Processing Audit Logs" -Status "Processing period: $ChunkStart to $ChunkEnd" -PercentComplete (($CurrentChunk / $TotalChunks) * 100)

        $Results = Search-UnifiedAuditLog -StartDate $ChunkStart -EndDate $ChunkEnd -Operations $Operations -SessionId $SessionId -SessionCommand ReturnLargeSet -ResultSize 5000

        foreach ($Result in $Results) {
            $Audit = $Result.AuditData | ConvertFrom-Json

            if ($Audit.LogonType -eq 0 -or $Audit.ExternalAccess) { continue }

            if ($Audit.LogonUserSId -ne $Audit.MailboxOwnerSid -or (($Audit.Operation -eq "SendAs" -or $Audit.Operation -eq "SendOnBehalf") -and $Audit.UserType -eq 0)) {
                $AccessedMB = if ($Audit.Operation -eq "SendAs") { $Audit.SendAsUserSMTP } elseif ($Audit.Operation -eq "SendOnBehalf") { $Audit.SendOnBehalfOfUserSmtp } else { $Audit.MailboxOwnerUPN }
                $AccessedBy = $Audit.UserId

                if ($AccessedMB -eq $AccessedBy) { continue }

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
        
        $CurrentChunk++
    }
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
