<#
.SYNOPSIS
Audits non-owner mailbox access in Exchange Online over the past 90 days.

.DESCRIPTION
This script uses Search-UnifiedAuditLog to identify non-owner mailbox access,
such as delegated access or send-as/send-on-behalf actions, within a rolling
90-day window. Results are chunked in 12-hour blocks to ensure completeness.

.PARAMETER ShowProgress
Switch to show or hide the progress bar (default: shown).

.EXAMPLE
.\Get-NonOwnerMailboxAuditReport.ps1
# Scans for non-owner access in the last 90 days with progress output.

.EXAMPLE
.\Get-NonOwnerMailboxAuditReport.ps1 -ShowProgress:$false
# Runs the same audit silently without progress display.
#>

param (
    [bool]$ShowProgress = $true
)

# === CONNECT TO EXCHANGE ONLINE ===
Connect-ExchangeOnline -ShowBanner:$false

# === CONFIGURATION ===
$StartDate = (Get-Date).AddDays(-90).Date
$EndDate = Get-Date
$OutputCSV = ".\NonOwnerMailboxAudit_$((Get-Date -Format 'ddMMyyyy_HHmm')).csv"

$Operations = @(
    'ApplyRecord','Copy','Create','FolderBind','HardDelete','MessageBind','Move',
    'MoveToDeletedItems','RecordDelete','SendAs','SendOnBehalf','SoftDelete',
    'Update','UpdateCalendarDelegation','UpdateFolderPermissions','UpdateInboxRules'
)

$TotalChunks = (($EndDate - $StartDate).Days + 1) * 2
$CurrentChunk = 1

# Ensure headers are added to CSV before loop starts
if (-not (Test-Path $OutputCSV)) {
    "" | Select-Object 'Access Time','Accessed by','Performed Operation','Accessed Mailbox','Logon Type','Result Status','External Access','More Info' |
        Export-Csv -Path $OutputCSV -NoTypeInformation
}

# === MAIN AUDIT LOOP ===
for ($Date = $StartDate; $Date -le $EndDate; $Date = $Date.AddDays(1)) {
    foreach ($Half in @(0,12)) {
        $ChunkStart = $Date.AddHours($Half)
        $ChunkEnd = $ChunkStart.AddHours(12).AddSeconds(-1)
        $SessionId = "NonOwnerMailboxAudit_$($ChunkStart.ToString('ddMMyyyy_HHmm'))"

        if ($ShowProgress) {
            Write-Progress -Activity "Processing Audit Logs" `
                           -Status "Processing period: $ChunkStart to $ChunkEnd" `
                           -PercentComplete (($CurrentChunk / $TotalChunks) * 100)
        }

        try {
            $Results = Search-UnifiedAuditLog -StartDate $ChunkStart -EndDate $ChunkEnd `
                        -Operations $Operations -SessionId $SessionId `
                        -SessionCommand ReturnLargeSet -ResultSize 5000
        } catch {
            Write-Warning "Failed to retrieve audit logs for $ChunkStart to $ChunkEnd: $_"
            $CurrentChunk++
            continue
        }

        foreach ($Result in $Results) {
            try {
                $Audit = $Result.AuditData | ConvertFrom-Json
            } catch {
                Write-Warning "Failed to parse AuditData JSON: $_"
                continue
            }

            if ($Audit.LogonType -eq 0 -or $Audit.ExternalAccess) { continue }

            $NonOwnerAccess = ($Audit.LogonUserSId -ne $Audit.MailboxOwnerSid) -or
                              (($Audit.Operation -in @("SendAs", "SendOnBehalf")) -and $Audit.UserType -eq 0)

            if ($NonOwnerAccess) {
                $AccessedMB = switch ($Audit.Operation) {
                    "SendAs"         { $Audit.SendAsUserSMTP }
                    "SendOnBehalf"   { $Audit.SendOnBehalfOfUserSmtp }
                    default          { $Audit.MailboxOwnerUPN }
                }

                $AccessedBy = $Audit.UserId
                if ($AccessedMB -eq $AccessedBy) { continue }

                [PSCustomObject]@{
                    'Access Time'      = (Get-Date $Audit.CreationTime).ToLocalTime()
                    'Accessed by'      = $AccessedBy
                    'Performed Operation' = $Audit.Operation
                    'Accessed Mailbox' = $AccessedMB
                    'Logon Type'       = switch ($Audit.LogonType) {
                                            1 { "Administrator" }
                                            2 { "Delegated" }
                                            default { "Microsoft Datacenter" }
                                         }
                    'Result Status'    = $Audit.ResultStatus
                    'External Access'  = $Audit.ExternalAccess
                    'More Info'        = $Result.AuditData
                } | Export-Csv -Path $OutputCSV -NoTypeInformation -Append
            }
        }

        $CurrentChunk++
    }
}

# === DISCONNECT ===
Disconnect-ExchangeOnline -Confirm:$false

Write-Host "`n✅ Audit complete. Results saved to: $OutputCSV`n"
