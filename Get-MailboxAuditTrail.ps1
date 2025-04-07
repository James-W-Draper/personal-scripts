<#
.SYNOPSIS
Exports raw and parsed mailbox audit logs for a specific user over a defined date range.

.DESCRIPTION
This script queries the Unified Audit Log in 6-hour chunks, retrieving all activity associated with the specified mailbox.
It exports both raw audit data and a filtered, human-readable summary where `AffectedItems` are present.

.PARAMETER Mailbox
The UPN (UserPrincipalName) of the mailbox to audit.

.PARAMETER StartDate
Start of the audit range (format: yyyy-MM-dd or full datetime).

.PARAMETER EndDate
End of the audit range (format: yyyy-MM-dd or full datetime).

.EXAMPLE
.\Export-MailboxAuditReport.ps1 -Mailbox "user@domain.com"

Audits the specified user's mailbox from 90 days ago to today.

.EXAMPLE
.\Export-MailboxAuditReport.ps1 -Mailbox "user@domain.com" -StartDate "2025-04-01" -EndDate "2025-04-07"

Exports audit log data for a specific mailbox between 1st and 7th April 2025.

.NOTES
- Requires the ExchangeOnlineManagement module
- Run with appropriate permissions to access Unified Audit Logs
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$Mailbox,

    [Parameter()]
    [datetime]$StartDate = (Get-Date).AddDays(-90),

    [Parameter()]
    [datetime]$EndDate = (Get-Date)
)

# Output paths (timestamped)
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$rawExportPath    = "C:\Scripts\AuditLogs_Raw_$($Mailbox)_$timestamp.csv"
$parsedExportPath = "C:\Scripts\AuditLogs_Parsed_$($Mailbox)_$timestamp.csv"

# Init
$allRawLogs = @()
$allParsedResults = @()

Write-Host "`n📋 Auditing mailbox: $Mailbox"
Write-Host "🗓️  Date Range: $StartDate → $EndDate"
Write-Host "⏱️  Querying logs in 6-hour chunks...`n"

# Loop in 6-hour windows
while ($StartDate -lt $EndDate) {
    $chunkEnd = $StartDate.AddHours(6)
    Write-Host "🔍 Querying from $StartDate to $chunkEnd..."

    try {
        $logs = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $chunkEnd -ResultSize 5000
    } catch {
        Write-Warning "❌ Failed to query logs from $StartDate to $chunkEnd: $($_.Exception.Message)"
        $StartDate = $chunkEnd
        continue
    }

    $allRawLogs += $logs

    $parsed = $logs | ForEach-Object {
        try {
            $data = $_.AuditData | ConvertFrom-Json
        } catch {
            return
        }

        if ($data.MailboxOwnerUPN -eq $Mailbox -and $data.AffectedItems) {
            foreach ($item in $data.AffectedItems) {
                [PSCustomObject]@{
                    Operation         = $_.Operations
                    Actor             = $_.UserIds
                    MailboxOwner      = $data.MailboxOwnerUPN
                    TimeStamp         = $_.CreationDate
                    Subject           = $item.Subject
                    Sender            = $item.Sender
                    DestinationFolder = $data.DestFolder?.Path
                    SourceFolder      = $data.ParentFolder?.Path
                    ClientIP          = $data.ClientIP
                    ClientInfo        = $data.ClientInfoString
                    UserAgent         = $data.UserAgent
                    Attachments       = $item.Attachments -join '; '
                }
            }
        }
    }

    $allParsedResults += $parsed
    $StartDate = $chunkEnd
}

# Export both files
$allRawLogs | Export-Csv -Path $rawExportPath -NoTypeInformation -Encoding UTF8
$allParsedResults | Export-Csv -Path $parsedExportPath -NoTypeInformation -Encoding UTF8

# Report
Write-Host "`n✅ Export complete!"
Write-Host "📂 Raw logs:     $rawExportPath"
Write-Host "📂 Parsed logs:  $parsedExportPath`n"
