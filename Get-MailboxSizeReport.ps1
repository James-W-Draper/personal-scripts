<#
.SYNOPSIS
    Generates a detailed mailbox size report for all Exchange Online mailboxes with quota thresholds, archive status, and color-coded Excel output.

.DESCRIPTION
    This script connects to Exchange Online, retrieves size and quota data for each mailbox, and generates a comprehensive Excel report. It includes:
        - Total mailbox size (MB)
        - Item count
        - Archive mailbox status
        - Prohibit send, send/receive, and warning quota thresholds
        - Color-coded Excel output based on threshold proximity

.PARAMETER OutputPath
    Optional. Directory to save the report. Defaults to "C:\Scripts".

.EXAMPLE
    .\Get-MailboxSizeReport.ps1

.EXAMPLE
    .\Get-MailboxSizeReport.ps1 -OutputPath "D:\Reports"

.NOTES
    Requires:
    - ExchangeOnlineManagement
    - ImportExcel (for enhanced formatting)
#>

param (
    [string]$OutputPath = "C:\Scripts"
)

# Ensure modules
foreach ($module in @("ExchangeOnlineManagement", "ImportExcel")) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Install-Module -Name $module -Scope CurrentUser -Force
    }
}

Import-Module ExchangeOnlineManagement
Import-Module ImportExcel

# Connect to Exchange
Connect-ExchangeOnline -ShowBanner:$false

# Get all mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox | Sort-Object DisplayName
$mailboxStats = Get-MailboxStatistics -ResultSize Unlimited | Group-Object DisplayName -AsHashTable -AsString

# Collect data
$report = foreach ($mbx in $mailboxes) {
    $stats = $mailboxStats[$mbx.DisplayName]
    [PSCustomObject]@{
        DisplayName                = $mbx.DisplayName
        UserPrincipalName         = $mbx.UserPrincipalName
        ArchiveEnabled            = $mbx.ArchiveStatus -eq 'Active'
        MailboxSizeMB             = [math]::Round($stats.TotalItemSize.Value.ToMB(), 2)
        ItemCount                 = $stats.ItemCount
        IssueWarningQuota         = $mbx.IssueWarningQuota.Value.ToMB()
        ProhibitSendQuota         = $mbx.ProhibitSendQuota.Value.ToMB()
        ProhibitSendReceiveQuota  = $mbx.ProhibitSendReceiveQuota.Value.ToMB()
    }
}

# Output path
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFile = Join-Path $OutputPath "MailboxSizeReport_$timestamp.xlsx"

# Export to Excel with conditional formatting
$report | Export-Excel -Path $outputFile -WorksheetName "MailboxSizes" -AutoSize -FreezeTopRow -BoldTopRow -ConditionalFormat @( 
    New-ConditionalText -TextCondition ">=90" -Range "D2:D9999" -BackgroundColor 'LightSalmon'
    New-ConditionalText -TextCondition ">=80" -Range "D2:D9999" -BackgroundColor 'LightYellow'
) -TableName MailboxSizeReport

Write-Host "Report saved to: $outputFile" -ForegroundColor Green

Disconnect-ExchangeOnline -Confirm:$false
