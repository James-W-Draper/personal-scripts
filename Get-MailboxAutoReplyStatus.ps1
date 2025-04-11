<#
.SYNOPSIS
    Retrieves the auto-reply (Out of Office) status for all user mailboxes in Exchange Online.

.DESCRIPTION
    This script connects to Exchange Online and retrieves auto-reply configurations for all user mailboxes.
    It reports whether automatic replies are enabled and flags those that are still active past their scheduled
    end date. The results are exported to an Excel report.

.PARAMETER OutputPath
    Optional. The directory where the Excel report will be saved. Defaults to "C:\Scripts".

.EXAMPLE
    .\Get-MailboxAutoReplyStatus.ps1

.EXAMPLE
    .\Get-MailboxAutoReplyStatus.ps1 -OutputPath "D:\Reports"

.NOTES
    Requires:
    - ExchangeOnlineManagement
    - ImportExcel
#>

[CmdletBinding()]
param (
    [string]$OutputPath = "C:\Scripts"
)

# Ensure required modules
foreach ($module in @("ExchangeOnlineManagement", "ImportExcel")) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Install-Module -Name $module -Scope CurrentUser -Force
    }
}

Import-Module ExchangeOnlineManagement
Import-Module ImportExcel

Connect-ExchangeOnline -ShowBanner:$false

# Retrieve all user mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox
$report = @()

foreach ($mailbox in $mailboxes) {
    $config = Get-MailboxAutoReplyConfiguration -Identity $mailbox.UserPrincipalName
    $isExpired = $false

    if ($config.AutomaticRepliesEnabled -and $config.EndTime -ne $null) {
        if ((Get-Date) -gt $config.EndTime) {
            $isExpired = $true
        }
    }

    $report += [pscustomobject]@{
        DisplayName       = $mailbox.DisplayName
        UserPrincipalName = $mailbox.UserPrincipalName
        AutoReplyEnabled  = $config.AutomaticRepliesEnabled
        StartTime         = $config.StartTime
        EndTime           = $config.EndTime
        IsExpired         = $isExpired
    }
}

# Export results to Excel
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFile = Join-Path $OutputPath "MailboxAutoReplyStatus_$timestamp.xlsx"
$report | Export-Excel -Path $outputFile -WorksheetName "AutoReplyStatus" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow

Write-Host "Report saved to: $outputFile" -ForegroundColor Green

Disconnect-ExchangeOnline -Confirm:$false
