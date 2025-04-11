<#
.SYNOPSIS
    Generates a summary report of mailbox permissions in Exchange Online.

.DESCRIPTION
    This script connects to Exchange Online and generates a detailed report of:
      - Full Access Permissions
      - Send As Permissions
      - Send On Behalf Permissions

    Each permission type is exported to its own worksheet within an Excel workbook
    using the ImportExcel module. Useful for access reviews, audits, and compliance.

.PARAMETER OutputPath
    Optional. Specifies the output directory. Defaults to "C:\Scripts" if not provided.

.EXAMPLE
    .\Get-MailboxPermissionSummary.ps1

.EXAMPLE
    .\Get-MailboxPermissionSummary.ps1 -OutputPath "D:\Reports"

.NOTES
    Requires:
    - ExchangeOnlineManagement module
    - ImportExcel module
    - Appropriate RBAC permissions in Exchange Online

    Author: James Draper
#>

[CmdletBinding()]
param (
    [string]$OutputPath = "C:\Scripts"
)

# Ensure required modules
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
}
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module ImportExcel -Scope CurrentUser -Force
}

Import-Module ExchangeOnlineManagement -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop

# Connect to Exchange Online
Connect-ExchangeOnline -ShowBanner:$false

# Initialize results
$fullAccessResults = [System.Collections.Generic.List[object]]::new()
$sendAsResults     = [System.Collections.Generic.List[object]]::new()
$sendOnBehalfResults = [System.Collections.Generic.List[object]]::new()

$mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox
$total = $mailboxes.Count
$counter = 1

foreach ($mbx in $mailboxes) {
    $display = "[$counter/$total] Processing: $($mbx.DisplayName)"
    Write-Progress -Activity "Retrieving Permissions" -Status $display -PercentComplete (($counter / $total) * 100)

    # Full Access
    try {
        Get-MailboxPermission -Identity $mbx.Identity | Where-Object {
            $_.User -notmatch "NT AUTHORITY|S-1-5|SELF" -and $_.AccessRights -contains 'FullAccess'
        } | ForEach-Object {
            $fullAccessResults.Add([pscustomobject]@{
                Mailbox       = $mbx.DisplayName
                User          = $_.User
                AccessRights  = $_.AccessRights -join ", "
                IsInherited   = $_.IsInherited
                Deny          = $_.Deny
            })
        }
    } catch {}

    # Send As
    try {
        Get-RecipientPermission -Identity $mbx.Identity | Where-Object {
            $_.Trustee -notmatch "NT AUTHORITY|S-1-5|SELF"
        } | ForEach-Object {
            $sendAsResults.Add([pscustomobject]@{
                Mailbox       = $mbx.DisplayName
                User          = $_.Trustee
                AccessRights  = $_.AccessRights -join ", "
                IsInherited   = $_.IsInherited
            })
        }
    } catch {}

    # Send On Behalf
    try {
        $delegates = $mbx.GrantSendOnBehalfTo
        if ($delegates) {
            foreach ($delegate in $delegates) {
                $resolved = Get-Recipient -Identity $delegate -ErrorAction SilentlyContinue
                $sendOnBehalfResults.Add([pscustomobject]@{
                    Mailbox = $mbx.DisplayName
                    User    = $resolved.DisplayName
                })
            }
        }
    } catch {}

    $counter++
}

# Prepare export path
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$excelPath = Join-Path $OutputPath "MailboxPermissionSummary_$timestamp.xlsx"

# Export to Excel with separate sheets
$fullAccessResults     | Export-Excel -Path $excelPath -WorksheetName "FullAccess" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
$sendAsResults         | Export-Excel -Path $excelPath -WorksheetName "SendAs" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
$sendOnBehalfResults   | Export-Excel -Path $excelPath -WorksheetName "SendOnBehalf" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow

Write-Host "Report saved to: $excelPath" -ForegroundColor Green

# Disconnect
Disconnect-ExchangeOnline -Confirm:$false
