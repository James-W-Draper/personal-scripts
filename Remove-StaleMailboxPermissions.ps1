<#
.SYNOPSIS
    Removes mailbox delegate permissions (FullAccess, SendAs, SendOnBehalf) assigned to users
    who no longer exist in Azure AD (stale accounts).

.DESCRIPTION
    This script audits all mailboxes in Exchange Online for delegated permissions assigned to stale users
    (i.e., users no longer in Azure AD). It supports preview mode to simulate changes before applying them.

    Permissions targeted for removal:
        - Full Access (MailboxPermission)
        - Send As (RecipientPermission)
        - Send On Behalf (GrantSendOnBehalfTo)

.PARAMETER PreviewOnly
    If specified, changes will only be simulated and logged, not applied.

.PARAMETER OutputPath
    Optional. Directory to save the report file. Defaults to "C:\Scripts".

.EXAMPLE
    .\Remove-StaleMailboxPermissions.ps1 -PreviewOnly

.EXAMPLE
    .\Remove-StaleMailboxPermissions.ps1 -OutputPath "D:\Reports"

.NOTES
    Requires:
    - ExchangeOnlineManagement
    - Microsoft.Graph.Users (for user validation)
    - ImportExcel (for report formatting)
#>

[CmdletBinding()]
param (
    [switch]$PreviewOnly,
    [string]$OutputPath = "C:\Scripts"
)

# Ensure required modules
foreach ($module in @("ExchangeOnlineManagement", "Microsoft.Graph.Users", "ImportExcel")) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Install-Module -Name $module -Scope CurrentUser -Force
    }
}

Import-Module ExchangeOnlineManagement
Import-Module Microsoft.Graph.Users
Import-Module ImportExcel

Connect-ExchangeOnline -ShowBanner:$false
Connect-MgGraph -Scopes "User.Read.All"

$report = [System.Collections.Generic.List[object]]::new()
$mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox
$total = $mailboxes.Count
$counter = 1

foreach ($mbx in $mailboxes) {
    $display = "[$counter/$total] Processing: $($mbx.UserPrincipalName)"
    Write-Progress -Activity "Auditing Permissions" -Status $display -PercentComplete (($counter / $total) * 100)
    $counter++

    # Full Access
    Get-MailboxPermission -Identity $mbx.Identity | Where-Object {
        $_.User -notmatch "NT AUTHORITY|S-1-5|SELF" -and $_.IsInherited -eq $false
    } | ForEach-Object {
        $user = $_.User.ToString()
        if (-not (Get-MgUser -Filter "userPrincipalName eq '$user'" -ErrorAction SilentlyContinue)) {
            $report.Add([pscustomobject]@{
                Mailbox = $mbx.UserPrincipalName
                Type = "FullAccess"
                AssignedTo = $user
                Action = if ($PreviewOnly) { "Would Remove" } else { "Removed" }
            })
            if (-not $PreviewOnly) {
                Remove-MailboxPermission -Identity $mbx.Identity -User $user -AccessRights FullAccess -Confirm:$false
            }
        }
    }

    # SendAs
    Get-RecipientPermission -Identity $mbx.Identity | Where-Object {
        $_.Trustee -notmatch "NT AUTHORITY|S-1-5|SELF"
    } | ForEach-Object {
        $user = $_.Trustee.ToString()
        if (-not (Get-MgUser -Filter "userPrincipalName eq '$user'" -ErrorAction SilentlyContinue)) {
            $report.Add([pscustomobject]@{
                Mailbox = $mbx.UserPrincipalName
                Type = "SendAs"
                AssignedTo = $user
                Action = if ($PreviewOnly) { "Would Remove" } else { "Removed" }
            })
            if (-not $PreviewOnly) {
                Remove-RecipientPermission -Identity $mbx.Identity -Trustee $user -AccessRights SendAs -Confirm:$false
            }
        }
    }

    # SendOnBehalf
    foreach ($delegate in $mbx.GrantSendOnBehalfTo) {
        $resolved = Get-Recipient -Identity $delegate -ErrorAction SilentlyContinue
        if ($resolved -and ($resolved.UserType -ne "Guest") -and (-not (Get-MgUser -Filter "id eq '$($resolved.ExternalDirectoryObjectId)'" -ErrorAction SilentlyContinue))) {
            $report.Add([pscustomobject]@{
                Mailbox = $mbx.UserPrincipalName
                Type = "SendOnBehalf"
                AssignedTo = $resolved.DisplayName
                Action = if ($PreviewOnly) { "Would Remove" } else { "Removed" }
            })
            if (-not $PreviewOnly) {
                Set-Mailbox -Identity $mbx.Identity -GrantSendOnBehalfTo @{remove = $resolved.Identity}
            }
        }
    }
}

# Output report
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFile = Join-Path $OutputPath "StaleMailboxPermissions_$timestamp.xlsx"
$report | Export-Excel -Path $outputFile -WorksheetName "StalePermissions" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow

Write-Host "Report saved to: $outputFile" -ForegroundColor Green

Disconnect-MgGraph
Disconnect-ExchangeOnline -Confirm:$false
