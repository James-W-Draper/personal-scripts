<#
.SYNOPSIS
Scans Exchange Online mailboxes and exports mailbox permissions including Send As, Full Access, and Send on Behalf.

.DESCRIPTION
This script connects to Exchange Online and audits all User and Shared mailboxes.
It exports permission assignments to a CSV and optionally groups them by 'AssignedTo' or 'Permission'.
You can suppress the progress bar and customise the grouping using parameters.

.PARAMETER ShowProgress
Optional. If specified (default = $true), shows a progress bar while scanning.

.PARAMETER GroupBy
Optional. Specify 'AssignedTo', 'Permission', or 'None' to control how the output is grouped.

.EXAMPLE
.\Get-MailboxPermissionsReportGrouped.ps1
# Runs the script with progress and no grouping.

.EXAMPLE
.\Get-MailboxPermissionsReportGrouped.ps1 -ShowProgress:$false
# Runs silently without showing a progress bar.

.EXAMPLE
.\Get-MailboxPermissionsReportGrouped.ps1 -GroupBy AssignedTo
# Shows progress and exports a grouped report by delegate.

.EXAMPLE
.\Get-MailboxPermissionsReportGrouped.ps1 -ShowProgress:$false -GroupBy Permission
# Silent run, grouped by permission type (e.g. FullAccess, SendAs, etc.).

.NOTES
- Requires the ExchangeOnlineManagement module
- Requires appropriate permissions (e.g. View-Only Recipients, Mailbox Reader)
- Exports CSVs to C:\Temp by default
#>

param (
    [bool]$ShowProgress = $true,
    [ValidateSet("AssignedTo", "Permission", "None")]
    [string]$GroupBy = "None"
)

Write-Host "`n📦 Fetching mailboxes..." -ForegroundColor Cyan
$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox, SharedMailbox -ResultSize Unlimited `
    -PropertySet Delivery `
    -Properties RecipientTypeDetails, DisplayName, GrantSendOnBehalfTo |
    Select DisplayName, UserPrincipalName, RecipientTypeDetails, GrantSendOnBehalfTo

if (-not $Mbx) {
    Write-Error "❌ No mailboxes found. Script exiting..." -ErrorAction Stop
}

$Report = [System.Collections.Generic.List[Object]]::new()
$Total = $Mbx.Count
$ProgressDelta = 100 / $Total
$PercentComplete = 0
$Index = 0

foreach ($M in $Mbx) {
    $Index++
    if ($ShowProgress) {
        $ProgressMsg = "$($M.DisplayName) [$Index of $Total]"
        Write-Progress -Activity "Auditing mailbox permissions..." -Status $ProgressMsg -PercentComplete $PercentComplete
        $PercentComplete += $ProgressDelta
    }

    # --- Send As Permissions ---
    try {
        $Permissions = Get-ExoRecipientPermission -Identity $M.UserPrincipalName | Where-Object { $_.Trustee -ne "NT AUTHORITY\SELF" }
        foreach ($Permission in $Permissions) {
            $Report.Add([PSCustomObject]@{
                Mailbox     = $M.DisplayName
                UPN         = $M.UserPrincipalName
                Permission  = $Permission.AccessRights
                AssignedTo  = $Permission.Trustee
                MailboxType = $M.RecipientTypeDetails
            })
        }
    } catch {
        Write-Warning "SendAs permissions failed for $($M.UserPrincipalName): $_"
    }

    # --- Full Access Permissions ---
    try {
        $Permissions = Get-ExoMailboxPermission -Identity $M.UserPrincipalName | Where-Object { $_.User -like "*@*" }
        foreach ($Permission in $Permissions) {
            $Report.Add([PSCustomObject]@{
                Mailbox     = $M.DisplayName
                UPN         = $M.UserPrincipalName
                Permission  = $Permission.AccessRights
                AssignedTo  = $Permission.User
                MailboxType = $M.RecipientTypeDetails
            })
        }
    } catch {
        Write-Warning "FullAccess permissions failed for $($M.UserPrincipalName): $_"
    }

    # --- Send on Behalf Of ---
    if ($M.GrantSendOnBehalfTo) {
        foreach ($Trustee in $M.GrantSendOnBehalfTo) {
            try {
                $recipient = Get-ExoRecipient -Identity $Trustee
                $Report.Add([PSCustomObject]@{
                    Mailbox     = $M.DisplayName
                    UPN         = $M.UserPrincipalName
                    Permission  = "Send on Behalf Of"
                    AssignedTo  = $recipient.PrimarySmtpAddress
                    MailboxType = $M.RecipientTypeDetails
                })
            } catch {
                Write-Warning "SendOnBehalf resolution failed for $Trustee on $($M.UserPrincipalName)"
            }
        }
    }
}

# === EXPORT: Main Report ===
$csvBasePath = "C:\Temp"
$timestamp = Get-Date -Format "yyyyMMdd-HHmm"
$mainReportPath = Join-Path $csvBasePath "MailboxPermissions_$timestamp.csv"

$Report | Sort-Object @{Expression = { $_.MailboxType }; Ascending = $false }, Mailbox |
    Export-Csv -Path $mainReportPath -NoTypeInformation -Encoding UTF8

Write-Host "`n✅ Main report exported to:`n$mainReportPath"

# === OPTIONAL GROUPED EXPORT ===
if ($GroupBy -ne "None") {
    $groupedReportPath = Join-Path $csvBasePath "MailboxPermissions_GroupedBy_$GroupBy`_$timestamp.csv"

    $Grouped = $Report | Group-Object -Property $GroupBy
    $ExportList = foreach ($Group in $Grouped) {
        foreach ($item in $Group.Group) {
            $item
        }
    }

    $ExportList | Export-Csv -Path $groupedReportPath -NoTypeInformation -Encoding UTF8
    Write-Host "📊 Grouped report (by $GroupBy) exported to:`n$groupedReportPath"
}

Write-Host "`n🔍 $($Mbx.Count) mailboxes scanned.`n"
