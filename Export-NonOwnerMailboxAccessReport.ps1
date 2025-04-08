<#
.SYNOPSIS
Exports a report of non-owner mailbox access events from Exchange Online audit logs.

.DESCRIPTION
This script queries the Unified Audit Log in Exchange Online to retrieve actions performed by users other than the mailbox owner. 
It supports optional filtering for external access and works with both MFA and non-MFA accounts. 
The output is saved to a timestamped CSV file for review.

.PARAMETER IncludeExternalAccess
Switch to include or exclude external user access events.

.PARAMETER StartDate
Start date for the audit log search (within 90-day window).

.PARAMETER EndDate
End date for the audit log search.

.PARAMETER Organization
Organization ID used for certificate-based authentication.

.PARAMETER ClientId
Client ID for app-based authentication.

.PARAMETER CertificateThumbprint
Certificate thumbprint for app-based authentication.

.PARAMETER UserName
User principal name for basic authentication (non-MFA).

.PARAMETER Password
Password for basic authentication (non-MFA).

.EXAMPLE
./Export-NonOwnerMailboxAccess.ps1 -IncludeExternalAccess \$true -StartDate '2024-01-01' -EndDate '2024-01-31'

#>

[CmdletBinding()]
param (
    [bool]$IncludeExternalAccess = $false,
    [Nullable[datetime]]$StartDate,
    [Nullable[datetime]]$EndDate,
    [string]$Organization,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$UserName,
    [string]$Password
)

function Connect-Exchange {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host "Installing Exchange Online Management module..."
        Install-Module ExchangeOnlineManagement -Force -Scope CurrentUser
    }

    if ($UserName -and $Password) {
        $securePassword = ConvertTo-SecureString -AsPlainText $Password -Force
        $cred = New-Object PSCredential($UserName, $securePassword)
        Connect-ExchangeOnline -Credential $cred -ShowBanner:$false
    } elseif ($Organization -and $ClientId -and $CertificateThumbprint) {
        Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -Organization $Organization -ShowBanner:$false
    } else {
        Connect-ExchangeOnline -ShowBanner:$false
    }
}

if ($StartDate -and -not $EndDate -or -not $StartDate -and $EndDate) {
    Write-Error "Please provide both StartDate and EndDate."
    exit
}

$StartDate = $StartDate ?? ((Get-Date).AddDays(-30)).Date
$EndDate = $EndDate ?? (Get-Date)

if ($StartDate -lt (Get-Date).AddDays(-90)) {
    Write-Error "StartDate cannot be more than 90 days ago."
    exit
}
if ($EndDate -lt $StartDate) {
    Write-Error "EndDate must be after StartDate."
    exit
}

Connect-Exchange

$outputFile = "./NonOwnerMailboxAccessReport_$((Get-Date -f 'yyyyMMdd_HHmm')).csv"
$intervalMinutes = 1440
$searchStart = $StartDate
$searchEnd = $searchStart.AddMinutes($intervalMinutes)
$ops = 'ApplyRecord','Copy','Create','FolderBind','HardDelete','MessageBind','Move','MoveToDeletedItem','RecordDelete','SendAs','SendOnBehalf','SoftDelete','Update','UpdateCalendarDelegation','UpdateFolderPermissions','UpdateInboxRules'

$totalResults = 0
$nonOwnerCount = 0

while ($true) {
    if ($searchStart -eq $searchEnd) {
        Write-Error "Start and end time are equal. Aborting."
        break
    }

    $logs = Search-UnifiedAuditLog -StartDate $searchStart -EndDate $searchEnd -Operations $ops -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000

    foreach ($entry in $logs) {
        $data = $entry.AuditData | ConvertFrom-Json

        if ($data.LogonType -eq 0) { continue }
        if (-not $IncludeExternalAccess -and $data.ExternalAccess -eq $true) { continue }

        if (($data.LogonUserSid -ne $data.MailboxOwnerSid) -or (("SendAs","SendOnBehalf" -contains $data.Operation) -and $data.UserType -eq 0)) {
            $timestamp = [datetime]::Parse($data.CreationTime).ToLocalTime()
            $logonType = switch ($data.LogonType) {
                1 { "Admin" }
                2 { "Delegated" }
                default { "Microsoft Datacenter" }
            }

            switch ($data.Operation) {
                "SendAs" { $accessedMailbox = $data.SendAsUserSMTP; $accessedBy = $data.MailboxOwnerUPN }
                "SendOnBehalf" { $accessedMailbox = $data.SendOnBehalfOfUserSMTP; $accessedBy = $data.MailboxOwnerUPN }
                default { $accessedMailbox = $data.MailboxOwnerUPN; $accessedBy = $data.UserId }
            }

            if ($accessedMailbox -ne $accessedBy) {
                $nonOwnerCount++
                [pscustomobject]@{
                    'Access Time'       = $timestamp
                    'Logon Type'        = $logonType
                    'Accessed by'       = $accessedBy
                    'Performed Operation' = $data.Operation
                    'Accessed Mailbox'  = $accessedMailbox
                    'Result Status'     = $data.ResultStatus
                    'External Access'   = $data.ExternalAccess
                } | Export-Csv -Path $outputFile -Append -NoTypeInformation
            }
        }
    }

    $totalResults += $logs.Count

    if ($searchEnd -ge $EndDate -or $logs.Count -lt 5000) { break }

    $searchStart = $searchEnd
    $searchEnd = $searchStart.AddMinutes($intervalMinutes)
    if ($searchEnd -gt $EndDate) { $searchEnd = $EndDate }
}

Disconnect-ExchangeOnline -Confirm:$false

if (Test-Path $outputFile) {
    Write-Host "Report complete: $outputFile ($nonOwnerCount non-owner records exported)" -ForegroundColor Green
    $prompt = New-Object -ComObject wscript.shell
    if ($prompt.Popup("Open the report now?", 0, "Export Complete", 4) -eq 6) {
        Invoke-Item $outputFile
    }
} else {
    Write-Warning "No records found."
}
