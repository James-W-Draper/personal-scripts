<#
.SYNOPSIS
Exports a report of all Microsoft 365 Unified Groups and their respective owners.

.DESCRIPTION
This script retrieves all Unified Groups (Microsoft 365 Groups) in the tenant using Exchange Online PowerShell.
It collects each group's display name along with the primary SMTP addresses of their owners, then exports the data
to an Excel file using the ImportExcel module.

.PARAMETER OutputPath
Specifies the output directory where the Excel report should be saved. Defaults to 'C:\scripts\'.

.EXAMPLE
.\Export-UnifiedGroupOwners.ps1 -OutputPath "D:\Reports\"

.NOTES
Requires the Exchange Online module and the ImportExcel module.
Run with sufficient permissions to query Unified Groups and Mailboxes.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "C:\scripts\"
)

# Ensure output path ends with a backslash
if (-not ($OutputPath.EndsWith('\'))) {
    $OutputPath += '\'
}

# Create the directory if it doesn't exist
if (-not (Test-Path -Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

# Retrieve all Unified Groups
$UnifiedGroups = Get-UnifiedGroup -ResultSize Unlimited

# Prepare results array
$GroupOwnershipResults = @()

# Loop through groups and collect owner email addresses
foreach ($Group in $UnifiedGroups) {
    $GroupOwners = Get-UnifiedGroupLinks -Identity $Group.Identity -LinkType Owners
    $OwnerEmails = @()

    foreach ($Owner in $GroupOwners) {
        try {
            $Mailbox = Get-Mailbox -Identity $Owner -ErrorAction Stop
            $OwnerEmails += $Mailbox.PrimarySmtpAddress
        } catch {
            Write-Warning "Failed to retrieve mailbox for owner '$Owner'"
        }
    }

    $GroupOwnershipResults += [PSCustomObject]@{
        GroupName   = $Group.DisplayName
        OwnerEmails = ($OwnerEmails -join '; ')
    }
}

# Generate timestamped filename
$Timestamp = Get-Date -Format 'yyyyMMddTHHmmss'
$OutputFile = "${OutputPath}${Timestamp}_UnifiedGroupOwners.xlsx"

# Export to Excel
$ExportExcelParams = @{
    Path         = $OutputFile
    WorksheetName = 'GroupOwners'
    BoldTopRow   = $true
    AutoSize     = $true
    AutoFilter   = $true
    FreezeTopRow = $true
}

$GroupOwnershipResults | Export-Excel @ExportExcelParams

Write-Output "Report saved to: $OutputFile"
