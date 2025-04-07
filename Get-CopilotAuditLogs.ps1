<#
.SYNOPSIS
Collects Copilot interaction audit logs from Exchange Online and appends new entries to a CSV.

.DESCRIPTION
This script:
- Connects to Exchange Online
- Collects CopilotInteraction records from the Unified Audit Log
- Tracks the last successfully processed timestamp using either the CSV or a fallback log
- Supports pagination (5000 record limit)
- Outputs data to a CSV and updates the timestamp log

.NOTES
- Requires the ExchangeOnlineManagement module
- Must be run by a user with permissions to access audit logs (Audit Reader or equivalent)
#>

# === MODULE LOAD ===
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    try {
        Write-Host "Installing Exchange Online module..."
        Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
    } catch {
        Write-Host "❌ Failed to install Exchange Online module: $_"
        exit
    }
}
Import-Module ExchangeOnlineManagement

# === CONNECT TO EXO ===
try {
    Connect-ExchangeOnline
    Write-Host "✅ Connected to Exchange Online."
} catch {
    Write-Host "❌ Failed to connect to Exchange Online: $_"
    exit
}

# === FILE PATHS ===
$folderPath = "C:\Scripts\CopilotAuditLogs"
if (-not (Test-Path $folderPath)) {
    New-Item -ItemType Directory -Path $folderPath -Force | Out-Null
    Write-Host "📁 Created directory: $folderPath"
}

$outputCsv  = Join-Path $folderPath "Copilot-Auditlogs.csv"
$outputLog  = Join-Path $folderPath "Copilot-LastUpdate.txt"

# === DETERMINE LAST PROCESSED TIME ===
$lastProcessedTime = $null

if (Test-Path $outputCsv) {
    try {
        $lastLine = Get-Content $outputCsv -Tail 1
        if ($lastLine) {
            $lastEntry = $lastLine | ConvertFrom-Csv
            if ($lastEntry.CreationDate) {
                $lastProcessedTime = ([DateTime]$lastEntry.CreationDate).AddSeconds(1)
            }
        }
    } catch {
        Write-Warning "⚠️ Failed to read last CSV entry: $_"
    }
}

if (-not $lastProcessedTime -and (Test-Path $outputLog)) {
    try {
        $logTime = Get-Content $outputLog | Select-Object -Last 1
        $lastProcessedTime = [DateTime]$logTime
    } catch {
        Write-Warning "⚠️ Failed to read timestamp log file."
    }
}

if (-not $lastProcessedTime) {
    $lastProcessedTime = (Get-Date).AddDays(-7)
    Write-Host "📅 Defaulting to 7 days ago: $lastProcessedTime"
}

# === LOOP THROUGH DATES ===
$startDate = $lastProcessedTime
$endDate = Get-Date
$csvExists = Test-Path $outputCsv

for ($date = $startDate.Date; $date -le $endDate.Date; $date = $date.AddDays(1)) {
    try {
        $sessionId = [guid]::NewGuid()
        $moreResults = $true
        $nextPage = $null

        while ($moreResults) {
            $searchStartDate = if ($date -eq $startDate.Date) { $startDate } else { $date }
            $searchEndDate   = if ($date -eq $endDate.Date) { $endDate } else { $date.AddDays(1).AddSeconds(-1) }

            $results = Search-UnifiedAuditLog -StartDate $searchStartDate -EndDate $searchEndDate `
                      -RecordType CopilotInteraction -ResultSize 5000 -SessionId $sessionId -NextPage $nextPage

            if ($results) {
                if (-not $csvExists) {
                    $results | Export-Csv -Path $output
