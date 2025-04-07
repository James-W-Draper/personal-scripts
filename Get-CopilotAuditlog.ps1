# Audit Log Collection Script for Copilot

# This script collects Copilot interaction audit logs from Exchange Online.
# It saves the last processed time to a text file to resume from where it left off.

# Check if the Exchange Online Management module is already installed
$moduleName = 'ExchangeOnlineManagement'
$moduleInstalled = Get-Module -ListAvailable | Where-Object { $_.Name -eq $moduleName }

if ($null -eq $moduleInstalled) {
    try {
        Write-Host "Installing $moduleName module..."
        Install-Module -Name $moduleName -Force -AllowClobber -Scope CurrentUser
    } catch {
        Write-Host "Failed to install $moduleName module: $($_)"
        exit
    }
}

Import-Module $moduleName

# Connect to Exchange Online
try {
    Connect-ExchangeOnline
    Write-Host "Connected to Exchange Online."
} catch {
    Write-Host "Failed to connect to Exchange Online: $($_)"
    exit
}

# Define the folder path where the audit logs will be stored
$folderPath = "C:\Scripts\Copilot-AuditLogs"

# Create the folder if it doesn't exist
if (-not (Test-Path -Path $folderPath)) {
    New-Item -ItemType Directory -Path $folderPath -Force | Out-Null
    Write-Host "Created directory: $folderPath"
}

# Define paths for the CSV file and the last processed timestamp
$outputCsv = Join-Path $folderPath "Copilot-Auditlogs.csv"
$lastProcessedDateFile = Join-Path $folderPath "LastProcessedDate.txt"

$lastProcessedTime = $null

# Read the last processed timestamp if available
if (Test-Path $lastProcessedDateFile) {
    try {
        $lastProcessedTimeString = Get-Content -Path $lastProcessedDateFile
        $lastProcessedTime = [DateTime]::Parse($lastProcessedTimeString, [System.Globalization.CultureInfo]::InvariantCulture)
        Write-Host "Resuming data collection from: $lastProcessedTime"
    } catch {
        Write-Warning "Failed to parse the last processed date: $($_)"
    }
}

if (-not $lastProcessedTime) {
    $lastProcessedTime = Get-Date -Year (Get-Date).Year -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
    Write-Host "No previous timestamp found. Starting from: $($lastProcessedTime)"
}

$startDate = $lastProcessedTime
$endDate = Get-Date
$csvExists = Test-Path $outputCsv
$timeIncrement = [TimeSpan]::FromHours(1)
$totalIntervals = [Math]::Ceiling(($endDate - $startDate).TotalHours / $timeIncrement.TotalHours)
$currentInterval = 0
$currentStartTime = $startDate

while ($currentStartTime -lt $endDate) {
    $currentEndTime = $currentStartTime.Add($timeIncrement)
    if ($currentEndTime -gt $endDate) {
        $currentEndTime = $endDate
    }

    $currentInterval++
    $percentComplete = [Math]::Round(($currentInterval / $totalIntervals) * 100, 2)

    Write-Progress -Activity "Collecting Audit Logs" `
                   -Status "Processing $($currentStartTime) to $($currentEndTime)" `
                   -PercentComplete $percentComplete

    try {
        $results = Search-UnifiedAuditLog `
            -StartDate $currentStartTime `
            -EndDate $currentEndTime `
            -RecordType CopilotInteraction `
            -ResultSize 5000

        if ($results) {
            $desiredProperties = @(
                'RecordId', 'CreationDate', 'RecordType', 'Operation',
                'UserId', 'AuditData', 'AssociatedAdminUnits', 'AssociatedAdminUnitsNames'
            )

            $formattedResults = $results | Select-Object $desiredProperties

            $formattedResults | ForEach-Object {
                $_.CreationDate = $_.CreationDate.ToString('yyyy-MM-ddTHH:mm:ss.fffffffZ')
            }

            if (-not $csvExists) {
                $formattedResults | Export-Csv -NoTypeInformation -Path $outputCsv
                $csvExists = $true
                Write-Host "CSV file created at: $outputCsv"
            } else {
                $formattedResults | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | Add-Content -Path $outputCsv
                Write-Host "Data appended to: $outputCsv"
            }

            if ($results.Count -eq 5000) {
                Write-Warning "Max results (5000) hit for interval $($currentStartTime) - $($currentEndTime). Consider reducing time range."
            }
        }
    } catch {
        Write-Warning "Error during audit search from $($currentStartTime) to $($currentEndTime): $($_)"
    }

    $currentStartTime = $currentEndTime
}

# Save the last processed time
$currentEndTimeString = $currentEndTime.ToString('yyyy-MM-ddTHH:mm:ss')
Set-Content -Path $lastProcessedDateFile -Value $currentEndTimeString
Write-Host "Last processed date saved."

Write-Progress -Activity "Collecting Audit Logs" -Completed
Write-Host "Audit log collection completed. Output file: $outputCsv"
