# Audit Log Collection Script for Copilot

# This script collects Copilot interaction audit logs from Exchange Online.
# It saves the last processed time to a text file to resume from where it left off.

# Check if the Exchange Online Management module is already installed
$moduleName = 'ExchangeOnlineManagement'
$moduleInstalled = Get-Module -ListAvailable | Where-Object { $_.Name -eq $moduleName }

if ($null -eq $moduleInstalled) {
    # If the module is not installed, attempt to install it
    try {
        Write-Host "Installing $moduleName module..."
        Install-Module -Name $moduleName -Force -AllowClobber -Scope CurrentUser
    } catch {
        # If installation fails, output the error and exit
        Write-Host "Failed to install $moduleName module: $($_)"
        exit
    }
}

# Import the Exchange Online Management module
Import-Module $moduleName

# Connect to Exchange Online
try {
    # Uncomment this line when using the script
    Connect-ExchangeOnline
    Write-Host "Connected to Exchange Online."
} catch {
    # If connection fails, output the error and exit
    Write-Host "Failed to connect to Exchange Online: $($_)"
    exit
}

# Define the folder path where the audit logs will be stored
$folderPath = "C:\Users\draperj\OneDrive - Enstargroup\copilot\"

# Check if the folder exists; if not, create it
if (-not (Test-Path -Path $folderPath)) {
    New-Item -ItemType Directory -Path $folderPath -Force | Out-Null
    Write-Host "Created directory: $folderPath"
}

# Path to the output CSV file
$outputCsv = Join-Path $folderPath "Copilot-Auditlogs.csv"

# Path to the last processed date text file
$lastProcessedDateFile = Join-Path $folderPath "LastProcessedDate.txt"

# Initialize the start date for the log collection
$lastProcessedTime = $null

# Check if the last processed date file exists
if (Test-Path $lastProcessedDateFile) {
    try {
        # Read the last processed date from the text file
        $lastProcessedTimeString = Get-Content -Path $lastProcessedDateFile
        $lastProcessedTime = [DateTime]::Parse($lastProcessedTimeString, [System.Globalization.CultureInfo]::InvariantCulture)
        Write-Host "Resuming data collection from: $lastProcessedTime"
    } catch {
        Write-Warning "Failed to parse the last processed date from file: $($_)"
    }
}

# If the last processed time is still null, set it to the start of the year
if (-not $lastProcessedTime) {
    $lastProcessedTime = Get-Date -Year (Get-Date).Year -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
    Write-Host "No last processed date found. Collecting logs from the start of the year: $($lastProcessedTime)"
}

# Define the start and end dates for the audit log retrieval
$startDate = $lastProcessedTime
$endDate = Get-Date

# Check if CSV file exists for export logic
$csvExists = Test-Path $outputCsv

# Define the time increment for each audit log search interval (e.g., 1 hour)
$timeIncrement = [TimeSpan]::FromHours(1)

# Calculate the total number of intervals for progress tracking
$totalIntervals = [Math]::Ceiling(($endDate - $startDate).TotalHours / $timeIncrement.TotalHours)
$currentInterval = 0  # Initialize the current interval counter

# Initialize the current start time for the loop
$currentStartTime = $startDate

# Loop through each time interval from the last processed time to now
while ($currentStartTime -lt $endDate) {
    # Determine the end time for the current interval
    $currentEndTime = $currentStartTime.Add($timeIncrement)
    if ($currentEndTime -gt $endDate) {
        # Adjust the end time if it exceeds the current time
        $currentEndTime = $endDate
    }

    # Increment the interval counter
    $currentInterval++

    # Calculate the percentage complete for the progress bar
    $percentComplete = [Math]::Round(($currentInterval / $totalIntervals) * 100, 2)

    # Display the progress update
    Write-Progress -Activity "Collecting Audit Logs" `
                   -Status "Processing $($currentStartTime) to $($currentEndTime)" `
                   -PercentComplete $percentComplete

    try {
        # Search the Unified Audit Log for the specified time range
        $results = Search-UnifiedAuditLog `
            -StartDate $currentStartTime `
            -EndDate $currentEndTime `
            -RecordType CopilotInteraction `
            -ResultSize 5000

        # Check if any results were found
        if ($results) {
            # Define the desired property order
            $desiredProperties = @(
                'RecordId',
                'CreationDate',
                'RecordType',
                'Operation',
                'UserId',
                'AuditData',
                'AssociatedAdminUnits',
                'AssociatedAdminUnitsNames'
            )

            # Select the desired properties in the specified order
            $formattedResults = $results | Select-Object $desiredProperties

            # Ensure consistent formatting for CreationDate
            $formattedResults | ForEach-Object {
                $_.CreationDate = $_.CreationDate.ToString('yyyy-MM-ddTHH:mm:ss.fffffffZ')
            }

            if (-not $csvExists) {
                # If the CSV file doesn't exist, export with headers
                $formattedResults | Export-Csv -NoTypeInformation -Path $outputCsv
                $csvExists = $true  # Update the flag since the file now exists
                Write-Host "CSV file created and saved at: $outputCsv"
            } else {
                # Append data without headers to avoid duplicates
                $formattedResults | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | Add-Content -Path $outputCsv
                Write-Host "Data appended to CSV file at: $outputCsv"
            }

            # Check if the maximum number of records was retrieved
            if ($results.Count -eq 5000) {
                Write-Warning "Maximum number of records retrieved (5000) between $($currentStartTime) and $($currentEndTime). Consider reducing the time increment."
            }
        }
    } catch {
        # Output a warning if an error occurs during the search
        Write-Warning "An error occurred while searching audit logs between $($currentStartTime) and $($currentEndTime): $($_)"
    }

    # Move to the next time interval
    $currentStartTime = $currentEndTime
}

# Save the last processed time to the text file
$currentEndTimeString = $currentEndTime.ToString('yyyy-MM-ddTHH:mm:ss')
Set-Content -Path $lastProcessedDateFile -Value $currentEndTimeString
Write-Host "Last processed date saved to: $lastProcessedDateFile"

# Clear the progress bar upon completion
Write-Progress -Activity "Collecting Audit Logs" -Completed

# Output the completion message
Write-Host "Audit log collection completed successfully. CSV file saved at: $outputCsv"
