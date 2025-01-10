# Audit log collection script for Copilot

# Check if Exchange Online module is already installed
$moduleInstalled = Get-Module -ListAvailable | Where-Object { $_.Name -eq 'ExchangeOnlineManagement' }

if ($null -eq $moduleInstalled) {

    # Exchange Online module is not installed, attempt to install it
    try {
        Write-Host "Installing Exchange Online module..."
        Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
    } catch {
        Write-Host "Failed to install Exchange Online module: $_"
        exit
    }
}

# Import the Exchange Online module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
try {
    Connect-ExchangeOnline
    Write-Host "Connected to Exchange Online."
} catch {
    Write-Host "Failed to connect to Exchange Online: $_"
    exit
}

# Define the folder path
$folderPath = "C:\Users\draperj\Enstargroup\Copilot CoE - Program Team - Program Team\new audit log"

# Check if the folder exists; create it if it doesn't
if (!(Test-Path -Path $folderPath)) {
    New-Item -ItemType Directory -Path $folderPath -Force | Out-Null
    Write-Host "Created directory: $folderPath"
}

# Path to the output CSV file
$outputCsv = Join-Path $folderPath "Copilot-Auditlogs.csv"

# Path to the log file storing the last processed timestamp
$outputlog = Join-Path $folderPath "Copilot-lastupdate.txt"

# Initialize $lastProcessedTime
$lastProcessedTime = $null

# Check if the CSV file exists
if (Test-Path $outputCsv) {
    try {
        # Read the last line of the CSV file efficiently
        $lastLine = Get-Content $outputCsv -Tail 1
        if ($lastLine) {
            # Convert the last line from CSV format to an object
            $lastEntry = $lastLine | ConvertFrom-Csv
            $creationDateString = $lastEntry.CreationDate
            if ($creationDateString) {
                $lastProcessedTime = [DateTime]$creationDateString
                # Add one second to avoid reprocessing the last record
                $lastProcessedTime = $lastProcessedTime.AddSeconds(1)
            }
        }
    } catch {
        Write-Warning "Failed to read the last entry from CSV: $_"
    }
}

# If $lastProcessedTime is still null, try to get it from the log file
if (-not $lastProcessedTime) {
    $lastProcessedTimeFromLog = Get-Content $outputlog -ErrorAction SilentlyContinue | Select-Object -Last 1
    if ($lastProcessedTimeFromLog) {
        $lastProcessedTime = [DateTime]$lastProcessedTimeFromLog
    }
}

# If $lastProcessedTime is still null, default to 7 days ago
if (-not $lastProcessedTime) {
    $lastProcessedTime = (Get-Date).AddDays(-7)
}

# Define start and end dates
$startDate = $lastProcessedTime
$endDate = Get-Date

# Check if CSV file exists for export logic
$csvExists = Test-Path $outputCsv

# Loop through each day from the last processed time to now
for ($date = $startDate.Date; $date -le $endDate.Date; $date = $date.AddDays(1)) {

    try {
        # Initialize variables for pagination
        $sessionId = [Guid]::NewGuid()
        $moreResults = $true
        $nextPage = $null

        while ($moreResults) {
            # Adjust start and end dates within the loop
            if ($date -eq $startDate.Date) {
                # On the first day, start from the exact $startDate
                $searchStartDate = $startDate
            } else {
                # On subsequent days, start from midnight
                $searchStartDate = $date
            }

            if ($date -eq $endDate.Date) {
                # On the last day, end at the exact $endDate
                $searchEndDate = $endDate
            } else {
                # On other days, end at the end of the day
                $searchEndDate = $date.AddDays(1).AddSeconds(-1)
            }

            # Search the Unified Audit Log for the specified time range
            $results = Search-UnifiedAuditLog -StartDate $searchStartDate -EndDate $searchEndDate -RecordType CopilotInteraction -ResultSize 5000 -SessionId $sessionId -NextPage $nextPage

            # Check if results are found
            if ($results) {
                # If CSV does not exist, include headers
                if (-not $csvExists) {
                    $results | Export-Csv -NoTypeInformation -Path $outputCsv
                    $csvExists = $true  # Set to true after creating the file
                } else {
                    $results | Export-Csv -NoTypeInformation -Append -Path $outputCsv -NoClobber
                }

                # Check if more results are available
                if ($results.Count -eq 5000) {
                    Write-Warning "Maximum number of records retrieved (5000) for date $date. There may be more records. Continuing to next page."
                    $moreResults = $true
                    # Continue pagination
                    $nextPage = $results[$results.Count - 1].ResultId
                } else {
                    $moreResults = $false
                }
            } else {
                $moreResults = $false
            }
        }

    } catch {
        Write-Warning "An error occurred while searching audit logs for date ${date}: $($_)"

        continue
    }
}

# Get the current timestamp
$timestamp = Get-Date

# Update the last processed time in the log file
Set-Content -Path $outputlog -Value $timestamp

# Print the last processed time
Write-Host "The file was last updated at: $timestamp"
