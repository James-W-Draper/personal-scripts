# Powershell script to move groups to OU's
# Summary of the Script in Layman’s Terms
# This script automates the process of moving Active Directory groups back to their original locations (called Organizational Units or OUs).
# It works by reading a list of group names and their old and current locations from an Excel file. 
# For each group, it checks if the group exists in Active Directory, and if it does, it moves it back to the specified "old" location.

# Created by James Draper
# Version 1.0 6th December 2024

# Path to the Excel file containing group information
$ExcelFilePath = "C:\Scripts\Group_Object_History.xlsx"

# Path to the log file where all actions will be recorded
$LogFilePath = "C:\Scripts\MoveGroupLog.txt"

# Function to handle logging
function Write-Log {
    param (
        [string]$Message # Message to be logged
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss" # Timestamp for log entries
    $LogMessage = "$Timestamp - $Message" # Combine timestamp and message
    Write-Output $LogMessage | Out-File -FilePath $LogFilePath -Append # Write to log file
    Write-Output $LogMessage # Write to console
}

# Ensure the ImportExcel module is installed
Write-Log "Checking for ImportExcel module..."
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Log "ImportExcel module not found. Installing..."
    try {
        Install-Module -Name ImportExcel -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
        Write-Log "Successfully installed ImportExcel module."
    } catch {
        Write-Log "Failed to install ImportExcel module. Error: $_"
        Exit 1
    }
} else {
    Write-Log "ImportExcel module is already installed."
}

# Import the Active Directory module
# This module provides cmdlets for managing Active Directory objects
# Ensure the ActiveDirectory module is installed
Write-Log "Checking for ActiveDirectory module..."
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Log "ActiveDirectory module not found. Installing..."
    try {
        Install-Module -Name ActiveDirectory -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
        Write-Log "Successfully installed ActiveDirectory module."
    } catch {
        Write-Log "Failed to install ActiveDirectory module. Error: $_"
        Exit 1
    }
} else {
    Write-Log "ActiveDirectory module is already installed."
}


# Ensure a fresh log file is created or the existing one is cleared
# This prevents old log entries from being mixed with new ones
New-Item -Path $LogFilePath -ItemType File -Force | Out-Null

# Log the start of the script execution
Write-Log "Starting the group move script..."

# Import data from the Excel file
Write-Log "Reading the spreadsheet..."
$ExcelData = Import-Excel -Path $ExcelFilePath

# Loop through each row in the Excel spreadsheet
foreach ($Row in $ExcelData) {
    # Extract group name, current OU, and old OU from the row
    $GroupName = $Row.'Group Name'
    $NewOU = $Row.'New OU' # This isn't really needed at this point, but it's probably useful in the future.
    $OldOU = $Row.'Old OU' # This is the OU that the group is going to be moved to, assuming its move is being reversed. I probably need to make these clearer.

    # Log the processing of this group
    Write-Log "Processing group: $GroupName"

    # Check if the group exists in Active Directory
    $Group = Get-ADGroup -Filter "Name -eq '$GroupName'" -ErrorAction SilentlyContinue
    if ($Group) {
        # If the group is found, log its discovery
        Write-Log "Group $GroupName found in Active Directory."

        # Attempt to move the group to its old OU
        # The Move-ADObject cmdlet moves the group back to the specified OU
        try {
            Move-ADObject -Identity $Group.DistinguishedName -TargetPath $OldOU
            Write-Log "Successfully moved group $GroupName back to $OldOU."
        } catch {
            # Log any errors that occur during the move
            Write-Log "Failed to move group $GroupName. Error: $_"
        }
    } else {
        # If the group is not found, log this information
        Write-Log "Group $GroupName not found in Active Directory."
    }
}

# Log the completion of the script execution
Write-Log "Script execution completed."
