# File Deletion Script
# (V4) LAUNCH POWERSHELL AS ADMIN AND DOMAIN ADMIN PERMISSIONS (V4)

# Set verbose preference to log detailed information for troubleshooting
$VerbosePreference = 'Continue'

# Import the excel Powershell module for handling Excel files
Install-Module -Name ImportExcel -Scope CurrentUser

# Define file paths and dynamic naming for output files
$inputCsvPath = "C:\scripts\CHG0077067 - USE2-PRD-FIL-01.csv"  # Path to the input CSV file

$computerName = $env:COMPUTERNAME  # Get the current computer name
$currentDate = Get-Date -Format "yyyy-MM-dd"  # Get the current date in yyyy-MM-dd format
$excelFileName = "DeletionReport_${computerName}_$currentDate.xlsx"  # Name for the output Excel file
$transcriptFileName = "Transcript_${computerName}_$currentDate.txt"  # Name for the transcript file
$outputExcelPath = "C:\scripts\$excelFileName"  # Full path for the output Excel file
$transcriptPath = "C:\scripts\$transcriptFileName"  # Full path for the transcript file

# Start logging all actions to a transcript file for later review
Start-Transcript -Path $transcriptPath -Force

# Function to check and override file attributes if needed
Function Get-FileAttributes {
    Param (
        [string]$filePath  # The file path to check and override attributes
    )

    $file = Get-Item -LiteralPath $filePath -Force  # Get the file item with force to bypass restrictions
    $attributes = $file.Attributes  # Retrieve the file attributes

    # Check if the file has Hidden, System, or ReadOnly attributes
    $isHidden = $attributes -band [System.IO.FileAttributes]::Hidden
    $isSystem = $attributes -band [System.IO.FileAttributes]::System
    $isReadOnly = $attributes -band [System.IO.FileAttributes]::ReadOnly

    # Override the attributes to Normal if any of the above attributes are set
    if ($isHidden -or $isSystem -or $isReadOnly) {
        Set-ItemProperty -LiteralPath $filePath -Name Attributes -Value 'Normal'
        return "Overridden"
    }
    return "No Override"
}

# Function to escape special characters in file paths
Function Optimize-SpecialCharacters {
    Param (
        [string]$filePath  # The file path to escape special characters in
    )

    # Escape special characters [] in the file path
    return $filePath -replace '([[]])', '`$1'
}

# Read the list of files from the CSV file
$fileList = Import-Csv -Path $inputCsvPath
$results = @()  # Initialize an array to store results

# Loop through each file in the CSV list
foreach ($file in $fileList) {
    $filePath = Optimize-SpecialCharacters -filePath $file.Path  # Escape special characters in the file path
    $fileExists = Test-Path -LiteralPath $filePath  # Check if the file exists
    $checkTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"  # Record the current time for the check
    $deletionTime = $null  # Initialize deletion time
    $deletionSuccess = $false  # Initialize deletion success status
    $deletionMessage = ""  # Initialize deletion message
    $postDeletionCheck = $null  # Initialize post-deletion check
    $ownershipTaken = $false  # Initialize ownership taken status
    $attributeOverrideResult = "Not Checked"  # Initialize attribute override result

    if ($fileExists) {
        # Override file attributes if necessary
        $attributeOverrideResult = Get-FileAttributes -filePath $filePath

        try {
            # Attempt to delete the file without confirmation
            Remove-Item -LiteralPath $filePath -Force
            $deletionSuccess = $true  # Set deletion success to true
            $deletionMessage = "File Deleted Successfully"  # Set deletion message to success
            $deletionTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"  # Record the deletion time
        } catch {
            # If an error occurs during deletion, record the error message
            $deletionMessage = "Error Deleting File: $_"
        }
    } else {
        # If the file does not exist, set the appropriate message
        $deletionMessage = "File Not Found"
    }

    # Check if the file is gone after the deletion attempt
    $postDeletionCheck = -not (Test-Path -LiteralPath $filePath)

    # Add the results to the results array
    $results += [PSCustomObject]@{
        FilePath = $filePath  # The file path
        ExistsInitially = $fileExists  # Whether the file existed initially
        HasDeletePermission = $hasPermission  # Whether the script has permission to delete the file
        OwnershipTaken = $ownershipTaken  # Whether ownership was taken (not implemented in this script)
        DeletionAttempted = $deletionTime  # The time the deletion was attempted
        DeletionSuccess = $deletionSuccess  # Whether the deletion was successful
        DeletionMessage = $deletionMessage  # Message regarding the deletion attempt
        FileGonePostDeletion = $postDeletionCheck  # Whether the file is gone after the deletion attempt
        CheckedOn = $checkTime  # The time the file was checked
        AttributeOverrideResult = $attributeOverrideResult  # The result of attribute override attempt
    }
}

# Export the results to an Excel file
$results | Export-Excel -Path $outputExcelPath -WorksheetName "DeletionReport" -AutoSize

Write-Host "Script completed. The results have been exported to $outputExcelPath"

# Stop the transcript logging
Stop-Transcript

Write-Host "Transcript saved to $transcriptPath"

