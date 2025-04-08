<#
.SYNOPSIS
    Moves Active Directory groups back to their original OUs as listed in an Excel file.

.DESCRIPTION
    This script automates the restoration of AD group objects to their original Organizational Units (OUs).
    It reads group names and OU paths from a spreadsheet and uses the Active Directory and ImportExcel modules
    to move the objects and log the results.

.PARAMETER ExcelFilePath
    Path to the Excel file containing columns: "Group Name", "Old OU", and "New OU".

.PARAMETER LogFilePath
    Optional: custom path to store the log file. If not provided, a dated log file is created in the script folder.

.EXAMPLE
    .\Revert-ADGroupMoves.ps1 -ExcelFilePath "C:\Scripts\Group_Object_History.xlsx"

.NOTES
    Created by James Draper - v1.1
    Updated: 6th April 2025
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$ExcelFilePath,

    [string]$LogFilePath = $(Join-Path -Path (Split-Path -Parent $ExcelFilePath) `
                                       -ChildPath ("MoveGroupLog_{0}.txt" -f (Get-Date -Format "yyyy-MM-dd_HHmmss")))
)

# Function to handle logging
function Write-Log {
    param (
        [string]$Message
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "$Timestamp - $Message"
    $LogMessage | Out-File -FilePath $LogFilePath -Append
    Write-Output $LogMessage
}

# Ensure required modules
function Ensure-Module {
    param (
        [string]$ModuleName
    )
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Log "$ModuleName module not found. Attempting to install..."
        try {
            Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
            Write-Log "$ModuleName installed successfully."
        } catch {
            Write-Log "Failed to install $ModuleName. Error: $_"
            Exit 1
        }
    } else {
        Write-Log "$ModuleName module is already available."
    }
}

# Start
Write-Log "Starting script..."
Ensure-Module -ModuleName "ImportExcel"
Ensure-Module -ModuleName "ActiveDirectory"

# Validate Excel file
if (-not (Test-Path $ExcelFilePath)) {
    Write-Log "ERROR: Excel file '$ExcelFilePath' not found. Exiting script."
    Exit 1
}

Write-Log "Reading spreadsheet from $ExcelFilePath..."
$ExcelData = Import-Excel -Path $ExcelFilePath

# Validate required columns
$requiredColumns = "Group Name", "Old OU"
foreach ($column in $requiredColumns) {
    if (-not ($ExcelData | Get-Member -Name $column -MemberType NoteProperty)) {
        Write-Log "ERROR: Required column '$column' is missing from spreadsheet. Exiting."
        Exit 1
    }
}

# Process each group
foreach ($Row in $ExcelData) {
    $GroupName = $Row.'Group Name'
    $OldOU     = $Row.'Old OU'
    $NewOU     = $Row.'New OU' # Optional

    Write-Log "Processing group: $GroupName"

    try {
        $Group = Get-ADGroup -Filter { Name -eq $GroupName } -ErrorAction Stop
        Write-Log "Found group '$GroupName'. Attempting to move to '$OldOU'..."

        Move-ADObject -Identity $Group.DistinguishedName -TargetPath $OldOU -ErrorAction Stop
        Write-Log "Successfully moved '$GroupName' to '$OldOU'."
    } catch {
        Write-Log "ERROR: Failed to move group '$GroupName'. $_"
    }
}

Write-Log "Script execution completed. Log saved to: $LogFilePath"
