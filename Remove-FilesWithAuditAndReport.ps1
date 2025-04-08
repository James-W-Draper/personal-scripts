<#
.SYNOPSIS
Deletes files listed in a CSV file and logs the results.

.DESCRIPTION
This script reads file paths from a CSV, checks for existence, handles protected attributes,
attempts deletion, and logs success or errors. It outputs results to an Excel file and creates a transcript log.

.PARAMETER inputCsvPath
Path to the input CSV file with a column named 'Path' containing file paths.

.PARAMETER outputFolder
Directory where the Excel report and transcript will be saved.

.EXAMPLE
.\Remove-FilesWithAuditAndReport.ps1 -inputCsvPath "C:\scripts\files.csv" -outputFolder "C:\scripts"
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$inputCsvPath,

    [Parameter(Mandatory = $true)]
    [string]$outputFolder
)

$VerbosePreference = 'Continue'

# Attempt to load ImportExcel module
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    try {
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
    } catch {
        Write-Host "Failed to install ImportExcel. Please install it manually." -ForegroundColor Red
        return
    }
}

Import-Module ImportExcel -ErrorAction Stop

# Generate dynamic file names
$computerName = $env:COMPUTERNAME
$currentDate = Get-Date -Format "yyyy-MM-dd"
$excelFileName = "DeletionReport_${computerName}_$currentDate.xlsx"
$transcriptFileName = "Transcript_${computerName}_$currentDate.txt"

$outputExcelPath = Join-Path -Path $outputFolder -ChildPath $excelFileName
$transcriptPath = Join-Path -Path $outputFolder -ChildPath $transcriptFileName

# Start transcript
try {
    Start-Transcript -Path $transcriptPath -Force
} catch {
    Write-Warning "Unable to start transcript. Logging will not be saved."
}

function Override-FileAttributes {
    param ([string]$filePath)

    $file = Get-Item -LiteralPath $filePath -Force
    $attributes = $file.Attributes

    if ($attributes -band [System.IO.FileAttributes]::Hidden -or
        $attributes -band [System.IO.FileAttributes]::System -or
        $attributes -band [System.IO.FileAttributes]::ReadOnly) {

        Set-ItemProperty -LiteralPath $filePath -Name Attributes -Value 'Normal'
        return "Overridden"
    }

    return "No Override"
}

function Escape-SpecialCharacters {
    param ([string]$filePath)
    return $filePath -replace '([[]])', '`$1'
}

# Read CSV
$fileList = Import-Csv -Path $inputCsvPath
$results = @()

foreach ($file in $fileList) {
    $escapedPath = Escape-SpecialCharacters -filePath $file.Path
    $fileExists = Test-Path -LiteralPath $escapedPath
    $checkTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $deletionTime = $null
    $deletionSuccess = $false
    $deletionMessage = ""
    $postDeletionCheck = $null
    $ownershipTaken = $false
    $attributeOverrideResult = "Not Checked"

    if ($fileExists) {
        $attributeOverrideResult = Override-FileAttributes -filePath $escapedPath
        try {
            Remove-Item -LiteralPath $escapedPath -Force
            $deletionSuccess = $true
            $deletionMessage = "File Deleted Successfully"
            $deletionTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        } catch {
            $deletionMessage = "Error Deleting File: $_"
        }
    } else {
        $deletionMessage = "File Not Found"
    }

    $postDeletionCheck = -not (Test-Path -LiteralPath $escapedPath)

    $results += [PSCustomObject]@{
        FilePath               = $file.Path
        ExistsInitially        = $fileExists
        OwnershipTaken         = $ownershipTaken
        DeletionAttempted      = $deletionTime
        DeletionSuccess        = $deletionSuccess
        DeletionMessage        = $deletionMessage
        FileGonePostDeletion   = $postDeletionCheck
        CheckedOn              = $checkTime
        AttributeOverrideResult= $attributeOverrideResult
    }
}

# Export results
$results | Export-Excel -Path $outputExcelPath -WorksheetName "DeletionReport" -AutoSize
Write-Host "✅ Report exported to $outputExcelPath"

# Stop transcript
Stop-Transcript
Write-Host "📄 Transcript saved to $transcriptPath"
