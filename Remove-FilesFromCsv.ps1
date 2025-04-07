<#
.SYNOPSIS
Deletes files listed in a CSV and logs all actions with optional attribute overrides.

.DESCRIPTION
This script reads file paths from a CSV and attempts to delete each file.
It overrides hidden/system/readonly attributes where needed, logs all outcomes,
and exports results to an Excel report and a transcript file.

.NOTES
- Must be run as Administrator
- Requires the ImportExcel module
- CSV must have a column named 'Path'
#>

# === SETTINGS ===
$inputCsvPath      = "C:\Scripts\FileList.csv"  # CSV with 'Path' column
$computerName      = $env:COMPUTERNAME
$currentDate       = Get-Date -Format "yyyy-MM-dd"
$excelFileName     = "DeletionReport_${computerName}_$currentDate.xlsx"
$transcriptFile    = "Transcript_${computerName}_$currentDate.txt"
$outputExcelPath   = "C:\Scripts\$excelFileName"
$transcriptPath    = "C:\Scripts\$transcriptFile"

# === MODULE CHECK ===
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel

# === LOGGING ===
$VerbosePreference = 'Continue'
Start-Transcript -Path $transcriptPath -Force

# === FUNCTION: Override attributes ===
function Get-FileAttributes {
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

# === FUNCTION: Escape characters ===
function Optimize-SpecialCharacters {
    param ([string]$filePath)
    return $filePath -replace '([[]])', '`$1'
}

# === PROCESSING ===
$fileList = Import-Csv -Path $inputCsvPath
$results = @()

foreach ($file in $fileList) {
    $filePath = Optimize-SpecialCharacters -filePath $file.Path.Trim()
    $fileExists = Test-Path -LiteralPath $filePath
    $checkTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $deletionTime = $null
    $deletionSuccess = $false
    $deletionMessage = ""
    $postDeletionCheck = $null
    $ownershipTaken = $false
    $attributeOverrideResult = "Not Checked"

    if ($fileExists) {
        $attributeOverrideResult = Get-FileAttributes -filePath $filePath

        try {
            Remove-Item -LiteralPath $filePath -Force
            $deletionSuccess = $true
            $deletionMessage = "File Deleted Successfully"
            $deletionTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        } catch {
            $deletionMessage = "Error Deleting File: $($_.Exception.Message)"
        }
    } else {
        $deletionMessage = "File Not Found"
    }

    $postDeletionCheck = -not (Test-Path -LiteralPath $filePath)

    $results += [PSCustomObject]@{
        FilePath                = $filePath
        ExistsInitially         = $fileExists
        OwnershipTaken          = $ownershipTaken  # Placeholder (not yet implemented)
        DeletionAttempted       = $deletionTime
        DeletionSuccess         = $deletionSuccess
        DeletionMessage         = $deletionMessage
        FileGonePostDeletion    = $postDeletionCheck
        CheckedOn               = $checkTime
        AttributeOverrideResult = $attributeOverrideResult
    }
}

# === EXPORT REPORT ===
$results | Export-Excel -Path $outputExcelPath -WorksheetName "DeletionReport" -AutoSize
# === EXPORT ERRORS SEPARATELY ===
$errorResults = $results | Where-Object { -not $_.DeletionSuccess }

if ($errorResults.Count -gt 0) {
    $errorReportPath = "C:\Scripts\DeletionErrors_${computerName}_$currentDate.xlsx"
    $errorResults | Export-Excel -Path $errorReportPath -WorksheetName "ErrorsOnly" -AutoSize
    Write-Host "`n⚠️ Errors were found. Error report saved to: $errorReportPath"
} else {
    Write-Host "`n✅ No errors found during file deletion."
}
