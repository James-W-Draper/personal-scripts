<#
.SYNOPSIS
Safely deletes files listed in a CSV, with auditing, WhatIf simulation, and error reporting.

.DESCRIPTION
This script reads a CSV with file paths and attempts to delete them, logging all actions to Excel.
It overrides restrictive file attributes, optionally takes ownership, and simulates deletions if -WhatIf is used.
Invalid or missing file paths are safely skipped and logged with warnings.

.PARAMETER inputCsvPath
Path to the CSV file containing file paths in a "Path" column.

.PARAMETER outputFolder
Directory where the Excel report and transcript log will be saved.

.PARAMETER WhatIf
Simulate all actions without deleting any files.

.EXAMPLE
.\Remove-FilesWithAudit.ps1 -inputCsvPath "C:\scripts\files.csv" -outputFolder "C:\scripts\logs" -WhatIf
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$inputCsvPath,

    [Parameter(Mandatory = $true)]
    [string]$outputFolder,

    [switch]$WhatIf
)

# ─────────────────────────────────────────────
# 📌 Environment Setup
# ─────────────────────────────────────────────
$VerbosePreference = 'Continue'

$repoName = "PSGallery"
if ((Get-PSRepository -Name $repoName).InstallationPolicy -ne 'Trusted') {
    Set-PSRepository -Name $repoName -InstallationPolicy Trusted
}
Install-Module -Name ImportExcel -Scope CurrentUser -Force

# ─────────────────────────────────────────────
# 📁 Path Preparation
# ─────────────────────────────────────────────
$computerName = $env:COMPUTERNAME
$currentDate = Get-Date -Format "yyyy-MM-dd"
$excelFileName = "DeletionReport_${computerName}_$currentDate.xlsx"
$transcriptFileName = "Transcript_${computerName}_$currentDate.txt"
$outputExcelPath = Join-Path $outputFolder $excelFileName
$transcriptPath = Join-Path $outputFolder $transcriptFileName

if (-not (Test-Path -LiteralPath $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder | Out-Null
}

Start-Transcript -Path $transcriptPath -Force

# ─────────────────────────────────────────────
# 🔧 Utility Functions
# ─────────────────────────────────────────────
function Escape-SpecialCharacters {
    param ([string]$filePath)
    return $filePath -replace '([[]])', '`$1'
}

function Override-FileAttributes {
    param ([string]$filePath)
    try {
        $file = Get-Item -LiteralPath $filePath -Force
        $attributes = $file.Attributes
        $requiresChange = $attributes -band ([System.IO.FileAttributes]::Hidden -bor `
                                             [System.IO.FileAttributes]::System -bor `
                                             [System.IO.FileAttributes]::ReadOnly)
        if ($requiresChange) {
            Set-ItemProperty -LiteralPath $filePath -Name Attributes -Value 'Normal'
            return "Overridden"
        }
        return "No Override"
    } catch {
        return "Attribute Override Failed: $_"
    }
}

function Take-FileOwnership {
    param ([string]$filePath)
    try {
        takeown.exe /F "`"$filePath`"" /A /R /D Y | Out-Null
        icacls.exe "`"$filePath`"" /grant "$env:USERNAME:F" /T /C | Out-Null
        return $true
    } catch {
        return $false
    }
}

# ─────────────────────────────────────────────
# 📊 Main Logic
# ─────────────────────────────────────────────
$fileList = Import-Csv -Path $inputCsvPath
$results = @()
$errorLog = @()

foreach ($file in $fileList) {

    if (-not $file.Path -or [string]::IsNullOrWhiteSpace($file.Path)) {
        Write-Warning "⚠️ Skipping row with missing or empty path."
        continue
    }

    $filePath = Escape-SpecialCharacters -filePath $file.Path
    $fileExists = Test-Path -LiteralPath $filePath
    $checkTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $deletionTime = $null
    $deletionSuccess = $false
    $deletionMessage = ""
    $postDeletionCheck = $null
    $ownershipTaken = $false
    $attributeOverrideResult = "Not Checked"

    if ($fileExists) {
        $attributeOverrideResult = Override-FileAttributes -filePath $filePath

        if ($WhatIf) {
            $ownershipTaken = $false
            $deletionMessage = "Simulated Deletion (WhatIf)"
        } else {
            $ownershipTaken = Take-FileOwnership -filePath $filePath
            try {
                Remove-Item -LiteralPath $filePath -Force
                $deletionSuccess = $true
                $deletionMessage = "File Deleted Successfully"
                $deletionTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            } catch {
                $deletionMessage = "Error Deleting File: $_"
                $errorLog += [PSCustomObject]@{
                    FilePath     = $filePath
                    TimeChecked  = $checkTime
                    ErrorMessage = $_.Exception.Message
                }
            }
        }
    } else {
        $deletionMessage = "File Not Found"
        $errorLog += [PSCustomObject]@{
            FilePath     = $filePath
            TimeChecked  = $checkTime
            ErrorMessage = "File Not Found"
        }
    }

    $postDeletionCheck = if ($WhatIf) { $null } else { -not (Test-Path -LiteralPath $filePath) }

    $results += [PSCustomObject]@{
        FilePath                = $filePath
        ExistsInitially         = $fileExists
        OwnershipTaken          = $ownershipTaken
        DeletionAttempted       = $deletionTime
        DeletionSuccess         = $deletionSuccess
        DeletionMessage         = $deletionMessage
        FileGonePostDeletion    = $postDeletionCheck
        CheckedOn               = $checkTime
        AttributeOverrideResult = $attributeOverrideResult
    }
}

# ─────────────────────────────────────────────
# 📤 Export Results
# ─────────────────────────────────────────────
$results | Export-Excel -Path $outputExcelPath -WorksheetName "DeletionReport" -AutoSize

if ($errorLog.Count -gt 0) {
    $errorLog | Export-Excel -Path $outputExcelPath -WorksheetName "Errors" -AutoSize
}

# ─────────────────────────────────────────────
# ✅ Completion Message
# ─────────────────────────────────────────────
if ($WhatIf) {
    Write-Host "`n⚠️ WhatIf mode enabled. No files were actually deleted." -ForegroundColor Yellow
}

Write-Host "`n✅ Script complete. Results saved to:`n$outputExcelPath"
Stop-Transcript
Write-Host "📝 Transcript saved to:`n$transcriptPath"
