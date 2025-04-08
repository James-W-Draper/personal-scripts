<#
.SYNOPSIS
Restores or simulates restoring AD objects (Groups, Users, Contacts) to their previous OUs based on Excel input.

.DESCRIPTION
This script reads data from an Excel file containing AD object names and their historical distinguished names (DNs), 
and attempts to move them back to their original Organizational Units (OUs).

It supports:
- Simulation mode (-WhatIf) or real move with -Confirm:$false
- Logging actions to a CSV file (optional)
- Handling of Users, Groups, and Contacts

.PARAMETER ExcelPath
The full path to the Excel file containing columns 'Object Name', 'Object Type', and 'Old Value'.

.PARAMETER PerformMove
If specified, performs the actual move using -Confirm:$false. Otherwise, only simulates using -WhatIf.

.PARAMETER LogPath
Optional path to save a CSV log of actions performed.

.EXAMPLE
.\Restore-ADObjectsFromExcel.ps1 -ExcelPath "C:\Scripts\RestoreList.xlsx" -PerformMove -LogPath "C:\Logs\restore_log.csv"

.NOTES
- Requires ActiveDirectory and ImportExcel modules.
- Must be run with appropriate AD permissions.
#>

param (
    [Parameter(Mandatory)]
    [string]$ExcelPath,

    [switch]$PerformMove,

    [string]$LogPath
)

# Ensure required modules
foreach ($module in @('ImportExcel', 'ActiveDirectory')) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        try {
            Install-Module -Name $module -Scope CurrentUser -Force -ErrorAction Stop
        } catch {
            Write-Host "Failed to install or load module: $module" -ForegroundColor Red
            exit
        }
    }
    Import-Module $module -ErrorAction Stop
}

# Load data from Excel
try {
    $data = Import-Excel -Path $ExcelPath -ErrorAction Stop
} catch {
    Write-Host "Unable to read Excel file at $ExcelPath" -ForegroundColor Red
    exit
}

# Initialize log collection if requested
if ($LogPath) {
    $log = @()
}

foreach ($row in $data) {
    $name  = $row.'Object Name'
    $type  = $row.'Object Type'
    $dn    = $row.'Old Value'

    if (-not $name -or -not $type -or -not $dn) {
        Write-Warning "Skipping row due to missing fields."
        continue
    }

    # Determine target OU from DN
    $ou = ($dn -split ',') | Where-Object { $_ -like 'OU=*' } | ForEach-Object { $_ } | Join-String -Separator ","
    if (-not $ou) {
        Write-Warning "Unable to parse OU from old DN for $name"
        continue
    }

    try {
        $object = switch ($type.ToLower()) {
            'user'    { Get-ADUser    -Identity $name -ErrorAction Stop }
            'group'   { Get-ADGroup   -Identity $name -ErrorAction Stop }
            'contact' { Get-ADObject  -LDAPFilter "(&(objectClass=contact)(cn=$name))" -ErrorAction Stop }
            default   { throw "Unsupported object type: $type" }
        }

        $action = if ($PerformMove) { '-Confirm:$false' } else { '-WhatIf' }
        Write-Host "[$type] $name => $ou" -ForegroundColor Cyan

        if ($PerformMove) {
            Move-ADObject -Identity $object.DistinguishedName -TargetPath $ou -Confirm:$false
        } else {
            Move-ADObject -Identity $object.DistinguishedName -TargetPath $ou -WhatIf
        }

        if ($LogPath) {
            $log += [pscustomobject]@{
                Timestamp = (Get-Date)
                Name      = $name
                Type      = $type
                OU        = $ou
                Action    = if ($PerformMove) { 'Moved' } else { 'Simulated' }
                Status    = 'Success'
            }
        }
    } catch {
        Write-Warning "Failed to process $name ($type): $_"
        if ($LogPath) {
            $log += [pscustomobject]@{
                Timestamp = (Get-Date)
                Name      = $name
                Type      = $type
                OU        = $ou
                Action    = if ($PerformMove) { 'Move' } else { 'Simulate' }
                Status    = $_.Exception.Message
            }
        }
    }
}

# Write log if requested
if ($LogPath -and $log.Count -gt 0) {
    $log | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
    Write-Host "Log written to: $LogPath" -ForegroundColor Green
}
