<#
.SYNOPSIS
    Manage Microsoft 365 licenses using Microsoft Graph PowerShell.

.DESCRIPTION
    This script provides a unified, menu-driven interface for reporting and managing Microsoft 365 licenses using Microsoft Graph.
    It supports license reports, assignment, removal, cleanup, and more across licensed and unlicensed users.

.EXAMPLE
    .\Manage-M365Licenses.ps1
    Interactively prompts for action and performs license management/reporting.

.EXAMPLE
    .\Manage-M365Licenses.ps1 -Action 1
    Automatically runs the "Get all licensed users" report.

.NOTES
    Author: O365Reports Team (Enhanced by ChatGPT)
    Version: 4.0
    Requires: Microsoft.Graph module
#>

param(
    [string]$LicenseName,
    [string]$LicenseUsageLocation,
    [int]$Action,
    [switch]$MultipleActionsMode
)

function Ensure-ModuleInstalled {
    param([string]$ModuleName)
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "$ModuleName not found. Installing..." -ForegroundColor Yellow
        Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber
    }
}

function Connect-ToGraph {
    Try {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes "Directory.ReadWrite.All" -ErrorAction Stop
        Set-MgProfile -Name "beta"
        Write-Host "Connected successfully." -ForegroundColor Green
    } catch {
        Write-Host "Failed to connect to Microsoft Graph: $_" -ForegroundColor Red
        Exit
    }
}

function Load-LicenseFriendlyNames {
    $filePath = ".\LicenseFriendlyName.txt"
    if (-not (Test-Path $filePath)) {
        Write-Warning "LicenseFriendlyName.txt not found. Defaulting to raw license IDs."
        return @{}
    }
    return Get-Content -Raw -Path $filePath | ConvertFrom-StringData
}

function Get-LicenseMappings {
    $skuMap = @{}
    $skuIdMap = @{}
    Get-MgSubscribedSku | ForEach-Object {
        $skuMap[$_.SkuPartNumber] = $_.SkuId
        $skuIdMap[$_.SkuId] = $_.SkuPartNumber
    }
    return @{
        SkuPartToId = $skuMap
        SkuIdToPart = $skuIdMap
    }
}

function Export-CSVFile {
    param(
        [string]$Path,
        [object]$Object
    )
    $Object | Export-Csv -Path $Path -NoTypeInformation -Append
}

function Export-LicensedUsers {
    param(
        [string]$Path,
        [hashtable]$FriendlyNames,
        [hashtable]$SkuIdToPart
    )
    Get-MgUser -All | Where-Object { $_.AssignedLicenses.Count -gt 0 } | ForEach-Object {
        $friendlyList = $_.AssignedLicenses.SkuId | ForEach-Object {
            $part = $SkuIdToPart[$_]
            if ($FriendlyNames.ContainsKey($part)) { $FriendlyNames[$part] } else { $part }
        }
        $obj = [pscustomobject]@{
            DisplayName = $_.DisplayName
            UPN = $_.UserPrincipalName
            LicensePlans = ($_.AssignedLicenses.SkuId -join ", ")
            FriendlyNames = ($friendlyList -join ", ")
            AccountStatus = if ($_.AccountEnabled) { "Enabled" } else { "Disabled" }
            Department = $_.Department
            JobTitle = $_.JobTitle
        }
        Export-CSVFile -Path $Path -Object $obj
    }
}

function Export-UnlicensedUsers {
    param(
        [string]$Path
    )
    Get-MgUser -All | Where-Object { $_.AssignedLicenses.Count -eq 0 } | ForEach-Object {
        $obj = [pscustomobject]@{
            DisplayName = $_.DisplayName
            UPN = $_.UserPrincipalName
            AccountStatus = if ($_.AccountEnabled) { "Enabled" } else { "Disabled" }
            Department = $_.Department
            JobTitle = $_.JobTitle
        }
        Export-CSVFile -Path $Path -Object $obj
    }
}

function Show-Menu {
    Write-Host "\nMicrosoft 365 License Management Menu" -ForegroundColor Cyan
    Write-Host "1. Export all licensed users"
    Write-Host "2. Export all unlicensed users"
    Write-Host "0. Exit"
}

function Main {
    Ensure-ModuleInstalled -ModuleName "Microsoft.Graph"
    Connect-ToGraph

    $friendlyNames = Load-LicenseFriendlyNames
    $maps = Get-LicenseMappings
    $SkuPartToId = $maps.SkuPartToId
    $SkuIdToPart = $maps.SkuIdToPart

    $OutputDir = ".\Reports"
    if (-not (Test-Path $OutputDir)) { New-Item -Path $OutputDir -ItemType Directory | Out-Null }

    do {
        Show-Menu
        if ($Action -eq $null) {
            $selection = Read-Host "Please select an option"
        } else {
            $selection = $Action
        }

        switch ($selection) {
            1 {
                $csv = Join-Path $OutputDir ("LicensedUsersReport_" + (Get-Date -Format yyyyMMdd_HHmm) + ".csv")
                Write-Host "Generating Licensed Users Report..." -ForegroundColor Cyan
                Export-LicensedUsers -Path $csv -FriendlyNames $friendlyNames -SkuIdToPart $SkuIdToPart
                Write-Host "Report saved to: $csv" -ForegroundColor Green
            }
            2 {
                $csv = Join-Path $OutputDir ("UnlicensedUsersReport_" + (Get-Date -Format yyyyMMdd_HHmm) + ".csv")
                Write-Host "Generating Unlicensed Users Report..." -ForegroundColor Cyan
                Export-UnlicensedUsers -Path $csv
                Write-Host "Report saved to: $csv" -ForegroundColor Green
            }
            0 {
                Write-Host "Exiting script..." -ForegroundColor Yellow
            }
            default {
                Write-Warning "Invalid selection. Please choose a valid option."
            }
        }

        if ($Action -ne $null -or $selection -eq 0) { break }

    } while ($true)

    Disconnect-MgGraph | Out-Null
    Write-Host "Disconnected from Microsoft Graph." -ForegroundColor DarkGray
}

Main
