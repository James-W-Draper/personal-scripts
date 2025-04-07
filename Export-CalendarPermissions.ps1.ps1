<#
=============================================================================================
.SYNOPSIS
    Generates detailed calendar permission reports for Office 365 mailboxes.

.DESCRIPTION
    This script connects to Exchange Online and exports mailbox calendar permissions, including default, external,
    and user-shared access. It supports optional filters and export file naming that includes the report type,
    target user (if applicable), and current date for easy reference.

.PARAMETER UserName
    Optional. The username of an account to authenticate with Exchange Online. Use with -Password.

.PARAMETER Password
    Optional. The plaintext password of the account. Use with -UserName.

.PARAMETER ShowAllPermissions
    If specified, shows all calendar permissions for all user mailboxes.

.PARAMETER DisplayAllCalendarsSharedTo
    Shows only calendars that are shared to a specific user (UPN/email).

.PARAMETER DefaultCalendarPermissions
    Shows only default user calendar permissions.

.PARAMETER ExternalUsersCalendarPermissions
    Shows only permissions shared externally (ExchangePublishedUser).

.PARAMETER CSVIdentityFile
    Optional. Path to a CSV file containing mailbox identities to report on.

.EXAMPLE
    .\Export-CalendarPermissions.ps1 -ShowAllPermissions

.EXAMPLE
    .\Export-CalendarPermissions.ps1 -DisplayAllCalendarsSharedTo "jane.doe@contoso.com"

.EXAMPLE
    .\Export-CalendarPermissions.ps1 -ExternalUsersCalendarPermissions -CSVIdentityFile "C:\Scripts\MailboxList.csv"
=============================================================================================
#>

param (
    [string] $UserName = $null,
    [string] $Password = $null,
    [Switch] $ShowAllPermissions,
    [String] $DisplayAllCalendarsSharedTo,
    [Switch] $DefaultCalendarPermissions,
    [Switch] $ExternalUsersCalendarPermissions,
    [String] $CSVIdentityFile    
)

$global:ExportPath = "C:\Scripts"
if (-not (Test-Path $global:ExportPath)) {
    New-Item -Path $global:ExportPath -ItemType Directory -Force | Out-Null
}

Function Connect_Exo {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host "Installing Exchange Online PowerShell module..."
        Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
    }
    Import-Module ExchangeOnlineManagement
    Write-Host "Connecting to Exchange Online..."
    if (($UserName -ne "") -and ($Password -ne "")) {
        $SecurePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
        $Credential = New-Object System.Management.Automation.PSCredential($UserName, $SecurePassword)
        Connect-ExchangeOnline -Credential $Credential
    } else {
        Connect-ExchangeOnline
    }
}

Function OutputFile_Declaration {
    $dateTag = (Get-Date -Format "yyyyMMdd")
    if ($DisplayAllCalendarsSharedTo) {
        $userTag = $DisplayAllCalendarsSharedTo -replace '[^a-zA-Z0-9]', '_'
        $global:ExportCSVFileName = "CalendarPermissions_DisplayAllCalendarsSharedTo_${userTag}_$dateTag.csv"
    }
    elseif ($ShowAllPermissions) {
        $global:ExportCSVFileName = "CalendarPermissions_AllPermissions_$dateTag.csv"
    }
    elseif ($DefaultCalendarPermissions) {
        $global:ExportCSVFileName = "CalendarPermissions_DefaultUsers_$dateTag.csv"
    }
    elseif ($ExternalUsersCalendarPermissions) {
        $global:ExportCSVFileName = "CalendarPermissions_ExternalUsers_$dateTag.csv"
    }
    else {
        $global:ExportCSVFileName = "CalendarPermissions_Report_$dateTag.csv"
    }
    $global:ExportCSVFileName = Join-Path -Path $global:ExportPath -ChildPath $global:ExportCSVFileName
}

Function RetrieveMBs {
    if ($CSVIdentityFile) {
        $IdentityList = Import-Csv -Header "IdentityValue" $CSVIdentityFile
        foreach ($Identity in $IdentityList) {
            $CurrUserData = Get-Mailbox -Identity $Identity.IdentityValue -ErrorAction SilentlyContinue
            if ($CurrUserData) { GetCalendars } else {
                Write-Host "$($Identity.IdentityValue) not found."
            }
        }
    } else {
        Get-Mailbox -ResultSize Unlimited | ForEach-Object {
            $CurrUserData = $_
            GetCalendars
        }
    }
}

Function GetCalendars {
    $global:MailboxCount++
    $EmailAddress = $CurrUserData.PrimarySmtpAddress
    $CalendarFolders = @()
    $CalendarStats = Get-MailboxFolderStatistics -Identity $EmailAddress -FolderScope Calendar
    foreach ($folder in $CalendarStats) {
        $folderPath = if ($folder.FolderType -eq "Calendar") { "$EmailAddress:\Calendar" } else { "$EmailAddress:\Calendar\$($folder.Name)" }
        $CalendarFolders += $folderPath
    }
    RetrieveCalendarPermissions
}

Function RetrieveCalendarPermissions {
    if ($DisplayAllCalendarsSharedTo) {
        $Flag = "DisplayAllCalendarsSharedTo"
        foreach ($CalendarFolder in $CalendarFolders) {
            $CalendarName = ($CalendarFolder -split "\\")[-1]
            $CurrCalendarData = Get-MailboxFolderPermission -Identity $CalendarFolder -User $DisplayAllCalendarsSharedTo -ErrorAction SilentlyContinue
            if ($CurrCalendarData) { SaveCalendarPermissionsData }
        }
    } elseif ($ShowAllPermissions) {
        foreach ($CalendarFolder in $CalendarFolders) {
            $CalendarName = ($CalendarFolder -split "\\")[-1]
            Get-MailboxFolderPermission -Identity $CalendarFolder | ForEach-Object {
                $CurrCalendarData = $_
                SaveCalendarPermissionsData
            }
        }
    } elseif ($DefaultCalendarPermissions) {
        $Flag = "DefaultUserCalendar"
        foreach ($CalendarFolder in $CalendarFolders) {
            $CalendarName = ($CalendarFolder -split "\\")[-1]
            $CurrCalendarData = Get-MailboxFolderPermission -Identity $CalendarFolder | Where-Object { $_.User -eq "Default" }
            if ($CurrCalendarData) { SaveCalendarPermissionsData }
        }
    } elseif ($ExternalUsersCalendarPermissions) {
        $Flag = "ExternalUserCalendarSharing"
        foreach ($CalendarFolder in $CalendarFolders) {
            $CalendarName = ($CalendarFolder -split "\\")[-1]
            Get-MailboxFolderPermission -Identity $CalendarFolder | Where-Object { $_.User.DisplayName -like "ExchangePublishedUser.*" } | ForEach-Object {
                $CurrCalendarData = $_
                SaveCalendarPermissionsData
            }
        }
    } else {
        foreach ($CalendarFolder in $CalendarFolders) {
            $CalendarName = ($CalendarFolder -split "\\")[-1]
            Get-MailboxFolderPermission -Identity $CalendarFolder | Where-Object { $_.User -ne "Default" -and $_.User -ne "Anonymous" } | ForEach-Object {
                $CurrCalendarData = $_
                SaveCalendarPermissionsData
            }
        }
    }
}

Function SaveCalendarPermissionsData {
    $SharedToMB = $CurrCalendarData.User.DisplayName
    $AllowedUser = if ($SharedToMB -like "ExchangePublishedUser.*") {
        $SharedToMB -replace "ExchangePublishedUser.", ""
    } else { $SharedToMB }
    $UserType = if ($SharedToMB -like "ExchangePublishedUser.*") { "External/Unauthorized" } else { "Member" }
    $AccessRights = $CurrCalendarData.AccessRights -join ","
    $PermissionFlag = if ($CurrCalendarData.SharingPermissionFlags) { $CurrCalendarData.SharingPermissionFlags -join "," } else { "-" }

    $ExportResult = [PSCustomObject]@{
        'Mailbox Name'             = $CurrUserData.Identity
        'Email Address'            = $CurrUserData.PrimarySmtpAddress
        'Mailbox Type'             = $CurrUserData.RecipientTypeDetails
        'Calendar Name'            = $CalendarName
        'Shared To'                = $AllowedUser
        'User Type'                = $UserType
        'Access Rights'            = $AccessRights
        'Sharing Permission Flags' = $PermissionFlag
    }

    $SelectFields = switch ($Flag) {
        "DisplayAllCalendarsSharedTo" { 'Mailbox Name','Email Address','Calendar Name','Access Rights','Sharing Permission Flags','Mailbox Type' }
        "DefaultUserCalendar"         { 'Mailbox Name','Email Address','Mailbox Type','Calendar Name','Access Rights' }
        "ExternalUserCalendarSharing" { 'Mailbox Name','Email Address','Calendar Name','Shared To','Access Rights' }
        default                        { 'Mailbox Name','Email Address','Mailbox Type','Calendar Name','Shared To','Access Rights','Sharing Permission Flags','User Type' }
    }
    $ExportResult | Select-Object $SelectFields | Export-Csv -Path $global:ExportCSVFileName -NoTypeInformation -Append
    $global:ReportSize++
}

# Main Execution
Connect_Exo
OutputFile_Declaration
$global:MailboxCount = 0
$global:ReportSize = 0

if ($DisplayAllCalendarsSharedTo) {
    $CurrUserData = Get-Mailbox -Identity $DisplayAllCalendarsSharedTo -ErrorAction SilentlyContinue
    if (-not $CurrUserData) {
        Write-Host "Invalid user: $DisplayAllCalendarsSharedTo" -ForegroundColor Red
        exit
    }
    GetCalendars
} else {
    Write-Host "Generating calendar permission report..."
    RetrieveMBs
}

if (Test-Path -Path $global:ExportCSVFileName) {
    Write-Host "\n✅ Exported $global:ReportSize record(s) to: $global:ExportCSVFileName" -ForegroundColor Green
    Invoke-Item $global:ExportCSVFileName
} else {
    Write-Host "\nNo data found matching your criteria." -ForegroundColor Yellow
}

Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore
