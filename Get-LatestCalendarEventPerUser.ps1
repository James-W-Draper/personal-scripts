<#
.SYNOPSIS
Retrieves the most recent calendar event (if any) within a specific date range for each user listed in a text file.

.DESCRIPTION
This script loops through a list of email addresses (e.g. room mailboxes), queries their calendar using Microsoft Graph,
and fetches their latest appointment within the last 6 months to 1 month in the future.

.NOTES
- Requires Microsoft.Graph module
- Requires appropriate app permissions (Calendars.Read / Calendars.Read.Shared)
#>

# === DATE RANGE ===
$StartDate = (Get-Date).AddMonths(-6).ToString("yyyy-MM-ddTHH:mm:ssZ")
$EndDate   = (Get-Date).AddMonths(1).ToString("yyyy-MM-ddTHH:mm:ssZ")

# === FILE INPUT ===
$userListPath = "C:\Scripts\roomlist.txt"  # <-- Update path if needed
$outputPath   = "C:\Scripts\LatestAppointments.csv"

$users = Get-Content $userListPath | Where-Object { $_.Trim() -ne "" }

# === RESULTS ARRAY ===
$results = $users | ForEach-Object {
    $mailbox = $_.Trim()

    try {
        $appointments = Get-MgUserCalendarView -UserId $mailbox -CalendarId "Calendar" -StartDateTime $StartDate -EndDateTime $EndDate

        if ($appointments) {
            $latest = $appointments | Sort-Object { [datetime]$_.Start.DateTime } -Descending | Select-Object -First 1

            [PSCustomObject]@{
                UserId      = $mailbox
                EventStart  = (Get-Date $latest.Start.DateTime).ToString("dd/MM/yyyy")
                Subject     = $latest.Subject
                BodyPreview = $latest.BodyPreview
            }
        } else {
            [PSCustomObject]@{
                UserId      = $mailbox
                EventStart  = "None found"
                Subject     = "No events"
                BodyPreview = ""
            }
        }
    } catch {
        Write-Warning "Failed to query mailbox: $mailbox. Error: $($_.Exception.Message)"
        [PSCustomObject]@{
            UserId      = $mailbox
            EventStart  = "Error"
            Subject     = "Error querying mailbox"
            BodyPreview = $_.Exception.Message
        }
    }
}

# === EXPORT ===
$results | Export-Csv -Path $outputPath -NoTypeInformation
Write-Host "`n✅ Latest calendar entries exported to:`n$outputPath"
