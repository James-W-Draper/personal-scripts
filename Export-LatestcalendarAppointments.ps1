$StartDate = (Get-Date).AddMonths(-6)  # Six month in the past
$EndDate = (Get-Date).AddMonths(1)    # One month in the future

# Load the user list from a text file
$users = Get-Content "C:\scripts\roomlist.txt"          # Replace with the actual path to your text file

# Fetch the latest appointment for each user
$results = $users | ForEach-Object {
    $mailbox = $_.Trim() # Ensure no extra whitespace in the email addresses
    try {
        Get-MgUserCalendarView -UserId "'$mailbox'" -CalendarId "Calendar" -StartDateTime $StartDate -EndDateTime $EndDate | 
        Sort-Object {[datetime]$_.Start.DateTime} -Descending | 
        Select-Object -First 1 -Property @{Name='UserId';Expression={$mailbox}}, 
                                            @{Name='EventStart';Expression={ (Get-Date $_.Start.DateTime).ToString("dd/MM/yyyy") }}, 
                                            Subject, 
                                            BodyPreview
    } catch {
        Write-Warning "Failed to query mailbox: $mailbox. Error: $_"
        [PSCustomObject]@{
            UserId      = $mailbox
            EventStart  = "Error"
            Subject     = "Error querying mailbox"
            BodyPreview = $_.Exception.Message
        }
    }
}

# Export the results to a CSV file
$results | Export-Csv "C:\scripts\LatestAppointments.csv" -NoTypeInformation
