# Get a list of all Unified Groups
$groups = Get-UnifiedGroup -resultsize unlimited

# Create an array to store the results
$results = @()

# Iterate through each group and retrieve owner's email addresses
foreach ($group in $groups) {
    $owners = Get-UnifiedGroupLinks -Identity $group.Identity -LinkType Owners
    $ownerEmailAddresses = @()

    foreach ($owner in $owners) {
        $ownerEmailAddress = (Get-Mailbox -Identity $owner).PrimarySmtpAddress
        $ownerEmailAddresses += $ownerEmailAddress
    }

    # Create a custom object to store the group name and owner's email addresses
    $result = [PSCustomObject]@{
        GroupName = $group.DisplayName
        OwnerEmails = ($ownerEmailAddresses -join '; ')
    }

    # Add the result to the array
    $results += $result
}

#Get date for file name  
$FileDate = Get-Date -Format yyyyMMddTHHmmss
#Excel Export Settings
$Common_ExportExcelParams = @{
    BoldTopRow   = $true
    AutoSize     = $true
    AutoFilter   = $true
    FreezeTopRow = $true
}

# Export the results to a CSV file
$results | Export-Excel @Common_ExportExcelParams -Path ("c:\scripts\" + $FileDate + "groupowners_report.xlsx") -WorksheetName report
