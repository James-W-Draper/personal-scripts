# Install the ImportExcel module if not already installed
# Install-Module -Name ImportExcel -Scope CurrentUser

# Path to the Excel file
$excelFilePath = "C:\Scripts\Group_Object_History.xlsx"

# Import the Excel data
$data = Import-Excel -Path $excelFilePath

# Iterate through each row and simulate reversing changes
foreach ($row in $data) {
    $groupName = $row.'Group Name'
    $oldValue = $row.'Old Value'

    # Extract the old Organizational Unit (OU) from the Old Value
    if ($oldValue -match "CN=(.+?),OU=(.+?)$") {
        $oldOU = $matches[2]

        Write-Host "Simulating restoration of group '$groupName' to OU '$oldOU'..."

        # Simulate the move using -WhatIf
        try {
            Move-ADObject -Identity $groupName -TargetPath $oldOU -WhatIf
        } catch {
            Write-Host "Simulation failed for '$groupName': $_"
        }
    } else {
        Write-Host "Skipping '$groupName' - Unable to parse Old Value."
    }
}
