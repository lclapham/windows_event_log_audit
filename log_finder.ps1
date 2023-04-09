# Specify the time zone for Pacific Time
$pacificTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("Pacific Standard Time")

# Get the current date and time in Pacific Time
$now = [System.TimeZoneInfo]::ConvertTimeFromUtc([System.DateTime]::UtcNow, $pacificTimeZone)

# Calculate the date and time 2 weeks ago in Pacific Time
$twoWeeksAgo = $now.AddDays(-14)

# Get the log IDs from the text file---Edit this to the location you want.
$logIDs = Get-Content "F:\Documents\Downloads\winIDs.txt"


foreach ($logID in $logIDs) {
    # Get the events from the System log with the specified log ID in the past 2 weeks
    $events = Get-EventLog -LogName Security -InstanceId $logID -After $twoWeeksAgo

    if ($events -ne $null) {
        # Create a new Excel file with the log ID and current timestamp in the file name
        $excelFileName = "$logID" + "_" + $now.ToString("yyyyMMddTHHmmss") + ".xlsx"
        $excelFilePath = "F:\Documents\Downloads\$excelFileName"
        $excel = New-Object -ComObject Excel.Application
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)

        # Write the header row to the Excel table
        $worksheet.Cells.Item(1,1) = "Time"
        $worksheet.Cells.Item(1,2) = "Event ID"
        $worksheet.Cells.Item(1,3) = "User"
        $worksheet.Cells.Item(1,4) = "IP Address"
        $worksheet.Cells.Item(1,5) = "Working Hours"

        # Loop through the event log entries and write the relevant data to the Excel table
        $row = 2
        foreach ($event in $events) {
            $time = $event.TimeGenerated
            $timeWithKind = [DateTime]::SpecifyKind($time, [DateTimeKind]::Utc)
            $eventTime = [System.TimeZoneInfo]::ConvertTime($timeWithKind, $pacificTimeZone)
            $eventID = $event.InstanceId
            $user = $event.ReplacementStrings[5]
            $ipAddress = $event.ReplacementStrings[19]
            $worksheet.Cells.Item($row,1) = $eventTime
            $worksheet.Cells.Item($row,2) = $eventID
            $worksheet.Cells.Item($row,3) = $user
            $worksheet.Cells.Item($row,4) = $ipAddress

            # Check if the event occurred outside of working hours
            $workdayStart = $eventTime.Date.AddHours(8)    # 8:00 AM Pacific
            $workdayEnd = $eventTime.Date.AddHours(17)    # 5:00 PM Pacific
            if ($eventTime.TimeOfDay -lt $workdayStart.TimeOfDay -or $eventTime.TimeOfDay -gt $workdayEnd.TimeOfDay) {
                # Highlight the row in yellow if the event occurred outside of working hours
                #$range = $worksheet.Range("A$row:D$row")
                #$range.Interior.ColorIndex = 6
                $range = $worksheet.Range("A$row:E$row")
                $range.Interior.ColorIndex = 6
                $worksheet.Cells.Item($row,5) = "No"
            } else {
             $worksheet.Cells.Item($row,5) = "Yes"
                         }

            $row++
        }

        # Auto-fit the columns in the Excel table
        $range = $worksheet.Range("A1:D$row")
        $range.EntireColumn.AutoFit() | Out-Null

        # Save and close the Excel file
        $workbook.SaveAs($excelFilePath)
        $excel.Quit()
    } else {
        Write-Host "No events found for log ID $logID"

        # Auto-fit the columns in the Excel table
        $range = $worksheet.Range("A1:D$row")
        $range.EntireColumn.AutoFit() | Out-Null

        # Create a new Excel file with the log ID and current timestamp in the file name
        $excelFileName = "$logID" + "_" + $now.ToString("yyyyMMddTHHmmss") + ".xlsx"
        $excelFilePath = "F:\Documents\Downloads\$excelFileName"
        $excel = New-Object -ComObject Excel.Application
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)

        # Write the header row to the Excel table
        $worksheet.Cells.Item(1,1) = "Time"
        $worksheet.Cells.Item(1,2) = "Event ID"
        $worksheet.Cells.Item(1,3) = "User"
        $worksheet.Cells.Item(1,4) = "IP Address"
        $worksheet.Cells.Item(1,5) = "Working Hours"

         # Save and close the Excel file
        $workbook.SaveAs($excelFilePath)
        $excel.Quit()
    }
}
