# windows_event_log_audit
This is a PowerShell script that allows you to extract data from windows logs so that you can review them for anomalies. The script uses a file called "winIDs.txt to iterate through the desired windows log IDs that you want to audit. The script reads each ID entry in the “winIDs.txt” file and outputs the results into a Microsoft Excel file label with the IDs that you selected for auditing.  While the output is sparse the idea is to flag events that occurred during non-business hours. The scripts output in Excel highlights events that occur outside of 0800-1700 Pacific so that you can easily see events that may need additional review.

How to use:

1) Update the $logIDs and $excelFilePath paths to the location for the winIDs.txt file and where you want to output the excel files.
2) Enter the Event ID numbers in the winIDs.txt file. Note: this script is tuned for "security" logs so if you want to pull events from other Log update the $events Logname to match log you want to audit. 
3) Run script as admin

This is a simple starting point. Edit script to you needs.