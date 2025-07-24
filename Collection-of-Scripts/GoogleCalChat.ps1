#####

#Purpose: Adds new staff to the LN Holiday Calendar and the All Company Google Chat Space

#Prerequisites: PowerShell, Gamadv-xtd3, Github

#Process: redacted

#####

# Global Variables
$CsvPath = "$HOME/Scripts/Temp/WorkingFolder/GoogleCal_Chat.csv"
$GamPath = "$HOME/bin/gamadv-xtd3/gam"
$AdminEmail = "redacted"
$SpaceID = "spaces/AAAA6vmpeSo"
$CalendarEmail = "redacted"
$CalendarName = "LN Holiday Calendar"

# Check if the CSV file exists and remove it if it does
if (Test-Path -Path $CsvPath) {
    Remove-Item -Path $CsvPath -Force
}

# Generate a new CSV file with Google Calendar data
& $GamPath redirect csv $CsvPath ou_and_children_ns /Staff print fields email,creationtime,ou

# Create a chat member for the specified user based on the filtered CSV data
& $GamPath config csv_input_row_filter "'creationTime:date>=-10d'" csv $CsvPath gam user `
$AdminEmail create chatmember $SpaceID user "~primaryEmail"

# Add a calendar for the specified user based on the filtered CSV data
& $GamPath config csv_input_row_filter "'creationTime:date>=-10d'" csv $CsvPath gam user `
"~primaryEmail" add calendar $CalendarEmail selected true summary $CalendarName

# Clean up - delete file created by gam
Remove-Item -Path $CsvPath -Force
