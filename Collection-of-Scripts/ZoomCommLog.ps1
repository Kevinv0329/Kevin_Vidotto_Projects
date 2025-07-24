#####

#Purpose: Filter a report from Zoom down to students, find the Hosts' ID and email, find the students' ID, find total meeting time for that student

#Prerequisites: Gamadv-xtd3 (installed in users Home directory), Powershell, GitHub Desktop (or CLI)

#Process: Redacted

#####

# Define the current date and time for file naming purposes
$CurrentDate = Get-Date -Format "yyyy-MM-dd-mmss"

# Prompt user for the file name of the downloaded Zoom report
$Downloaded = Read-Host 'Enter File Name Here'

# Set paths for the CSV files
$ZoomReportFilePath = "$HOME/Downloads/$Downloaded"
$StudentIDFromGoogle = "$HOME/Scripts/Github/Systems/OfflineData/OfflineStudentSchema.csv"
$ZoomAPICsv = "$HOME/Scripts/Github/Systems/OfflineData/Zoom_dn_offline.csv"
$TeacherIDFromGoogle = "$HOME/Scripts/Github/Systems/OfflineData/OfflineTeacherSchema.csv"
$ZoomandGoogleMerge = "$HOME/Scripts/Temp/WorkingFolder/ZoomandGoogleMerge-$CurrentDate.csv"
$CommLogImport = "$HOME/Downloads/ZoomCommUpdate-$CurrentDate/CommUpdate-$CurrentDate.csv"
$CommLogImportNull = "$HOME/Downloads/ZoomCommUpdate-$CurrentDate/Null-CommUpdate-$CurrentDate.csv"
$MissingStudentIDLog = "$HOME/Downloads/ZoomCommUpdate-$CurrentDate/MissingStudentIDLog-$CurrentDate.txt"

# Create a new directory and a CSV file for logging communication updates
New-Item -ItemType Directory -Path "$HOME/Downloads/ZoomCommUpdate-$CurrentDate" -Force
New-Item -ItemType File -Path "$CommLogImport" -Force

# List of allowed domains to filter student emails
$AllowedDomains = @("redacted")

# Create a new CSV file to store filtered student data
New-Item -ItemType File -Path "$HOME/Scripts/Temp/WorkingFolder/StudentOnlyFilter-$CurrentDate.csv"
$StudentOnlyFilter = "$HOME/Scripts/Temp/WorkingFolder/StudentOnlyFilter-$CurrentDate.csv"

# Function to filter and export data specific to students based on allowed domains
function Start-StudentOnlyData {

    # Import the Zoom report CSV file
    $FilteredData = Import-Csv -Path $ZoomReportFilePath | Where-Object {
        $UserEmail = $_."User Email"
        $UserJoinStatus = $_."User Join Status"
        $JoinTime = $_."Join Time"
        $LeaveTime = $_."Leave Time"


        # Extract domain from user email
        $StudentEmailDomain = $UserEmail.split('@')[1]

        # Filter based on UserJoinStatus and allowed email domains
        $UserJoinStatus -eq "In Meeting" -and $AllowedDomains -contains $StudentEmailDomain `
        -and -not [string]::IsNullOrEmpty($JoinTime) `
        -and -not [string]::IsNullOrEmpty($LeaveTime)
    }

    # Export the filtered data to the StudentOnlyFilter CSV file without changing the data format
    $FilteredData | Export-Csv -Path $StudentOnlyFilter -NoTypeInformation -Force
}

# Function to check if value is null
function Get-StudentNullValue {

    # Import csv and iterate
    Import-Csv -Path $StudentIDFromGoogle | ForEach-Object {
        $StudentID = [long]$_.'StudentIndex'
        $PrimaryEmail = $_.Email
        
        # Ensure StudentID value exists
        if (-not($StudentID -gt 0)) {
            
            Write-Warning "Missing StudentID Data for $PrimaryEmail"
        }
    }
}

# Function to check if value is null
function Get-TeacherNullValue {

    # Import csv and iterate
    Import-Csv -Path $TeacherIDFromGoogle | ForEach-Object {
        $TeacherID = [long]$_.'UserIndex'
        $PrimaryEmail = $_.Email
        
        # Ensure StudentID value exists
        if (-not($TeacherID -gt 0)) {
            
            Write-Warning "Missing UserID Data for $PrimaryEmail"
        }
    }
}

function Get-ZoomAndGoogleData {

    # Import both CSV files
    $ZoomData = Import-Csv -Path $ZoomAPICsv
    $GoogleData = Import-Csv -Path $TeacherIDFromGoogle

    # Create an array to store the merged data
    $MergedData = @()

    # Iterate through each entry in the ZoomData and find matching emails in the GoogleData
    foreach ($zoomRecord in $ZoomData) {
        # Find the matching row in GoogleData by email
        $googleRecord = $GoogleData | Where-Object { $_.Email -eq $zoomRecord.email }
            
        # If a match is found, create a new object with the required fields
        if ($googleRecord) {

            $MergedData += [PSCustomObject]@{
                Email = $zoomRecord.email
                #ConcatName = $fullName
                Display_Name = $zoomRecord.display_name
                TeacherID = $googleRecord.'UserIndex'
            }
        }
    }

    # Export the merged data to a new CSV file
    $MergedData | Export-Csv -Path $ZoomandGoogleMerge -NoTypeInformation
}

# Function to calculate and display total meeting time for each unique meeting
function Start-TotalMeetingTime {

    # Import the student data with StudentIDs into a hashtable for quick lookups
    $StudentIDLookup = @{}
    Import-Csv -Path $StudentIDFromGoogle | ForEach-Object {

        $StudentIDLookup[$_.Email] = $_.'StudentIndex'
    }

    # Import the teacher data from Get-ZoomAndGoogleData into a hashtable
    $TeacherIDLookup = @{}
    Import-Csv -Path $ZoomandGoogleMerge | ForEach-Object {

        $Display_Name = $_.display_name
        $TeacherIDLookup[$Display_Name] = $_.TeacherID
    }

    # Create an array to store the output records
    $OutputData = @()

    # Create an array to store the output records
    $OutputDataNull = @()

    # Group data by unique combination of UserEmail, MeetingID, Topic, Host, and the parsed DateTime fields
    $Groups = Import-Csv -Path $StudentOnlyFilter | Group-Object -Property {
        $UserEmail = $_."User Email"
        $MeetingID = $_."Meeting ID"
        $Topic = $_.Topic
        $ZoomHost = $_.Host
        $ParsedStartTime = Get-DateTime -DateTimeString $_."Start Time" -UserEmail $UserEmail -MeetingID $MeetingID
        $ParsedEndTime = Get-DateTime -DateTimeString $_."End Time" -UserEmail $UserEmail -MeetingID $MeetingID
        "$UserEmail|$MeetingID|$Topic|$ZoomHost|$ParsedStartTime|$ParsedEndTime"
    }

    # Process each group to calculate total meeting time
    foreach ($group in $Groups) {
        $FirstRecord = $group.Group[0]
        $UserEmail = $FirstRecord."User Email"
        $MeetingID = $FirstRecord."Meeting ID"
        $Topic = $FirstRecord.Topic
        $ZoomHost = $FirstRecord.Host
        $ParsedStartTime = Get-DateTime -DateTimeString $FirstRecord."Start Time" -UserEmail $UserEmail -MeetingID $MeetingID
        $ParsedEndTime = Get-DateTime -DateTimeString $FirstRecord."End Time" -UserEmail $UserEmail -MeetingID $MeetingID
        $StartTime = $ParsedStartTime.ToString("MM/dd/yyyy HH:mm")

        $totalMeetingTime = 0

        # Convert values to DateTime
        foreach ($record in $group.Group) {
            $joinTime = Get-DateTime -DateTimeString $record."Join Time" -UserEmail $UserEmail -MeetingID $MeetingID
            $leaveTime = Get-DateTime -DateTimeString $record."Leave Time" -UserEmail $UserEmail -MeetingID $MeetingID

            # Calculate the duration of the meeting and accumulate the total meeting time
            $meetingDuration = ($leaveTime - $joinTime).TotalMinutes
            $totalMeetingTime += $meetingDuration

        }

        # Look up the StudentID from the hashtable using the email
        $StudentID = $StudentIDLookup[$UserEmail]

        # Initialize an empty array to store log messages
        $LogMessages = @()

        # Your logic to process entries
        if ($totalMeetingTime -gt 4) {
            if ($null -eq $StudentID) {
                # Log the missing StudentID with details
                $LogMessage = "Missing StudentID for email: $UserEmail | Start Time: $StartTime | Topic: $Topic"
                
                # Add the message to the array instead of immediately writing to a file
                $LogMessages += $LogMessage
            }
        }

        # After processing, check if there are any log messages
        if ($LogMessages.Count -gt 0) {
            # Write all log messages to the file
            $LogMessages | Out-File -FilePath $MissingStudentIDLog -Append
        }

        # Check if StudentID is $null and log the necessary information
        if ($totalMeetingTime -gt 4) {
            if ($null -eq $StudentID) {
                # Log the missing StudentID with details
                #Write-Output "Missing StudentID for email: $UserEmail | Start Time: $StartTime | Topic: $Topic"
            }

            # Check if TeacherID is $null and log the necessary information
            if ($null -eq $TeacherID) {
                # Log the missing TeacherID with details
                #Write-Output "Missing TeacherID for Zoom host: $ZoomHost | Start Time: $StartTime | Topic: $Topic"
            }
        }

        # Look up the TeacherID based on the Zoomhost name
        $TeacherID = $null
        if ($TeacherIDLookup.ContainsKey($ZoomHost)) {
            $TeacherID = $TeacherIDLookup[$ZoomHost]
        }

        # Collect the data into a custom object if the total meeting time is greater than 5 minutes
        if ($totalMeetingTime -gt 4) {

            if ($TeacherID -gt 0 -and $StudentID -gt 0){

                # Determine the TwoWay value based on the UserEmail
                $TwoWay = if ($UserEmail -match "co@|az@") {
                    "0"
                } elseif ($UserEmail -match "wa@") {
                    "1"
                } else {
                    "0" 
                }


                $OutputData += [PSCustomObject]@{
                    SentByUserIndex    = $TeacherID
                    StudentIndex       = $StudentID
                    Date               = $StartTime
                    Category           = "Live Session"
                    Subject            = $Topic
                    Contents           = "Host: $ZoomHost<br>Type: Zoom Meeting<br>Student Participation Duration: $totalMeetingTime minutes"
                    DispositionIndex   = ""
                    TwoWay             = "$TwoWay"
                    CampaignIndex      = ""

                
                }
            }

            if ($null -eq $StudentID){

                # Determine the TwoWay value based on the UserEmail
                $TwoWay = if ($UserEmail -match "co@|az@") {
                    "0"
                } elseif ($UserEmail -match "wa@") {
                    "1"
                } else {
                    "0" 
                }


                $OutputDataNull += [PSCustomObject]@{
                    SentByUserIndex    = $TeacherID
                    StudentIndex       = $StudentID
                    Date               = $StartTime
                    Category           = "Live Session"
                    Subject            = $Topic
                    Contents           = "Host: $ZoomHost<br>Type: Zoom Meeting<br>Student Participation Duration: $totalMeetingTime minutes"
                    DispositionIndex   = ""
                    TwoWay             = "$TwoWay"
                    CampaignIndex      = ""
            
                }
            }
        }
    }

    # Export the collected data to the CSV file
    $OutputData | Export-Csv -Path $CommLogImport -NoTypeInformation -Append

    $OutputDataNull | Export-Csv -Path $CommLogImportNull -NoTypeInformation -Append

    Get-FileCheck
}

function Get-FileCheck {

    # Import the CommLogImport CSV file
    $LogData = Import-Csv -Path $CommLogImport

    # Define the headers to check for empty values
    $headersToCheck = @("Date", "Category", "Subject", "Contents")

    # Iterate through each record in the CSV with an index for tracking the row number
    $rowIndex = 1  # CSV rows typically start at 1
    foreach ($record in $LogData) {
        # Check each header for an empty value
        foreach ($header in $headersToCheck) {
            if (-not $record.$header) {
                # Output a warning with the row number and details of the problematic record
                Write-Warning "Missing value for '$header' in record on row $rowIndex : $($record | Out-String)"
            }
        }
        $rowIndex++
    }
}

function Start-FileCleanup {

    if (Test-Path -Path $StudentOnlyFilter) {
        Remove-Item -Path $StudentOnlyFilter -Force
    }

}

function Get-DateTime {
    param (
        [string]$DateTimeString,
        $UserEmail,
        $MeetingID
    )

    # Define date/time formats to attempt parsing
    $formats = @(
        'MM/dd/yyyy HH:mm',  
        'MM/dd/yyyy h:mm tt',
        'MM/dd/yyyy h:mm'    
    )

    # Try parsing the date/time string with each format
    foreach ($format in $formats) {
        $parsedDate = [datetime]::MinValue 
        if ([datetime]::TryParseExact($DateTimeString, $format, $null, [System.Globalization.DateTimeStyles]::None, [ref]$parsedDate)) {
            return $parsedDate
        }
    }

    # Throw an error if parsing fails
    throw "Unable to get date/time: $DateTimeString for $UserEmail in Meeting: $MeetingID"
}

function Start-FileUpload {

    $gam_path = "$HOME/bin/gamadv-xtd3/gam"
    $admin_email = "redacted"
    $current_date = Get-Date -Format "MM.dd.yy--mm.ss"
    $file_id_upload1 = "1fGzhF6d30oovFtTHWgQonHqKvuf--I_CL034o0OnWYw"
    $file_id_upload2 = "11xJ9aD420Ft78vsAtmCpRxhdjZEGlcoU-XAB-RfE9DA"
    $file_id_upload3 = "1XpA_85aTE0XRhZ479Jhqw5sPioj7iYW2WDzwqqsrpf8"
    $file_id_upload4 = "1CuttH0WTqUklmHYaWs5gXy21EmPvo6C_YJVPBdF4Rp8"
    
    & $gam_path user $admin_email update drivefile $file_id_upload1 retainname localfile $ZoomReportFilePath addsheet "$current_date"

    & $gam_path user $admin_email update drivefile $file_id_upload2 retainname localfile $CommLogImport addsheet "$current_date"

    if (Test-Path -Path $MissingStudentIDLog) {

        & $gam_path user $admin_email update drivefile $file_id_upload3 retainname localfile $CommLogImportNull addsheet "$current_date"

        & $gam_path user $admin_email update drivefile $file_id_upload4 retainname localfile $MissingStudentIDLog addsheet "$current_date"
    }
}

function Start-PostNotification {

    do {
        $VerifyUpload = Read-Host "
        -
        -
        -

        Did you complete the Communication Import into Genius? (Type 'Yes' or 'Skip')

        -
        -
        -"
        
        # Check if the input is either 'Yes' or 'Skip'
        if ($VerifyUpload -notmatch '^(Yes|Skip)$') {
            Write-Host "Invalid input. Please type 'Yes' or 'Skip'"
        }

    } while ($VerifyUpload -notmatch '^(Yes|Skip)$')

    if ($VerifyUpload -eq "Yes") {

        Invoke-WebRequest -Uri "redacted" | Out-Null

        Start-Process "redacted"
    }
}


Start-StudentOnlyData
Get-StudentNullValue
Get-TeacherNullValue
Get-ZoomAndGoogleData
Start-TotalMeetingTime
Start-FileCleanup
Start-FileUpload
Start-PostNotification