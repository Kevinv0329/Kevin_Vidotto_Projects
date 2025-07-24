$Gam = "$HOME/bin/gamadv-xtd3/gam"

# Get email input from console
$PreTrimEmail = Read-Host "Enter the original email here"
Write-Host "-----Processing-----"
$PreTrimNewEmail = Read-Host "Enter the NEW email here"

# Removed leading or trailing spaces to ensure there are no "ghost" characters
$Email = $PreTrimEmail.Trim()
$NewEmail = $PreTrimNewEmail.Trim()

# This is a saftey check to ensure the input has more than 6 characters
if ($Email.Length -lt 5 -or $NewEmail.Length -lt 5) {
    Write-Host "Email must have at least 5 characters. Exiting script."
    exit
}

$Confirmation = Read-Host "

Type YES to confirm if the following emails are correct.

Old: $Email

New: $NewEmail

"

if ($Confirmation -notlike "YES") {
    Write-Host "Exiting Script"
    Exit
}

# Get the current date
$CurrentDate = Get-Date -Format "yyyy-MM-dd-mmss"
### Get file path for output of gam schema info
$LocalFile = "$HOME/Scripts/Temp/WorkingFolder/NameChange-$Username-$CurrentDate.csv"
# Get username by removing all characters after the @
$Username = $Email.split('@')[0]

# Gam command to get the schema data
& $Gam config csv_output_header_filter `
"primaryEmail,customSchemas.StudentData.LmsID,customSchemas.StudentData.StudentID,customSchemas.StudentData.UserID" `
redirect csv $LocalFile user $Email print full

$CsvData = Import-Csv -Path $LocalFile
$LmsID = [long]$CsvData.'customSchemas.StudentData.LmsID'
$StudentID = [long]$CsvData.'customSchemas.StudentData.StudentID'
$UserID = [long]$CsvData.'customSchemas.StudentData.UserID'

if ($LmsID -gt 5 -and $StudentID -gt 5 -and $UserID -gt 5) {

    New-Item -ItemType Directory -Path "$HOME/Downloads/$Username-$CurrentDate" -Force
    New-Item -ItemType File -Path "$HOME/Downloads/$Username-$CurrentDate/$Username-UsersUpdate.csv" -Force
    New-Item -ItemType File -Path "$HOME/Downloads/$Username-$CurrentDate/$Username-StudentsUpdate.csv" -Force
    New-Item -ItemType File -Path "$HOME/Downloads/$Username-$CurrentDate/$Username-LmsUpdate.csv" -Force
    Add-Content -Path "$HOME/Downloads/$Username-$CurrentDate/$Username-UsersUpdate.csv" -Value "userindex,Email,Username,GoogleEmail"
    Add-Content -Path "$HOME/Downloads/$Username-$CurrentDate/$Username-StudentsUpdate.csv" -Value "StudentIndex,Email"
    Add-Content -Path "$HOME/Downloads/$Username-$CurrentDate/$Username-LmsUpdate.csv" -Value "Action,userid,username,email"

    Add-Content -Path "$HOME/Downloads/$Username-$CurrentDate/$Username-UsersUpdate.csv" -Value "$UserID,$NewEmail,$NewEmail,$NewEmail"
    Add-Content -Path "$HOME/Downloads/$Username-$CurrentDate/$Username-StudentsUpdate.csv" -Value "$StudentID,$NewEmail"
    Add-Content -Path "$HOME/Downloads/$Username-$CurrentDate/$Username-LmsUpdate.csv" -Value "Edit,$LmsID,$NewEmail,$NewEmail"

    Write-Host "Updating Google Email..."
    & $Gam update user $Email primaryemail $NewEmail

    Write-Host "Opening Genuis and Buzz for Import. View CSV files in your Downloads folder"
    Write-Host "Send Email to redacted"
    Start-Sleep -Seconds 5
    Start-Process "https://portal.redacted.org/AdmBulkImport.aspx"
    Start-Process "https://ssi.redacted.com/admin/29495357/users"

}
else {
    Write-Host "This student is missing Schema data"
    Write-Host "LmsID:$LmsID StudentID:$StudentID UserID:$UserID"
    Exit
}

function Start-Zoom {

    $ZoomAPI = & /usr/bin/security find-generic-password -a "ZoomAPI" -w

    Import-module PSZoom
    Connect-PSZoom -AccountID 'redacted' -ClientID 'redacted' -ClientSecret $ZoomAPI
    $ZoomID = Get-ZoomUser -EncryptedEmail $Email | Select-Object -ExpandProperty id
    Update-ZoomUserEmail -UserId "$ZoomID" -Email "$NewEmail"
    
}

Start-Zoom
