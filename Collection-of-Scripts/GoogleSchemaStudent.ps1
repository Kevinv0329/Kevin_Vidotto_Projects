#####

#Purpose: Updates LMS ID, UserIndex ID, and Student ID in custom user info fields located in the users Google profile

#Prerequisites: Gamadv-xtd3 (installed in users Home directory), Powershell, GitHub Desktop (or CLI)

#Process: redacted

#####

# Global Variables
$current_date = Get-Date -Format "MM.dd.yy"
$admin_email = "redacted"
$gam_path = "$HOME/bin/gamadv-xtd3/gam"
$file_id_download = "redacted"
$sheet_id = "0"
$file_id_upload = "redacted"
$targetfolder = "$HOME/Scripts/Temp/WorkingFolder/"
$targetname = "WebhookLink-Student"
$webhook = "$targetfolder" + "$targetname.csv"
$student_data = "$HOME/Scripts/Temp/WorkingFolder/StudentSchema.csv"
$offline_data = "$HOME/Scripts/Github/Systems/OfflineData/OfflineStudentSchema.csv"
$difference = "$HOME/Scripts/Temp/WorkingFolder/Difference_$current_date.csv"


function Start-StudentSchema {

    if (Test-Path -Path $webhook) {

        Remove-Item -Path $webhook
    }

    & $gam_path user $admin_email get drivefile $file_id_download gsheet id:$sheet_id format csv targetfolder $targetfolder targetname $targetname

    if (Test-Path -Path $webhook) {

        $links = Import-Csv -Path $webhook

        $last_link = $links | Select-Object -Last 2 | Select-Object -First 1 | Select-Object -ExpandProperty Link

        $ProgressPreference = 'SilentlyContinue'

        Invoke-WebRequest -Uri "$last_link" -OutFile $student_data

        Start-GoogleUpdate

        Copy-Item -Path $student_data -Destination $offline_data -Force

        $ProgressPreference = 'Continue'

        & $gam_path user $admin_email update drivefile $file_id_upload retainname localfile $student_data addsheet "$current_date"

        $ticket_id = $links | Select-Object -Last 1
        $webhook_url = "redacted" 
        
        $payload = @{
            ticket_id = $ticket_id
        } | ConvertTo-Json -Depth 10
        
        # Try to send the webhook
        try {
            $response = Invoke-RestMethod -Uri $webhook_url -Method Post -Body $payload -ContentType "application/json"
            Write-Host "Webhook sent successfully. Zapier response:"
            Write-Host $response
        } catch {
            Write-Error "Failed to send webhook. Error details: $_"
        }
        
        if (Test-Path -Path $student_data) {
            
            Remove-Item -Path $webhook -Force 
        }
        
    }
    else {
        
        Write-Error "Script did not run"
    }
}

function Start-GoogleUpdate {

    $student_data_import = Import-Csv $student_data
    $offline_data_import = Import-Csv $offline_data

    $diff = Compare-Object -ReferenceObject $offline_data_import -DifferenceObject $student_data_import -Property Email,StudentIndex,UserIndex,LMSID -PassThru | Where-Object { $_.SideIndicator -eq '=>' }

    $diff | Export-Csv -Path $difference

    # Update Student ID using Gam multiprocess
    & $gam_path redirect csv - multiprocess csv $difference gam update user '~Email' StudentData.StudentID '~StudentIndex'

    # Update UserID using Gam multiprocess
    & $gam_path redirect csv - multiprocess csv $difference gam update user '~Email' StudentData.UserID '~UserIndex'

    # Update LmsID using Gam multiprocess
    & $gam_path redirect csv - multiprocess csv $difference gam update user '~Email' StudentData.LmsID '~LMSID'
}

Start-StudentSchema