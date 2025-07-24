#####

#Purpose: Updates LMS ID, UserIndex ID, and Teacher ID in custom user info fields located in the users Google profile

#Prerequisites: Gamadv-xtd3 (installed in users Home directory), Powershell, GitHub Desktop (or CLI)

#Process: Redacted

#####

# Global Variables
$current_date = Get-Date -Format "MM.dd.yy"
$admin_email = "redacted"
$gam_path = "$HOME/bin/gamadv-xtd3/gam"
$file_id_download = "redacted"
$sheet_id = "1840163001"
$file_id_upload = "redacted"
$targetfolder = "$HOME/Scripts/Temp/WorkingFolder/"
$targetname = "WebhookLink-Teacher"
$webhook = "$targetfolder" + "$targetname.csv"
$teacher_data = "$HOME/Scripts/Temp/WorkingFolder/TeacherSchema.csv"
$offline_data = "$HOME/Scripts/Github/Systems/OfflineData/OfflineTeacherSchema.csv"

function Start-TeacherSchema {

    if (Test-Path -Path $webhook) {

        Remove-Item -Path $webhook
    }

    & $gam_path user $admin_email get drivefile $file_id_download gsheet id:$sheet_id format csv targetfolder $targetfolder targetname $targetname

    if (Test-Path -Path $webhook) {

        $links = Import-Csv -Path $webhook

        $last_link = $links | Select-Object -Last 2 | Select-Object -First 1 | Select-Object -ExpandProperty Link

        $ProgressPreference = 'SilentlyContinue'

        Invoke-WebRequest -Uri "$last_link" -OutFile $teacher_data

        Copy-Item -Path $teacher_data -Destination $offline_data -Force

        $ProgressPreference = 'Continue'

        & $gam_path user $admin_email update drivefile $file_id_upload retainname localfile $teacher_data addsheet "$current_date"

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
        
        if (Test-Path -Path $teacher_data) {
            
            Remove-Item -Path $webhook -Force 
        }
        
    }
    else {
        
        Write-Error "Script did not run"
    }
}

Start-TeacherSchema