#####

#Purpose: Handles user offboarding task in Asana

#Prerequisites: Gamadv-xtd3 (installed in users Home directory), Powershell, GitHub Desktop (or CLI), Zoom Powershell Wrapper/Module, 

#Process: Redacted

#####

$Gam = "$HOME/bin/gamadv-xtd3/gam"

# Get email input from console
$PreTrimEmail = Read-Host "Enter the email here"
$Email = $PreTrimEmail.Trim()

# This is a saftey check to ensure the input has more than 6 characters
if ($Email.Length -lt 6) {
    Write-Host "Email must have at least 6 characters. Exiting script."
    exit
}

$Confirmation = Read-Host "

Type YES to confirm if the following email is correct.

Email:$Email

"

if ($Confirmation -notlike "YES") {
    Write-Host "Exiting Script"
    Exit
}

function Start-Google {

    # Signout guser
    & $Gam user $Email signout
    #Change password
    & $Gam update user $Email password random
    Write-Output "Set random password"
    #Update recovery email
    & $Gam update user $Email recoveryemail "Redacted"
    Write-Output "Set recovery email"
    #Update recovery phone
    & $Gam update user $Email recoveryphone "Redacted"
    Write-Output "Set recovery phone"
    #Move guser to Inactive OU
    & $Gam user $Email update user ou '/Staff/Z-Inactive Employees'
    Write-Output "$Email has been moved to Suspended OU"
    #Delete all groups from user
    & $Gam user $Email delete groups
    Write-Output "$Email has been removed from all groups"
    #Remove vacation reminder
    & $Gam user $Email vacation off
    Write-Output "Remove OOO emails"

}

function Start-Kandji {

    $KandjiAPI = & /usr/bin/security find-generic-password -a "KandjiAPI" -w

    # This is an API call to Kandji to get the device ID based off the email
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Authorization", "Bearer $KandjiAPI")

    $response = Invoke-RestMethod "https://Redacted.api.kandji.io/api/v1/devices?user_email=$Email&limit=300" -Method 'GET' -Headers $headers
    #$response | ConvertTo-Json

    $deviceId = $response.device_id
    $ID = "$deviceId"

    Write-Host "$Email has the device id of:$deviceId"

    # This is an API call to Kandji to lock the device ID
    $headers2 = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers2.Add("Content-Type", "application/json")
    $headers2.Add("Authorization", "Bearer $KandjiAPI")

    $body2 = @"
    {
    `"Message`": `"This device is locked!`"
    }
"@

    $response2 = Invoke-RestMethod "https://Redacted.api.kandji.io/api/v1/devices/$ID/action/lock" -Method 'POST' -Headers $headers2 -Body $body2
    $response2 | ConvertTo-Json

    Write-Host "$Email computer has been locked"
}

function Start-Dialpad {

    $DialpadAPI = & /usr/bin/security find-generic-password -a "DialpadAPI" -w
    
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Accept", "application/json")
    $headers.Add("Authorization", "Bearer $DialpadAPI")

    $response = Invoke-RestMethod "https://dialpad.com//api/v2/users?email=$Email" -Method 'GET' -Headers $headers

    $DialpadUserID = $response.items[0].id
    
    $response2 = Invoke-RestMethod "https://dialpad.com//api/v2/users/$DialpadUserID" -Method 'DELETE' -Headers $headers
    $response2 | ConvertTo-Json


}

function Start-Zoom {

    $ZoomAPI = & /usr/bin/security find-generic-password -a "ZoomAPI" -w

    Import-module PSZoom
    Connect-PSZoom -AccountID 'Redacted' -ClientID 'Redacted' -ClientSecret $ZoomAPI
    $ZoomID = Get-ZoomUser -EncryptedEmail $Email | Select-Object -ExpandProperty id
    Remove-ZoomUser -UserId "$ZoomID" -Action delete
    
}

Start-Google
Start-Kandji
Start-Dialpad
Start-Zoom