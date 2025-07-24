#####

#Purpose: Handles "Email Domain Affiliation Mismatch" tickets from Freshdesk 

#Prerequisites: Gamadv-xtd3 (installed in users Home directory), Powershell, GitHub Desktop (or CLI), Zoom Powershell Wrapper/Module

#Process: Redacted

#####


$GamPath = "$HOME/bin/gamadv-xtd3/gam"

$CurrentDate = Get-Date -Format "yyyy-MM-dd-mmss"

$Downloaded = Read-Host 'Enter File Name Here'
$FilePath = "$HOME/Downloads/$Downloaded"


function Start-Gam {

    Import-Csv -Path $FilePath | ForEach-Object {

        $NewDomain = $_.'Affiliation Domain'
        $Email = $_.UserName
        $EmailSplit = $Email.split('@')[0]
        $EmailSubstring = $EmailSplit.Substring(0, $EmailSplit.Length - 2)
        $NewEmail = $EmailSubstring + "$NewDomain"
    
        & $GamPath update user $Email primaryemail $NewEmail
    }

}

function Start-Genius {

    New-Item -ItemType Directory -Path "$HOME/Downloads/BulkDomainUpdate-$CurrentDate" -Force
    New-Item -ItemType File -Path "$HOME/Downloads/BulkDomainUpdate-$CurrentDate/Bulk-UsersUpdate.csv" -Force
    New-Item -ItemType File -Path "$HOME/Downloads/BulkDomainUpdate-$CurrentDate/Bulk-StudentsUpdate.csv" -Force
    New-Item -ItemType File -Path "$HOME/Downloads/BulkDomainUpdate-$CurrentDate/Bulk-LmsUpdate.csv" -Force

    Add-Content -Path "$HOME/Downloads/BulkDomainUpdate-$CurrentDate/Bulk-UsersUpdate.csv" -Value "userindex,Email,Username,GoogleEmail"
    Add-Content -Path "$HOME/Downloads/BulkDomainUpdate-$CurrentDate/Bulk-StudentsUpdate.csv" -Value "StudentIndex,Email"
    Add-Content -Path "$HOME/Downloads/BulkDomainUpdate-$CurrentDate/Bulk-LmsUpdate.csv" -Value "Action,userid,username,email"

    Import-Csv -Path $FilePath | ForEach-Object {

        $LmsID = $_.LMSID
        $StudentID = $_.StudentIndex
        $UserID = $_.userindex
        $NewDomain = $_.'Affiliation Domain'
        $Email = $_.UserName
        $EmailSplit = $Email.split('@')[0]
        $EmailSubstring = $EmailSplit.Substring(0, $EmailSplit.Length - 2)
        $NewEmail = $EmailSubstring + $NewDomain

        Add-Content -Path "$HOME/Downloads/BulkDomainUpdate-$CurrentDate/Bulk-UsersUpdate.csv" -Value "$UserID,$NewEmail,$NewEmail,$NewEmail"
        Add-Content -Path "$HOME/Downloads/BulkDomainUpdate-$CurrentDate/Bulk-StudentsUpdate.csv" -Value "$StudentID,$NewEmail"
        Add-Content -Path "$HOME/Downloads/BulkDomainUpdate-$CurrentDate/Bulk-LmsUpdate.csv" -Value "Edit,$LmsID,$NewEmail,$NewEmail"
    }

    Start-Process "https://portal.myschool.org/AdmBulkImport.aspx"
    Start-Process "https://ssi.agilixbuzz.com/admin/29495357/users"

}

function Start-Zoom {

    Import-module PSZoom

    $ZoomAPI = & /usr/bin/security find-generic-password -a "ZoomAPI" -w

    Connect-PSZoom -AccountID 'Redacted' -ClientID 'Redacted' -ClientSecret $ZoomAPI

    
    Import-Csv -Path $FilePath | ForEach-Object {

        $NewDomain = $_.'Affiliation Domain'
        $Email = $_.UserName
        $EmailSplit = $Email.split('@')[0]
        $EmailSubstring = $EmailSplit.Substring(0, $EmailSplit.Length - 2)
        $NewEmail = $EmailSubstring + "$NewDomain"

        $ZoomID = Get-ZoomUser -EncryptedEmail $Email -ErrorAction SilentlyContinue | Select-Object -ExpandProperty id
        Update-ZoomUserEmail -UserId "$ZoomID" -Email "$NewEmail" -ErrorAction SilentlyContinue
    }
}

Start-Gam
Start-Genius
Start-Zoom

