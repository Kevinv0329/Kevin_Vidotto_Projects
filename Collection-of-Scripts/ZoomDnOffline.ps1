

$ZoomAPI = & /usr/bin/security find-generic-password -a "ZoomAPI" -w

Import-module PSZoom
Connect-PSZoom -AccountID 'XnQnxTAPQvi-uZ0ti0gSpQ' -ClientID 'p2b7bwRJQkSNLnlrB054_A' -ClientSecret $ZoomAPI

$CurrentDate = Get-Date -Format "yyyy-MM-dd-mmss"
$APIExportCsv = "$HOME/Scripts/Temp/WorkingFolder/ZoomApiRawExport_$CurrentDate.csv"
$FinalCsv = "$HOME/Scripts/Github/Systems/OfflineData/Zoom_dn_offline.csv"

# Initialize an array to store all users
$AllUsers = @()

# Loop through all pages
$TotalPages = 30
for ($PageNumber = 1; $PageNumber -le $TotalPages; $PageNumber++) {
    # Make the API request
    $Url = "https://api.zoom.us/v2/users?status=active&page_size=300&page_number=$pageNumber"
    $Response = Invoke-ZoomRestMethod -Uri $Url -Method 'GET' -Headers $Headers
    $Response

    # Add the users from the current page to the array
    $AllUsers += $Response.users
}

$AllUsers | Export-Csv -Path $APIExportCsv

$Rows = Import-csv -Path $APIExportCsv

$DataArray = @()

foreach ($row in $Rows) {
    
    $First = $row.first_name
    $Last = $row.last_name
    $DN = $row.display_name
    $Email = $row.email
    
    if ($Email -like "redacted") {

        $DataArray += [PSCustomObject]@{
            first_name      = $First
            last_name       = $Last
            display_name    = $DN
            email           = $Email

        }
    }
}

$DataArray | Export-Csv -Path $FinalCsv -Force