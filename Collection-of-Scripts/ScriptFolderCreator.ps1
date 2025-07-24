#####

#Purpose: Creates certain folders in the users directory used by multiple other scripts

#Prerequisites: Powershell, GitHub

#Process: redacted

#####


# Define the base path using the user's home directory
$UserFolder = "$HOME/Scripts"

# Directories to create
$Directories = @(
    "$UserFolder",
    "$UserFolder/Github"
    "$UserFolder/Temp",
    "$UserFolder/Temp/WorkingFolder",
    "$UserFolder/Temp/TrashFiles"
)

# Loop through and create each directory if it does not already exist
foreach ($Directory in $Directories) {
    if (-Not (Test-Path -Path $Directory)) {
        New-Item -Path $Directory -ItemType Directory
        Write-Host "Created directory: $Directory"
    } else {
        Write-Host "Directory already exists: $Directory"
    }
}
