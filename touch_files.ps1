# Get the script's current directory as the default folder path
$defaultFolderPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Ask the user to confirm the default folder path or enter a new one
$folderPath = Read-Host "Press Enter to use the default folder path ($defaultFolderPath) or enter a new folder path"

# If the user does not input anything, use the default folder path
if ([string]::IsNullOrWhiteSpace($folderPath)) {
    $folderPath = $defaultFolderPath
}

# Get the current date and time
$currentDate = Get-Date

# Recursively touch all files in the folder
Get-ChildItem -Path $folderPath -Recurse | ForEach-Object {
    # Update the last write time to the current date and time
    $_.LastWriteTime = $currentDate
}
