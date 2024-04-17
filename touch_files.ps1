# Define the path to the OneDrive folder
$folderPath = 'C:\path\to\your\onedrive\folder'

# Get the current date and time
$currentDate = Get-Date

# Recursively touch all files in the folder
Get-ChildItem -Path $folderPath -Recurse | ForEach-Object {
    # Update the last write time to the current date and time
    $_.LastWriteTime = $currentDate
}
