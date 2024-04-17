# PowerShell equivalent of the VBA script for archiving emails

# Set the path where emails will be archived
$archivePath = "C:\Users\XXX\OneDrive\Archives emails\" + (Get-Date -Format "yyyy-MM-dd") + "\"

# Create a new Outlook application object
$outlook = New-Object -ComObject Outlook.Application

# Get the MAPI namespace
$namespace = $outlook.GetNameSpace("MAPI")

# Get the default Inbox folder
$inbox = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)

# Function to replace invalid characters in file names
function Replace-InvalidChars {
    param (
        [String]$fileName
    )
    $invalidChars = @("'", "*", "/", "\", ":", "?", "`"", "<", ">", "|")
    foreach ($char in $invalidChars) {
        $fileName = $fileName.Replace($char, "-")
    }
    return $fileName
}

# Function to save emails
#function Save-Emails {
#    param (
#        [Microsoft.Office.Interop.Outlook.Folder]$folder
#    )
#    $items = $folder.Items.Restrict("[Categories] <> 'Archived'")
#    foreach ($item in $items) {
#        if ($item.MessageClass -eq "IPM.Note") {
#            $mail = $item
#            $subject = Replace-InvalidChars -fileName $mail.Subject
#            $receivedTime = $mail.ReceivedTime.ToString("yyyyMMdd-HHmmss")
#            $fileName = $receivedTime + "-" + $subject + ".msg"
#            $mail.SaveAs($archivePath + $fileName, [Microsoft.Office.Interop.Outlook.OlSaveAsType]::olMSG)
#            $mail.Categories = "Archived"
#            $mail.Save()
#        }
#    }
#}

# Function to save emails with size check
function Save-Emails {
    param (
        [Microsoft.Office.Interop.Outlook.Folder]$folder
    )
    $items = $folder.Items.Restrict("[Categories] <> 'Archived'")
    foreach ($item in $items) {
        if ($item.MessageClass -eq "IPM.Note") {
            $mail = $item
            $subject = Replace-InvalidChars -fileName $mail.Subject
            $receivedTime = $mail.ReceivedTime.ToString("yyyyMMdd-HHmmss")
            $fileName = $receivedTime + "-" + $subject + ".msg"
            $filePath = $archivePath + $fileName

            # Check if the file already exists and compare sizes
            if (Test-Path -Path $filePath) {
                $existingFileSize = (Get-Item $filePath).length
                $mailSize = $mail.Size
                # If sizes are different, save with a unique name
                if ($existingFileSize -ne $mailSize) {
                    $uniqueFileName = $receivedTime + "-" + $subject + "-" + (Get-Random) + ".msg"
                    $filePath = $archivePath + $uniqueFileName
                }
            }

            $mail.SaveAs($filePath, [Microsoft.Office.Interop.Outlook.OlSaveAsType]::olMSG)
            $mail.Categories = "Archived"
            $mail.Save()
        }

# Check if the archive path exists, if not, create it
if (-not (Test-Path -Path $archivePath)) {
    New-Item -ItemType Directory -Path $archivePath
}

# Loop through each folder in the Inbox and save emails
foreach ($subfolder in $inbox.Folders) {
    Save-Emails -folder $subfolder
}

# Save emails in the Inbox itself
Save-Emails -folder $inbox

Write-Host "All activities done"
