' Version 1.0 Automatic Email Archive

Dim oMail As Outlook.MailItem
Dim objItem As Object
Dim sPath As String
Dim dtDate As Date
Dim sName As String
Dim YesNo As Long
Dim ToArchive As Long
Dim i As Long
Dim ns As Outlook.NameSpace
Dim myfolder As Outlook.Folder
Dim mysubfolder As Outlook.Folder
Dim Folderlist As String
Dim Foldercount As Long
Dim Duration As Long
Dim ExitMacro As Boolean
Dim Mailscount As Long
Dim NoFurtherQuestions As Boolean

Option Explicit

Public Sub SaveMessageAsMsg()

  ' Get the OneDrive path using the environment variable
  sPath = Environ("ONEDRIVE") & "\Archives emails\" & Format$(Date, "yyyy-mm-dd") & "\"

  ' Debug: Display the calculated path 
  Debug.Print "Archive Path: " & sPath

  ' Get the default inbox folder
  Set ns = Application.GetNamespace("MAPI")
  Set myfolder = ns.GetDefaultFolder(olFolderInbox)

  ' Get the number of unarchived emails
  ToArchive = 0
  For Each mysubfolder In myfolder.Folders
    If mysubfolder.Items.Restrict("[Categories] <> 'Archived'").Count > 0 Then
      For Each objItem In mysubfolder.Items.Restrict("[Categories] <> 'Archived'")
        If objItem.MessageClass = "IPM.Note" Then ToArchive = ToArchive + 1
      Next
    End If
    Mailscount = Mailscount + mysubfolder.Items.Count
  Next

  Set mysubfolder = myfolder
  If mysubfolder.Items.Restrict("[Categories] <> 'Archived'").Count > 0 Then
    For Each objItem In mysubfolder.Items.Restrict("[Categories] <> 'Archived'")
      If objItem.MessageClass = "IPM.Note" Then ToArchive = ToArchive + 1
    Next
  End If

  ' Confirm archive with option to disable further questions
  YesNo = MsgBox("There are " & ToArchive & " emails to archive out of " & Mailscount & " mails in Outlook. Proceed? " & vbCrLf & _
                 "**Clicking 'Yes' will archive all emails without further confirmation.**", vbYesNoCancel, "Archival")

  If YesNo = vbCancel Then Exit Sub 
  If YesNo = vbYes Then NoFurtherQuestions = True 

  ' Create new folder if doesn't exist
  If Dir(sPath, vbDirectory) = "" Then
    ' Create folder using shell command
    Shell ("cmd /c mkdir """ & sPath & """")
  End If

  ' Loop through each folder and archive emails
  For Each mysubfolder In myfolder.Folders
    If Not SaveEmails(mysubfolder, NoFurtherQuestions) Then Exit Sub ' Exit if SaveEmails returns False (user canceled)
  Next

  If ExitMacro = True Then Exit Sub

  ' Loop in all folders (redundant with previous loop)
  Set mysubfolder = myfolder
  If Not SaveEmails(mysubfolder, NoFurtherQuestions) Then Exit Sub 

  i = MsgBox("All activities done", vbOKOnly)

End Sub

Private Sub ReplaceCharsForFileName(sName As String, _
  sChr As String _
)
  sName = Replace(sName, "'", sChr)
  sName = Replace(sName, "*", sChr)
  sName = Replace(sName, "/", sChr)
  sName = Replace(sName, "\", sChr)
  sName = Replace(sName, ":", sChr)
  sName = Replace(sName, "?", sChr)
  sName = Replace(sName, Chr(34), sChr)
  sName = Replace(sName, "<", sChr)
  sName = Replace(sName, ">", sChr)
  sName = Replace(sName, "|", sChr)
End Sub

Function SaveEmails(Folder As Outlook.Folder, ByVal NoFurtherQuestions As Boolean) As Boolean
  On Error Resume Next 'Utile si email trop grand

    ' Count how many emails should be archived
    ToArchive = 0
    If mysubfolder.Items.Restrict("[Categories] <> 'Archived'").Count > 0 Then
      For Each objItem In mysubfolder.Items.Restrict("[Categories] <> 'Archived'")
        If objItem.MessageClass = "IPM.Note" Then ToArchive = ToArchive + 1
      Next
    End If

    ' Archive emails
    If ToArchive <> 0 Then
      If Not NoFurtherQuestions Then 
        Duration = ToArchive * 2
        YesNo = MsgBox("Do you want to archive " & ToArchive & " emails from " & mysubfolder & " ? To abort during execution, press Ctrl + break. It should take at most " & Duration & " seconds. ", vbYesNoCancel, mysubfolder)
        If YesNo = vbCancel Then 
          SaveEmails = False 
          Exit Function 
        End If 
      End If 

      If NoFurtherQuestions Or YesNo = 6 Then
        For i = mysubfolder.Items.Count To 1 Step -1
          Set objItem = mysubfolder.Items(i)
          If objItem.Categories <> "Archived" And objItem.MessageClass = "IPM.Note" Then
            Set oMail = objItem
            sName = oMail.Subject
            ReplaceCharsForFileName sName, "-"
            dtDate = oMail.ReceivedTime
            sName = Format(dtDate, "yyyymmdd", vbUseSystemDayOfWeek, _
              vbUseSystem) & Format(dtDate, "-hhnnss", _
              vbUseSystemDayOfWeek, vbUseSystem) & "-" & sName & ".msg"
            Debug.Print sPath & sName
            oMail.SaveAs sPath & sName, olMSG
            oMail.Categories = "Archived"
            oMail.Save
          End If

        Next
      End If
    End If

    SaveEmails = True 
End Function
