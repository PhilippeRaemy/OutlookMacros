Attribute VB_Name = "Archiver"
Option Explicit

Sub ExportSelectedMails()
Dim o As Object
Dim fso As New FileSystemObject
    For Each o In Application.Explorers(1).Selection
        If TypeName(o) = "MailItem" Then
            ExportMail o, FileSystemFolder("a:\MailArchives")
        End If
    Next o
End Sub

Sub ExportOldArchiveMails()
Dim exported As ExportStatus
    Set exported = ExportOldMails(ThisOutlookSession.GetArchivesFolder, FileSystemFolder("a:\MailArchives"), CDate((Year(Now) - 2) & "-01-01"), delete:=False, force:=False)
    Debug.Print "ExportOldArchiveMails: " & exported.ToString()
End Sub
Sub ExportOldArchiveMailsAndDelete()
Dim exported As ExportStatus
    Set exported = ExportOldMails(ThisOutlookSession.GetArchivesFolder, FileSystemFolder("a:\MailArchives"), CDate((Year(Now) - 3) & "-01-01"), delete:=True, force:=True)
    Debug.Print exported.ToString()
End Sub

Function ExportOldMails(root As Outlook.Folder, fileRoot As Scripting.Folder, maxSentOnDate As Date, Optional delete As Boolean = False, Optional force As Boolean = False) As ExportStatus
Dim fld As Outlook.Folder
Dim o As Object, mi As MailItem
Dim mis() As MailItem, idx As Integer, i As Integer
ReDim mis(100)
Set ExportOldMails = New ExportStatus

    For Each fld In root.folders
        ExportOldMails fld, fileRoot, maxSentOnDate, delete, force
    Next fld
    Debug.Print "ExportOldMails " & root.folderPath
    For Each o In root.items
        If TypeName(o) = "MailItem" Then
            Set mi = o
            If mi.SentOn < maxSentOnDate Then
                If idx > UBound(mis) Then ReDim Preserve mis(UBound(mis) + 100)
                Set mis(idx) = mi: idx = idx + 1
            End If
        End If
    Next o
    For i = 0 To idx - 1
        Set ExportOldMails = ExportOldMails.add(ExportMail(mis(i), fileRoot, delete, force))
        DoEvents
    Next i
    Debug.Print "ExportOldMails " & root.folderPath & " " & ExportOldMails.ToString
End Function

Public Function ExportMail(mi As MailItem, rootDirectory As Scripting.Folder, Optional delete As Boolean = False, Optional force As Boolean = False) As ExportStatus
Dim fld As Scripting.Folder, mailFileName As String
Dim att As Attachment, attachmentName As String
Dim filename As String
Dim fso As FileSystemObject: Set fso = New FileSystemObject
Set ExportMail = New ExportStatus
On Error GoTo proc_err
GoTo proc
proc_err:
    MsgBox trace.trace("Error", Err.Number & " " & Err.Description & " in ExportMail"), vbCritical
    Resume Next
    Exit Function
    Resume
proc:
    
    ExportMail.countMails = 1
    Set fld = FileSystemFolder(rootDirectory.path & "\" & mi.parent.FullFolderPath)
    mailFileName = MakeFileName(mi)
    If mi.Attachments.Count > 0 Then
        Set fld = FileSystemFolder(fld.path & "\" & mailFileName)
        For Each att In mi.Attachments
            attachmentName = ""
            On Error Resume Next
            attachmentName = att.filename
            On Error GoTo proc_err
            If Not attachmentName = "" Then
                filename = TruncateFileName(fld.path & "\" & attachmentName)
                If force Or Not fso.FileExists(filename) Then
                    ExportMail.countFiles = ExportMail.countFiles + 1
                    att.SaveAsFile filename
                End If
            End If
        Next att
    End If
    filename = TruncateFileName(fld.path & "\" & mailFileName & ".rtf")
    If force Or Not fso.FileExists(filename) Then
        ExportMail.countFiles = ExportMail.countFiles + 1
        mi.SaveAs filename, OlSaveAsType.olRTF
    End If
    ' if we're reached here without error, we can delete the e-mail
    If delete Then
        mi.delete
        ExportMail.countDeleted = ExportMail.countDeleted + 1
    End If
End Function

