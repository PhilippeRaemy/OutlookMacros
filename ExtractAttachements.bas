Attribute VB_Name = "ExtractAttachements"
Option Explicit
Dim fso As New Scripting.FileSystemObject
Sub main()
  Dim stArchive As store
  Set stArchive = openStore("D:\Users\Philippe\Documents\Outlook Files\philippe_raemy@swissonline.ch.pst")
  Dim oFolder As Folder
  Set oFolder = Application.Session.GetDefaultFolder(olFolderInbox)
  Set stArchive = oFolder.store

  ExtractFiles "D:\Users\Philippe\Documents\FINANCE\factures\Telecoms\Sunrise", stArchive.GetRootFolder().folders("eBusiness").folders("Sunrise")
End Sub

Sub ExtractO3CPieces()
    ExtractPieces "O3C", 2018, "D:\Users\Philippe\Documents\O3C finances\"
End Sub

Sub ExtractCambristiPieces()
    ExtractPieces "Cambristi", 2018, "D:\Users\Philippe\Documents\Cambristi\"
End Sub

Sub ExtractPartitionsPieces()
    ExtractPieces "Inbox", 2019, "D:\Users\Public\Partitions\o3c\", "partitionstroischene@gmail.com"
End Sub


Sub ExtractPieces(ByVal FolderName As String, ByVal year As Integer, ByVal target As String, Optional store As String)
    Dim oFolder As Folder
    Dim obj As Object, mi As MailItem, att As Attachment
    If store = "" Then
        Set oFolder = Application.Session.GetDefaultFolder(olFolderInbox).parent.folders(FolderName)
    Else
        Set oFolder = Application.Session.folders(store).folders(FolderName)
    End If
    Dim fileroot As String
    Dim targetFolder As Scripting.Folder
    Set targetFolder = Utilities.FileSystemFolder(target & year & "\Attachments")
    For Each obj In oFolder.Items
        If TypeName(obj) = "MailItem" Then
            Set mi = obj
            If DatePart("yyyy", mi.SentOn) = year Then
                fileroot = targetFolder.path & "\" & Format(mi.SentOn, "yyyymmdd_hhmmss") & "_" & CleanName(mi.subject)
                Debug.Print fileroot
                For Each att In mi.Attachments
                    SaveAttachment att, fileroot
                Next att
            End If
        End If
    Next obj
End Sub

Sub ClassifyO3CMails()
    Dim oFolder As Folder
    Dim obj As Object, mi As MailItem, ri As ReportItem
    Set oFolder = Application.Session.GetDefaultFolder(olFolderInbox).parent.folders("Archives").folders("O3C")
    Dim target As Outlook.Folder
    Dim year As Integer
    Dim i As Integer
    For i = oFolder.Items.Count To 1 Step -1
        Set obj = oFolder.Items(i)
        Select Case TypeName(obj)
            Case "MailItem":
                Set mi = obj
                year = DatePart("yyyy", mi.SentOn)
                Debug.Print year, mi.SentOn, mi.subject
            Case "ReportItem":
                Set ri = obj
                year = DatePart("yyyy", ri.CreationTime)
                Debug.Print year, ri.CreationTime, ri.subject
        End Select
        If year <= 2018 Then
            Set target = Utilities.EnsureFolderExists(oFolder, CStr(year))
            obj.Move target
        End If
    Next i
End Sub

Function GetFolder(path As String) As Scripting.Folder
    Dim subfolderName As String
    If fso.FolderExists(path) Then
        Set GetFolder = fso.GetFolder(path)
    Else
        subfolderName = Mid(path, Len(fso.GetParentFolderName(path)) + 1)
        If Left(subfolderName, 1) = "\" Then subfolderName = Mid(subfolderName, 2)
        Set GetFolder = GetFolder(fso.GetParentFolderName(path)).SubFolders.add(subfolderName)
    End If
End Function
Function openStore(archiveFileName As String) As store
  Dim myNameSpace As NameSpace, st As store
  Set myNameSpace = Application.GetNamespace("MAPI")
  myNameSpace.AddStore archiveFileName
  Set openStore = myNameSpace.Stores(myNameSpace.Stores.Count)
  Debug.Print "Store " & openStore.filepath & " is open."
End Function

Sub ExtractFiles(path As String, oFolder As Outlook.Folder, Optional SaveAttachements As Boolean = True, Optional SaveMailMessage As Boolean = False, Optional FromDate As Date = #1/1/1900#, Optional ToDate As Date = #1/1/2100#)
 
  Dim mi As MailItem, subfld As Outlook.Folder, obj As Object
  Dim ai As AppointmentItem
  Dim att As Attachment, atts As Attachments
  Dim FileName As String, fileroot As String
  Dim fFolder As Scripting.Folder
  Set fFolder = GetFolder(path)
  Dim i As Integer, j As Integer
  On Error GoTo err_proc
  GoTo proc
err_proc:
  Debug.Print "Error " & Err.Number & ", " & Err.Description & vbCrLf & TypeName(obj) & vbCrLf & mi.parent.folderPath & "\" & mi.subject & " - " & mi.SentOn
  Resume Next
proc:
  For Each obj In oFolder.Items
    Set atts = Nothing
    Select Case TypeName(obj)
      Case "MailItem"
        Set mi = obj
        fileroot = fFolder.path & "\" & Format(mi.SentOn, "yyyymmdd_hhmmss") & "_" & CleanName(mi.subject)
        If SaveAttachements Then
          For Each att In mi.Attachments
              SaveAttachment att, fileroot
          Next att
        End If
        On Error GoTo err_proc
        If SaveMailMessage Then
            mi.SaveAs Left(fileroot, 251) & ".msg", olMSG
        End If
      Case "AppointmentItem"
        Set ai = obj
        fileroot = fFolder.path & "\" & Format(ai.Start, "yyyymmdd_hhmmss") & "_" & CleanName(ai.subject)
        If SaveAttachements Then
          For Each att In ai.Attachments
              SaveAttachment att, fileroot
          Next att
        End If
        If SaveMailMessage Then
            ai.SaveAs fileroot & ".msg", olMSG
        End If
    End Select
    
  Next obj
  For Each subfld In oFolder.folders
    ExtractFiles path & "\" & CleanName(subfld.name), subfld
  Next subfld
End Sub
Function CleanName(name As String) As String
  Const undesired = ":\/?*<>|&""”“+%!"
  Dim i As Integer
  CleanName = name
  For i = 1 To Len(undesired)
    CleanName = Replace(CleanName, Mid(undesired, i, 1), "_")
  Next i
  For i = 1 To 31
    CleanName = Replace(CleanName, Chr(i), "_")
  Next i
End Function
Sub SaveAttachment(ByVal att As Attachment, ByVal fileroot As String)
  On Error GoTo err_proc
  GoTo proc
err_proc:
  Debug.Print "Error " & Err.Number & ", " & Err.Description
  Resume Next
proc:
    Dim FileName As String
    Dim FileSuffix As String
    Dim NameParts As Variant, Extension As String
    On Error Resume Next
    FileSuffix = "_" & CleanName(att.FileName)
    If Err.Number <> 0 Then
      Exit Sub
    End If
    Err.Clear
    On Error GoTo 0
    Dim j As Integer
    FileName = fileroot & FileSuffix
    NameParts = Split(FileName, ".")
    If UBound(NameParts) = 0 Then
      Extension = ""
    Else
      Extension = "." & NameParts(UBound(NameParts))
      FileName = Left(FileName, Len(FileName) - Len(Extension))
    End If
    If Len(FileName) + Len(Extension) > 250 Then
      FileName = Left(FileName, 250 - Len(Extension))
    End If
    j = 1
    fileroot = FileName
'    While fso.FileExists(FileName + Extension)
'        j = j + 1
'        FileName = fileroot & "(" & j & ")"
'    Wend
    If Not fso.FileExists(FileName + Extension) Then
        Debug.Print " ==> " & FileName + Extension
        Dim attPathName As String
        On Error Resume Next
        attPathName = att.pathname
        If Err = 0 And Left(attPathName, 4) = "http" Then
            MsgBox "Cannot save locally a cloud file", vbExclamation Or vbOKOnly, "ExtractAttachments"
            att.parent.Display
        Else
            att.SaveAsFile FileName + Extension
        End If
        If Err.Number <> 0 Then
            MsgBox "Cannot save this attachment:" & vbCrLf & Err.Description, vbExclamation Or vbOKOnly, "ExtractAttachments"
            att.parent.Display
        End If
        Err.Clear
    End If
End Sub
