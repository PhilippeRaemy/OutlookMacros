Attribute VB_Name = "Utilities"
Option Explicit
Global trace As New Tracer
Sub moveItem(ByVal miv As Variant, ByVal fld As Outlook.Folder, ByVal context As String)
On Error GoTo proc_err
GoTo proc
proc_err:
  trace.trace context, "ERROR", Err.Number & " " & Err.Description & " in moveItem"
  Exit Sub
  Resume
proc:
  If fld Is Nothing Then
    trace.trace context, "DELETED", miv.subject & "(" & miv.CreationTime & ")"
    trace.trace context, "FROM-->", miv.parent.folderPath
    miv.delete
  ElseIf miv.parent.folderPath = fld.folderPath Then
    trace.trace context, "NOT MOVED ", miv.subject & "(" & miv.CreationTime & ")"
    trace.trace context, "ON SAME FOLDER->", miv.parent.folderPath
  Else
    trace.trace context, "MOVED ", miv.subject & "(" & miv.CreationTime & ")"
    trace.trace context, "FROM->", miv.parent.folderPath
    trace.trace context, "--->TO", fld.folderPath
    miv.Move fld
  End If
End Sub

Public Function FileSystemFolder(ByVal path As String) As Scripting.Folder
Dim fso As New FileSystemObject
Dim parent As Scripting.Folder
Dim pathSegments As Variant
    
    If fso.FolderExists(path) Then
        Set FileSystemFolder = fso.GetFolder(path)
    Else
        pathSegments = Split(path, "\")
        If UBound(pathSegments) < 1 Then
            On Error GoTo 0
            Err.Raise vbObjectError, "FileSystemFolder", "Could not create a missing drive or UNC path specification"
        Else
            ReDim Preserve pathSegments(UBound(pathSegments) - 1)
            On Error Resume Next
            Set parent = FileSystemFolder(Join(pathSegments, "\"))
            Dim pathname As String: pathname = Left(path, (255 + Len(parent.path)) / 2)
            If fso.FolderExists(pathname) Then
                Set FileSystemFolder = fso.GetFolder(pathname)
                Exit Function
            End If
            Set FileSystemFolder = fso.CreateFolder(pathname)
            If Err.Number <> 0 Then
                Dim ErrNum As Integer: ErrNum = Err.Number
                On Error GoTo 0
                Err.Raise ErrNum, "FileSystemFolder", "Error finding path `" & path & "`"
            End If
        End If
    End If
End Function

Public Function MakeFileName(ByVal mi As MailItem) As String
Static specialChars As String
Dim i As Integer
    If specialChars = "" Then
        specialChars = ":|{}\/%?*^&<>""'"
    End If
    Dim name As Variant: name = Split(mi.sender.name, "=")
    MakeFileName = Format(mi.SentOn, "yyyy-mm-dd hh.mm.ss") & " [" & name(UBound(name)) & "] " & mi.subject
    For i = 1 To Len(specialChars)
        MakeFileName = Replace(MakeFileName, Mid(specialChars, i, 1), "_")
    Next i
End Function

Public Function TruncateFileName(ByVal FileName As String, Optional ByVal maxlength As Integer = 255) As String
    If Len(FileName) <= maxlength Then
        TruncateFileName = FileName
        Exit Function
    End If
    Dim chunks As Variant: chunks = Split(FileName, ".")
    Dim ext As String: ext = "." & chunks(UBound(chunks))
    ReDim Preserve chunks(UBound(chunks) - 1)
    TruncateFileName = Left(Join(chunks, "."), maxlength - Len(ext)) & ext
End Function
Public Function EnsureFolderExists(parentFolder As Outlook.Folder, folderPath As String) As Outlook.Folder
Dim a As Variant
Dim i As Integer
Dim parentPath As String

Set EnsureFolderExists = parentFolder

a = Split(folderPath, "\")
For i = 0 To UBound(a)
  If a(i) <> "" Then
    On Error Resume Next
    Set EnsureFolderExists = EnsureFolderExists.folders(a(i))
    If Err.Number <> 0 Then
      On Error GoTo 0
      Set EnsureFolderExists = EnsureFolderExists.folders.add(a(i))
    End If
  End If
Next i
End Function

Public Sub BubbleSort(list As Variant)
'   Sorts an array using bubble sort algorithm
    Dim First As Integer, Last As Long
    Dim i As Long, j As Long
    Dim Temp As Variant
    
    First = LBound(list)
    Last = UBound(list)
    For i = First To Last - 1
        For j = i + 1 To Last
            If list(i) > list(j) Then
                Temp = list(j)
                list(j) = list(i)
                list(i) = Temp
            End If
        Next j
    Next i
End Sub


