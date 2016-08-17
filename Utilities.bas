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
    trace.trace context, "DELETED", miv.Subject & "(" & miv.CreationTime & ")"
    trace.trace context, "FROM-->", miv.parent.folderPath
    miv.delete
  ElseIf miv.parent.folderPath = fld.folderPath Then
    trace.trace context, "NOT MOVED ", miv.Subject & "(" & miv.CreationTime & ")"
    trace.trace context, "ON SAME FOLDER->", miv.parent.folderPath
  Else
    trace.trace context, "MOVED ", miv.Subject & "(" & miv.CreationTime & ")"
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
    Dim name As Variant: name = Split(mi.Sender.name, "=")
    MakeFileName = Format(mi.SentOn, "yyyy-mm-dd hh.mm.ss") & " [" & name(UBound(name)) & "] " & mi.Subject
    For i = 1 To Len(specialChars)
        MakeFileName = Replace(MakeFileName, Mid(specialChars, i, 1), "_")
    Next i
End Function

Public Function TruncateFileName(ByVal filename As String, Optional ByVal maxlength As Integer = 255) As String
    If Len(filename) <= maxlength Then
        TruncateFileName = filename
        Exit Function
    End If
    Dim chunks As Variant: chunks = Split(filename, ".")
    Dim ext As String: ext = "." & chunks(UBound(chunks))
    ReDim Preserve chunks(UBound(chunks) - 1)
    TruncateFileName = Left(Join(chunks, "."), maxlength - Len(ext)) & ext
End Function
