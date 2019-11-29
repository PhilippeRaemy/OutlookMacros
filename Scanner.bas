Attribute VB_Name = "Scanner"
Option Explicit
Const ScannerExe = """C:\Program Files (x86)\NAPS2\NAPS2.Console.exe"""
' Const ScannerExe = """C:\Program Files (x86)\naps2-3.3.0-experimental-standalone\NAPS2.Console.exe"""

' Const ScannerExe = """C:\Program Files (x86)\naps2-3.3.0-experimental-standalone\NAPS2.Console.exe"""

Private Declare PtrSafe Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public Function Scan( _
    ByVal ScannerProfile As String, _
    ByVal DestinationFolder As String, _
    ByVal NamePattern As String, _
    ByVal OpenFileWhenDone As Boolean, _
    ByVal ReplaceIfExist As Boolean, _
    ByRef defaults As String, _
    Optional ByVal EditFileName As Boolean = False, _
    Optional ByVal AutoScript As String = "" _
) As String
' returns the newly scanned file name
Dim FileName As String, FolderName As String
Dim WShell As New WshShell
Dim cmd As String
Dim fso As New FileSystemObject
Dim f    As FileDialog
Dim startTime As Date
Static scanning As Boolean
On Error GoTo proc_err:
GoTo proc:

proc_err:
    MsgBox Err.Description
    Exit Function
    Resume
proc:

    If scanning Then
        Debug.Print Now & " already scanning, waiting..."
        While scanning: DoEvents: Wend
        Debug.Print Now & " All good, move on..."
    End If

    Dim defaultsDic As Scripting.Dictionary
    Set defaultsDic = ParseToDic(defaults)

    FolderName = ResolvePlaceholders(DestinationFolder, defaultsDic)
    Scanner.EnsureFolderExists FolderName
    FileName = ResolvePlaceholders(FolderName & "\" & NamePattern, defaultsDic) ' the seq placeholder needs full path
    defaults = DicToString(defaultsDic)
    If Not ReplaceIfExist Then
        FileName = EnsureUnique(FileName)
    End If
    If EditFileName Then
        FileName = InputBox("File will be saved as...", "File name", FileName)
        If Trim(FileName) = "" Then
            MsgBox "Scan aborted", vbExclamation
            Exit Function
        End If
    End If
    cmd = ScannerExe & " --profile """ & ScannerProfile & """ --output """ & FileName & """"
    startTime = Now
    Debug.Print Now & " " & cmd
    scanning = True
    Dim rc As Integer
    rc = WShell.Run(cmd, WshWindowStyle.WshWSDisplayChildMinimized, True)
    Debug.Print Now & " " & rc & " " & Int(((Now - startTime)) * 86400) & "[s]"
    scanning = False
    If Not fso.FileExists(FileName) Then
        Err.Raise vbObjectError, "Scan", "The file " & FileName & " was not created. Problem with scanner?"
    End If
    If AutoScript <> "" Then
        WShell.Run """" & AutoScript & """ """ & FileName & """", WshWindowStyle.WshWSActivateChildNoFocus
    End If
    If OpenFileWhenDone Then
        WShell.Run """" & FileName & """", WshWindowStyle.WshWSDisplayChild
    End If
    Scan = FileName
End Function

Private Function ParseToDic(ByVal defaults As String) As Scripting.Dictionary
    Dim dic As Scripting.Dictionary: Set dic = New Scripting.Dictionary
    Dim a As Variant, b As Variant
    Dim key As String, value As String
    For Each a In Split(defaults, ";")
        If a <> "" Then
            b = Split(CStr(a), ":", 2)
            key = Replace(Replace(CStr(b(0)), "&semicolon;", ";"), "&colon;", ":")
            If UBound(b) > 0 Then
                value = Replace(Replace(CStr(b(1)), "&semicolon;", ";"), "&colon;", ":")
            Else
                value = ""
            End If
            dic.add key, value
        End If
    Next a
    Set ParseToDic = dic
End Function

Private Function DicToString(ByVal dic As Scripting.Dictionary) As String
Dim k As Variant
    For Each k In dic.Keys
        DicToString = DicToString _
        & Replace(Replace(CStr(k), ";", "&semicolon;"), ":", "&colon;") & ":" _
        & Replace(Replace(CStr(dic(k)), ";", "&semicolon;"), ":", "&colon;") & ";"
    Next k
End Function
Public Sub LaunchPDFXChange(ByVal FileName As String)
Dim exe As String, wTitle

exe = """C:\Program Files\Tracker Software\PDF Editor\PDFXEdit.exe"""
wTitle = "PDF-XChange Editor"

Dim WShell As New WshShell
' Dim WScript As New w
WShell.Run exe, WshWindowStyle.WshWSActivateChildFocus
 Debug.Print "Activated " & wTitle
 Sleep 1500
 WShell.SendKeys "%FN{UP}{UP}"
 Debug.Print "Done"
'Dim rc As Integer
'rc = WShell.AppActivate(wTitle)
'Debug.Print "rc="; rc
'If rc = 0 Then
'
'    Sleep 500
'    If WShell.AppActivate(wTitle) = 0 Then
'        Debug.Print "Unable to start " & wTitle
'        Exit Sub
'    End If
'    Debug.Print "Started " & wTitle
'    Sleep 500
'End If

End Sub

Public Function EnsureUnique(ByVal FileName As String, Optional ByVal seq As Integer = 1)
    Dim fso As New FileSystemObject
    If Not fso.FileExists(FileName) Then
        EnsureUnique = FileName
        Exit Function
    End If
    Dim parts As Variant
    parts = Split(FileName, ".")
    If seq <= 1 Then
        ReDim Preserve parts(UBound(parts) + 1)
        parts(UBound(parts)) = parts(UBound(parts) - 1)
    End If
    seq = seq + 1
    parts(UBound(parts) - 1) = Format(seq, "000")
    EnsureUnique = EnsureUnique(Join(parts, "."), seq)
End Function
Public Function ResolvePlaceholders(ByVal Model As String, ByRef defaults As Scripting.Dictionary) As String
    Dim re As New RegExp
    Dim Patterns As Variant, Pattern As Variant
    Dim results  As String
    Patterns = Array("\{((input):([^\}:]+)(:([^\}]+))?)\}" _
                   , "\{((inputd):([^\}:]+)(:([+-]\d+)([dmy]))?)\}" _
                   , "\{((now):([^\}]+))\}" _
                   , "\{((seq):(\d+))\}" _
    )
    ResolvePlaceholders = Replace(Model, "\\", "\")
    For Each Pattern In Patterns
        ResolvePlaceholders = TestRegex(CStr(Pattern), ResolvePlaceholders, defaults)
    Next Pattern
End Function

Function TestRegex(ByVal Pattern As String, ByVal Model As String, ByRef dic As Scripting.Dictionary) As String
    Dim re As New RegExp
    Dim Matches As MatchCollection
    Dim Match As Match
    Dim i As Integer
    Dim key As String, previousValue As String, newValue As String
On Error GoTo proc_err:
GoTo proc:

proc_err:
    MsgBox Err.Description
    Exit Function
    Resume
proc:
    
    re.Pattern = Pattern
    TestRegex = Model
    Do
        Set Matches = re.Execute(TestRegex)
        If Matches.Count = 0 Then Exit Do
        For Each Match In Matches
            If Match.SubMatches.Count >= 2 Then
                key = Match.SubMatches(2)
                If dic.Exists(key) Then
                    previousValue = dic(key)
                End If
            End If
            Select Case Match.SubMatches(1)
                Case "input"
                    If Match.SubMatches(4) <> "" Then
                        previousValue = Match.SubMatches(4)
                    End If
                    newValue = InputBox(Match.SubMatches(2), "Provide a value", previousValue)
                    dic(key) = newValue
                    TestRegex = re.Replace(TestRegex, newValue)
                Case "inputd"
                    Dim dte As Date: dte = Now
                    If Match.SubMatches(3) <> "" Then ' offset
                        dte = DateAdd(Match.SubMatches(5), Match.SubMatches(4), dte)
                    ElseIf previousValue <> "" And IsDate(previousValue) Then
                        dte = CDate(previousValue)
                    Else
                        dte = 0
                    End If
                    newValue = InputBox(Match.SubMatches(2), "Provide a value", Format(dte, Match.SubMatches(2)))
                    dic(key) = newValue
                    TestRegex = re.Replace(TestRegex, newValue)
                Case "now"
                    TestRegex = re.Replace(TestRegex, Format(Now, Match.SubMatches(2)))
                Case "seq"
                    TestRegex = GetFreeFileName(re, TestRegex, Match.SubMatches(2))
            End Select
        Next Match
    Loop
End Function

Private Function GetFreeFileName( _
    ByVal re As RegExp, _
    ByVal FilePattern As String, _
    ByVal sequenceFormat As String _
    )
    Dim fs As Scripting.FileSystemObject
    Set fs = New FileSystemObject
    Dim seq As Integer
    Do
        GetFreeFileName = re.Replace(FilePattern, Format(seq, sequenceFormat))
        seq = seq + 1
    Loop Until Not fs.FileExists(GetFreeFileName)
End Function

Public Sub ExploreRegex()

    Dim re As New RegExp
    Dim Matches As MatchCollection
    Dim Match As Match
    Dim sma As Integer

    re.Pattern = "\{((inputd):([^\}:]+)(:([+-]\d+)([dmy]))?)\}"
    Set Matches = re.Execute("Séverine {inputd:yyyy-MM:-1m}.pdf")
    For Each Match In Matches
        Debug.Print Match.SubMatches.Count
        For sma = 0 To Match.SubMatches.Count - 1:
            Debug.Print sma, Match.SubMatches(sma)
        Next
    Next
End Sub

Public Sub EnsureFolderExists(FolderName As String)
    Dim fs As New FileSystemObject
    Dim parentFolderName As String
    Dim a As Variant, i As Integer
    If Not fs.FolderExists(FolderName) Then
        a = Split(FolderName, "\")
        ReDim Preserve a(UBound(a) - 1)
        parentFolderName = Join(a, "\")
        EnsureFolderExists parentFolderName
        If Not fs.FolderExists(FolderName) Then fs.CreateFolder FolderName
    End If
End Sub

