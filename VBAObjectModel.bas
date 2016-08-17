Attribute VB_Name = "VBAObjectModel"
Option Explicit

Public Sub ExportCode()
  Dim c As Integer, l As Integer
  Dim VBProj
  Dim Extension As Scripting.Dictionary
  Set Extension = New Scripting.Dictionary
  Extension.add 1, ".bas"
  Extension.add 2, ".cls"
  Extension.add 3, ".frm"
  Extension.add 100, ".ws.bas"
  
  Dim fso As FileSystemObject: Set fso = New FileSystemObject
  Dim filename As String
  Dim ts As Scripting.TextStream
  Dim code As String
  Dim fileChanged As Boolean
  
  Set VBProj = Application.VBE.ActiveVBProject
  Debug.Print VBProj.filename
  For c = 1 To VBProj.VBComponents.Count
    filename = VBProj.filename & "." & VBProj.VBComponents(c).name & Extension(VBProj.VBComponents(c).Type)
    code = "' ####################" & vbCrLf & "' " & filename & vbCrLf & "' ####################"
    For l = 1 To VBProj.VBComponents(c).CodeModule.CountOfLines
      code = code & vbCrLf & VBProj.VBComponents(c).CodeModule.Lines(l, 1)
    Next l
    fileChanged = True
    If fso.FileExists(filename) Then
      Set ts = fso.OpenTextFile(filename)
      If ts.ReadAll = code Then
        fileChanged = False
        ' Debug.Print " file "; filename; " didn't change"
      End If
      ts.Close
    End If
    If fileChanged Then
      Debug.Print " file "; filename; " changed"
      Set ts = fso.CreateTextFile(filename, True, False)
      ts.Write code
      ts.Close
    End If
  Next c
End Sub




' ==============================================================
' * Please note that Microsoft provides programming examples
' * for illustration only, without warranty either expressed or implied,
' * including, but not limited to, the implied warranties of merchantability
' * and/or fitness for a particular purpose. Any use by you of the code provided
' * in this blog is at your own risk.
'===============================================================

Sub CheckIfVBAAccessIsOn()

'[HKEY_LOCAL_MACHINE/Software/Microsoft/Office/10.0/Excel/Security]
'"AccessVBOM"=dword:00000001
 
Dim strRegPath As String
strRegPath = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Excel\Security\AccessVBOM"

If TestIfKeyExists(strRegPath) = False Then
'   Dim WSHShell As Object
'   Set WSHShell = CreateObject("WScript.Shell")
'   WSHShell.RegWrite strRegPath, 3, "REG_DWORD"
   MsgBox "A change has been introduced into your registry configuration. Pease restart Excel."
   WriteVBS
   Application.Quit
End If

Dim VBAEditor As Object     'VBIDE.VBE
Dim VBProj    As Object     'VBIDE.VBProject
Dim tmpVBComp As Object     'VBIDE.VBComponent
Dim VBComp    As Object     'VBIDE.VBComponent
    
Set VBAEditor = Application.VBE
Set VBProj = Application.ActiveWorkbook.VBProject
   

Dim counter As Integer

For counter = 1 To VBProj.References.Count
  Debug.Print VBProj.References(counter).FullPath
  'Debug.Print VBProj.References(counter).Name
  Debug.Print VBProj.References(counter).Description
  Debug.Print "---------------------------------------------------"
Next
 
End Sub

Function TestIfKeyExists(ByVal path As String)
 Dim WshShell As Object
 Set WshShell = CreateObject("WScript.Shell")
 On Error Resume Next
 WshShell.RegRead path
 
    If Err.Number <> 0 Then
       Err.Clear
       TestIfKeyExists = False
    Else
       TestIfKeyExists = True
    End If
 On Error GoTo 0
End Function

Sub WriteVBS()
Dim objFile     As Object
Dim objFSO      As Object
Dim codePath    As String

codePath = Application.ActiveDocument.path & "\reg_setting.vbs"

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(codePath, 2, True)

objFile.WriteLine (" On Error Resume Next")
objFile.WriteLine ("")
objFile.WriteLine ("Dim WshShell")
objFile.WriteLine ("Set WshShell = CreateObject(""WScript.Shell"")")
objFile.WriteLine ("")
objFile.WriteLine ("MsgBox ""Click OK to complete the setup process.""")
objFile.WriteLine ("")
objFile.WriteLine ("Dim strRegPath")
objFile.WriteLine ("Dim Application_Version")
objFile.WriteLine ("Application_Version = """ & Application.Version & """")
objFile.WriteLine ("strRegPath = ""HKEY_CURRENT_USER\Software\Microsoft\Office\"" & Application_Version & ""\Excel\Security\AccessVBOM""")
objFile.WriteLine ("WScript.echo strRegPath")
objFile.WriteLine ("WshShell.RegWrite strRegPath, 1, ""REG_DWORD""")
objFile.WriteLine ("")
objFile.WriteLine ("If Err.Code <> o Then")
objFile.WriteLine ("   MsgBox ""Error"" & Chr(13) & Chr(10) & Err.Source & Chr(13) & Chr(10) & Err.Message")
objFile.WriteLine ("End If")
objFile.WriteLine ("")
objFile.WriteLine ("WScript.Quit")

objFile.Close
Set objFile = Nothing
Set objFSO = Nothing

'run the VBscript code
Shell "cscript " & codePath, vbNormalFocus

End Sub

