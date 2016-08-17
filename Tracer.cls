VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tracer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private logfile As Scripting.TextStream
Private startDate As Date
Private fn As Integer
Private Sub Class_Initialize()
Dim fs As New Scripting.FileSystemObject
Set logfile = fs.OpenTextFile(Environ("USERPROFILE") & "\logs\mailAutoclassify_" & Format(Now, "yyyymmdd") & ".log", ForAppending, True)
startDate = Now
End Sub
Public Function trace(ByVal context As String, ParamArray strings() As Variant) As String
Dim i As Long
On Error GoTo proc_err
GoTo proc
proc_err:
  logfile.Write "tracer.trace ERROR:" & Err.Number & " " & Err.Description
  Resume Next
proc:
  trace = Format(Now, "yyyy-mm-dd hh:mm:ss [") & context & "]"
  trace = context
  For i = LBound(strings) To UBound(strings)
    'Debug.Print strings(i),
    trace = trace & vbTab & CStr(strings(i))
  Next i
  logfile.WriteLine trace
  Debug.Print trace
  OptionallyFlush
End Function
Private Sub OptionallyFlush()
  If Int(Now) <> Int(startDate) Or Now - startDate > 1 / 1440 * 5 Then
    Class_Terminate
    Class_Initialize
  End If
End Sub
Private Sub Class_Terminate()
logfile.Close
Set logfile = Nothing
End Sub
