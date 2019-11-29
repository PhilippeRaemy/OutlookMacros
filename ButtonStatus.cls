VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ButtonStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private PrevCaption As String
Dim ctrls As Variant

Public Sub init(CommandBarName As String, CommandBarControlName As String, Optional exactMatch As Boolean = True)
Dim co As Integer
  ctrls = FindAllCommandBarControls(CommandBarName, CommandBarControlName, exactMatch)
  For co = LBound(ctrls) To UBound(ctrls)
    ctrls(co).Tag = ctrls(co).caption
  Next co
End Sub
Public Sub setDefaultCaption(s As String)
Dim co As Integer
  For co = LBound(ctrls) To UBound(ctrls)
    ctrls(co).Tag = s
    ctrls(co).caption = s
  Next co
End Sub
Public Sub SetCaption(s As String)
Dim co As Integer
  If IsEmpty(ctrls) Then Exit Sub
  For co = LBound(ctrls) To UBound(ctrls)
    ctrls(co).caption = ctrls(co).Tag & "-" & s
    ctrls(co).Enabled = False
  Next co
End Sub
Private Function FindAllCommandBarControls(CommandBarName As String, CommandBarControlName As String, Optional exactMatch As Boolean = True) As Variant
On Error Resume Next
Dim ctrls As Variant
ctrls = Array()
Dim ex As Explorer
Dim bar As CommandBar
Dim co As CommandBarControl
  For Each ex In Application.Explorers
    Set bar = Nothing
    On Error Resume Next
    Set bar = ex.CommandBars(CommandBarName)
    On Error GoTo 0
    If Not (bar Is Nothing) Then
      If exactMatch Then
        Set co = Nothing
        On Error Resume Next
        Set co = bar.Controls(CommandBarControlName)
        On Error GoTo 0
        If Not (co Is Nothing) Then
          ReDim Preserve ctrls(UBound(ctrls) + 1)
          Set ctrls(UBound(ctrls)) = co
        End If
      Else 'use like ...
        For Each co In bar.Controls
          If co.caption Like "*" & CommandBarControlName & "*" Then
            ReDim Preserve ctrls(UBound(ctrls) + 1)
            Set ctrls(UBound(ctrls)) = co
          End If
        Next co
      End If
    End If
  Next ex
  FindAllCommandBarControls = ctrls
End Function

Public Sub Class_Terminate()
'restore initial command caption
  Dim co As Integer
  If IsEmpty(ctrls) Then Exit Sub
  For co = LBound(ctrls) To UBound(ctrls)
    ctrls(co).caption = ctrls(co).Tag
    ctrls(co).Enabled = True
  Next co
End Sub
