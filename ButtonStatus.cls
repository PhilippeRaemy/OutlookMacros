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

Public Sub init(ByVal CommandBarName As String, ByVal CommandBarControlName As String)
Dim co As Integer
  ctrls = FindAllCommandBarControls(CommandBarName, CommandBarControlName)
  If LBound(ctrls) > UBound(ctrls) Then ctrls = CreateCommandBarControls(CommandBarName, CommandBarControlName)
  For co = LBound(ctrls) To UBound(ctrls)
    ctrls(co).Tag = ctrls(co).Caption
  Next co
End Sub
Public Sub SetCaption(ByVal s As String)
Dim co As Integer
Dim ctrl As CommandBarButton
  For co = LBound(ctrls) To UBound(ctrls)
    Set ctrl = ctrls(co)
    ctrl.Caption = ctrl.Tag & "-" & s
    ctrl.Enabled = True
    ' ctrl.SetFocus
    ctrl.Enabled = False
    ctrl.Visible = False
    ctrl.Visible = True
    DoEvents
  Next co
End Sub
Private Function CreateCommandBarControls(ByVal CommandBarName As String, ByVal CommandBarControlName As String) As Variant
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
    If (bar Is Nothing) Then
      Set bar = ex.CommandBars.add(name:=CommandBarName)
      bar.Visible = True
    End If
    Set co = bar.Controls.add(Type:=msoControlButton)
    co.Tag = CommandBarControlName
    co.Caption = CommandBarControlName
    ReDim Preserve ctrls(UBound(ctrls) + 1)
    Set ctrls(UBound(ctrls)) = co
  Next ex
  CreateCommandBarControls = ctrls

End Function
Private Function FindAllCommandBarControls(ByVal CommandBarName As String, ByVal CommandBarControlName As String) As Variant
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
      Set co = Nothing
      On Error Resume Next
      Set co = bar.Controls(CommandBarControlName)
      On Error GoTo 0
      If Not (co Is Nothing) Then
        ReDim Preserve ctrls(UBound(ctrls) + 1)
        Set ctrls(UBound(ctrls)) = co
      End If
    End If
  Next ex
  FindAllCommandBarControls = ctrls
End Function

Public Sub Class_Terminate()
'restore initial command caption
  Dim co As Integer
  For co = LBound(ctrls) To UBound(ctrls)
    ctrls(co).Caption = ctrls(co).Tag
    ctrls(co).Enabled = True
  Next co
End Sub
