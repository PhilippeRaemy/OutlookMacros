VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IndexingStatus 
   Caption         =   "Initial Indexing Status"
   ClientHeight    =   645
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   8565.001
   OleObjectBlob   =   "IndexingStatus.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "IndexingStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Public MaxValue As Long
Private lValue As Long
Public LabelPrefix As String
Public Sub Setlabel(str As String)
  Me.Label.caption = LabelPrefix & str
End Sub
Public Property Get value() As Long
  value = lValue
End Property
Public Property Let value(v As Long)
  lValue = v
  UserForm_Resize
End Property
Private Sub UserForm_Resize()
  Dim wi As Long
  If MaxValue = 0 Then
    wi = 0
  Else
    wi = lValue * Me.InsideWidth / MaxValue
  End If
  Me.Label.Move 0, 0, Me.InsideWidth
'  wi = 100
  Me.ProgressBar.Move 0, Me.Label.Height, wi, Me.InsideHeight - Me.Label.Height
  Me.ProgressBar.caption = lValue & " - " & CLng(wi * 100 / Me.InsideWidth) & "%"
  DoEvents
End Sub
