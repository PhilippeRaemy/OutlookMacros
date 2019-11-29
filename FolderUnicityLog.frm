VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FolderUnicityLog 
   Caption         =   "Folder Unicity Log"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10170
   OleObjectBlob   =   "FolderUnicityLog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FolderUnicityLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private running As Boolean
Public fld As MAPIFolder
Public recurse As Boolean
Private Sub cmdCancel_Click()
  log.Text = log.Text & vbCrLf & "Cancelled."
  running = False
  Me.Hide
End Sub

Private Sub cmdRun_Click()
Dim o As Object
On Error GoTo proc_err
GoTo proc
proc_err:
  trace.trace "ERROR", Err.Number & " " & Err.Description & " in FolderUnicityLog.cmdRun_Click"
  MsgBox Err.Number & " " & Err.Description & " in FolderUnicityLog.cmdRun_Click", vbCritical
  Exit Sub
  Resume
proc:

running = True
  
  For Each o In fld.Items
    If TypeName(o) = "MailItem" Then
      On Error Resume Next 'maybe we have a type mismatch if the MI is till in the collection but has been deleted already
      ThisOutlookSession.CheckOneMailUnicity o, False, Me.log
      On Error GoTo proc_err
      DoEvents
      If Not running Then Exit Sub
    End If
  Next o
  running = False
  log.Text = log.Text & vbCrLf & "Done."
End Sub

Private Sub UserForm_Resize()
  Me.log.Height = Me.InsideHeight - 3 * Me.log.Top - Me.cmdRun.Height
  Me.log.Width = Me.InsideWidth - 2 * Me.log.Left
  Me.cmdRun.Top = Me.log.Height + 2 * Me.log.Top
  Me.cmdCancel.Top = Me.log.Height + 2 * Me.log.Top
  Me.cmdCancel.Left = Me.InsideWidth - Me.log.Left - Me.cmdCancel.Width
  Me.cmdRun.Left = Me.InsideWidth - 2 * Me.log.Left - Me.cmdCancel.Width - Me.cmdRun.Width
End Sub
