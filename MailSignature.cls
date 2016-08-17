VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MailSignature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private lBody As String
Private lAttach As String
Private lMi As MailItem
Public Property Get mail() As MailItem
  Set mail = lMi
End Property
Public Property Get body() As String
  body = lBody
End Property
Public Property Get Attach() As String
  Attach = lAttach
End Property
Public Function contains(ByVal s As MailSignature) As Boolean
  contains = (Me.CompareTo(s) > 0)
End Function
Public Function isContainedIn(ByVal s As MailSignature) As Boolean
  isContainedIn = (Me.CompareTo(s) < 0)
End Function
Public Function CompareTo(ByVal s As MailSignature) As Integer
  If InStr(1, lBody, s.body, vbTextCompare) > 0 Then
    CompareTo = 1
  ElseIf InStr(1, s.body, lBody, vbTextCompare) > 0 Then
    CompareTo = -1
  Else
    CompareTo = 0
  End If
  If CompareTo <> 0 Then
    If lAttach = s.Attach Then
      'do nothing: keep prev value...
    ElseIf InStr(1, lAttach, s.Attach, vbTextCompare) > 0 And CompareTo = 1 Then
      CompareTo = 1
    ElseIf InStr(1, s.Attach, lAttach, vbTextCompare) > 0 And CompareTo = -1 Then
      CompareTo = -1
    Else
      CompareTo = 0
    End If
  End If
End Function
Public Sub init(ByVal mi As MailItem)
  lBody = mailBodySignature(mi)
  lAttach = mailAttachmentsSignature(mi)
  Set lMi = mi
End Sub
Private Function mailAttachmentsSignature(ByVal mi As MailItem) As String
Dim a As Attachment
mailAttachmentsSignature = ""
For Each a In mi.Attachments
  If mailAttachmentsSignature <> "" Then mailAttachmentsSignature = mailAttachmentsSignature & ", "
  mailAttachmentsSignature = mailAttachmentsSignature & _
    a.Class & ":" & a.Position & ":" & a.DisplayName
Next a
End Function
Private Function mailBodySignature(ByVal mi As MailItem) As String
  mailBodySignature = Replace(Replace(Replace(Replace(Replace(mi.body _
    , vbCrLf, "") _
    , vbCr, "") _
    , vbLf, "") _
    , vbTab, "") _
    , " ", "")
End Function
Public Property Get ExtendedSubject() As String
  ExtendedSubject = "«" & lMi.Subject & "»(" & lMi.SenderName & ", " & lMi.SentOn & ")"
End Property
