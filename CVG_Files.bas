Attribute VB_Name = "CVG_Files"
Public Sub StoreDailySheetsAttachments(item As MailItem)
  StoreAttachments item, "S:\ENERGY\Convergence\Daily_sheets"
End Sub
Public Sub StoreAttachments(item As MailItem, FolderName As String)
Dim i As Integer
Dim att As Attachment
Dim filepath As String
Dim save As Boolean
Dim fs As New FileSystemObject
For i = 1 To item.Attachments.Count
  Set att = item.Attachments.item(i)
  
  filepath = FolderName & "\" & Format(item.SentOn, "yyyymmdd") & "_" & att.FileName
  If fs.FileExists(filepath) Then
    Select Case MsgBox("Overwrite " & filepath & ", saved on " & FileDateTime(filepath) & " with new attachment?", vbYesNoCancel)
      Case vbYes: save = True
      Case vbNo: save = False
      Case vbCancel: Exit Sub
    End Select
  Else
    save = True
  End If
  If save Then
    att.SaveAsFile filepath
  End If
Next i
End Sub

