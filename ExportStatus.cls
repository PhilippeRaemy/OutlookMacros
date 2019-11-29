VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExportStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public countDeleted As Integer
Public countFiles As Integer
Public countMails As Integer

Public Function add(other As ExportStatus) As ExportStatus
Set add = New ExportStatus
    add.countDeleted = countDeleted + other.countDeleted
    add.countFiles = countFiles + other.countFiles
    add.countFiles = countMails + other.countMails
    add.countMails = countMails + other.countMails
End Function

Public Function ToString() As String
    ToString = "Deleted mails: " & countDeleted & _
             ", Exported mails: " & countMails & _
             ", Exported files: " & countFiles
End Function
