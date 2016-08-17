VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExportStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public countMails As Integer
Public countFiles As Integer
Public countDeleted As Integer

Public Function add(other As ExportStatus) As ExportStatus
    Set add = New ExportStatus
    add.countMails = Me.countMails + other.countMails
    add.countFiles = Me.countFiles + other.countFiles
    add.countDeleted = Me.countDeleted + other.countDeleted
End Function

Public Function ToString() As String
    ToString = _
        "Count Mails = " & countMails & ", " & _
        "Count Files = " & countFiles & ", " & _
        "Count Deleted = " & countDeleted
End Function
