Attribute VB_Name = "MoveItems"
Option Explicit
Sub MoveItems()
    MoveItemsImpl ThisOutlookSession.GetFolder("philippe_raemy@swissonline.ch (old)"), ThisOutlookSession.GetFolder("philippe_raemy@swissonline.ch")
End Sub

Function MoveItemsImpl(ByVal source As Outlook.Folder, ByVal destination As Outlook.Folder, Optional subfolderName As String = "")
Dim miv() As Variant, fld As Outlook.Folder
Dim mi As Variant
Dim i As Integer
    trace.trace "MoveItemsImpl " & source.folderPath
    DoEvents
    If Not subfolderName = "" Then
        Set destination = Utilities.EnsureFolderExists(destination, subfolderName)
    End If
    If source.Items.Count > 0 Then
        ReDim miv(1 To source.Items.Count)
        For Each mi In source.Items
            i = i + 1
            Set miv(i) = mi
        Next mi
        For i = 1 To UBound(miv)
            Utilities.moveItem miv(i), destination, source.name
            DoEvents
        Next i
    End If
    trace.trace "Moved " & i & " items out of " & source.folderPath
    For Each fld In source.folders
        i = i + MoveItemsImpl(fld, destination, fld.name)
    Next fld
    trace.trace "Moved " & i & " items below " & source.folderPath
    MoveItemsImpl = i
End Function
