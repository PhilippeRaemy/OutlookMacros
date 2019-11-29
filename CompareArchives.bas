Attribute VB_Name = "CompareArchives"
Option Explicit
Sub main()
  Dim stDestination As store
  Set stDestination = openStore("C:\temp\sevMail\SéverineOnServer.pst")
'  MergeArchives "C:\temp\sevMail\severine.laliveraemy@swissonline.ch.pst", stDestination.GetRootFolder()
'  MergeArchives "C:\temp\sevMail\severine_lalive@swissonline.ch.pst", stDestination.GetRootFolder()
  MergeArchives "C:\temp\sevMail\SéverineOnJaune.pst", stDestination.GetRootFolder()
'  MergeArchives "C:\temp\sevMail\SéverineOnServer.pst", stDestination.GetRootFolder()
End Sub
Function openStore(archiveFileName As String) As store
  Dim myNameSpace As NameSpace, st As store
  Set myNameSpace = Application.GetNamespace("MAPI")
  myNameSpace.AddStore archiveFileName
  Set openStore = myNameSpace.Stores(myNameSpace.Stores.Count)
  Debug.Print "Store " & openStore.filepath & " is open."
End Function

Sub MergeArchives(archiveFileName As String, destination As Folder)
  Dim st As store
  Set st = openStore(archiveFileName)
  mergeFolders st.GetRootFolder(), destination
  st.parent.RemoveStore st.GetRootFolder()
End Sub

Sub mergeFolders(f As Folder, destination As Folder)
  Dim dic As Scripting.Dictionary
  Set dic = New Dictionary
  Dim item As MailItem, subfld As Folder, obj As Object
  Dim hash As String
  Dim i As Integer
  Debug.Print "mergeFolders " & f.folderPath & " into " & destination.folderPath
  For Each obj In destination.Items
    Select Case TypeName(obj)
      Case "MailItem"
        hash = MailHash(obj)
        If Not dic.Exists(hash) Then
          dic.add hash, obj
        End If
    End Select
  Next obj
  Debug.Print dic.Count & " distinct items in " & destination.folderPath
  mergeFolderWithDic dic, f, destination
  For Each subfld In f.folders
    mergeFolders subfld, Utilities.EnsureFolderExists(destination, subfld.name)
  Next subfld
End Sub

Sub mergeFolderWithDic(dic As Scripting.Dictionary, f As Folder, destination As Folder)
Dim item As MailItem, obj As Object
Dim hash As String, i As Integer
  For i = f.Items.Count To 1 Step -1
    Select Case TypeName(f.Items(i))
      Case "MailItem"
        Set item = f.Items(i)
        hash = MailHash(item)
        If dic.Exists(hash) Then
          Debug.Print "Exists : " & hash
        Else
          dic.add hash, item
          item.Move destination
          Debug.Print "Moved  : " & hash
        End If
    End Select
  Next i
End Sub
Function MailHash(mi As MailItem) As String
On Error Resume Next
  MailHash = TypeName(mi)
  MailHash = MailHash & "|" & mailSender(mi)
  MailHash = MailHash & "|" & mi.subject
  MailHash = MailHash & "|" & Format(mi.SentOn, "yyyymmdd hhmmss")
  MailHash = MailHash & "|" & Len(mi.body)
  
End Function
Function mailSender(mi As MailItem) As String
  If mi.sender Is Nothing Then Exit Function
  mailSender = mi.sender.address
End Function
