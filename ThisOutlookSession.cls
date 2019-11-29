VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Const PrincipalMail = "philippe_raemy@swissonline.ch"
Public ClassifyMailSearchComp As Boolean
Private CheckUnicityMailSearchComp(1) As Boolean
'mailIndex is a dictionnary of dictionnaries.
'1st level indexed with conversation topics
'2nd level indexed with folder-ids and containing TopicEntry objects (folderpath + count of mails)
Private MailIndex As Scripting.Dictionary
Private LastIndexDump As Date
Public disableAutoClassify As Boolean
Private WithEvents GlobFldInbox As Folder
Attribute GlobFldInbox.VB_VarHelpID = -1
Private WithEvents Items As Outlook.Items
Attribute Items.VB_VarHelpID = -1
Private Sub Application_Startup()
  Set GlobFldInbox = Application.Session.GetDefaultFolder(olFolderInbox)
  StartMakeMailIndex True
  Dim olApp As Outlook.Application
  Dim objNS As Outlook.NameSpace
  Set olApp = Outlook.Application
  Set objNS = olApp.GetNamespace("MAPI")
  ' default local Inbox
  Set Items = objNS.GetDefaultFolder(olFolderInbox).Items
End Sub

Private Sub Items_ItemAdd(ByVal item As Object)

  On Error GoTo ErrorHandler
  Dim Msg As Outlook.MailItem
  If TypeName(item) = "MailItem" Then
    AutoCategorizeSenderMails "", item
  End If
ProgramExit:
  Exit Sub
ErrorHandler:
  MsgBox Err.Number & " - " & Err.Description
  Resume ProgramExit
End Sub
Public Sub RedoAllCategorization()
Dim item As Object
    For Each item In Application.Session.GetDefaultFolder(olFolderInbox).Items
        Items_ItemAdd item
    Next item
End Sub

Private Sub Application_Quit()
  mailIndexDump
End Sub

Private Sub Application_ItemSend(ByVal item As Object, Cancel As Boolean)
Dim mi As MailItem
  Debug.Print
  Select Case TypeName(item)
    Case "MailItem"
      If Not disableAutoClassify Then ClassifyMail item, Cancel
  End Select
  Exit Sub
  Cancel = True
End Sub

'Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Search)
'  trace.trace "AdvancedSearchComplete Event", "tag=" & SearchObject.Tag & ", search=" & SearchObject.Filter
'  Select Case SearchObject.Tag
'    Case "AutoClassifyMail": ClassifyMailSearchComp = True
'    Case "CheckMailUnicity_0": CheckUnicityMailSearchComp(0) = True
'    Case "CheckMailUnicity_1": CheckUnicityMailSearchComp(1) = True
'  End Select
'End Sub
Private Sub ClassifyMail(mi As MailItem, ByRef Cancel As Boolean)
     If mi.To = "Spam, Collection - Collection_Spam@cargill.com" Then
     Else
          ClassifyOneMail mi, True, Cancel
     End If
End Sub
Private Function ClassifyOneMail(mi As MailItem, fromEvent As Boolean, Optional ByRef Cancel As Boolean = False, Optional unAttended As Boolean = False, Optional suggestedFld As Folder = Nothing) As Outlook.Folder
     Dim reindexRequired As Boolean
     Dim emsg     As String
On Error GoTo proc_err
GoTo proc
proc_err:
     If Err.Number = -2147220991 Then ' The messaging interface has returned an unknown error. If the problem persists, restart Outlook.
          reindexRequired = True
          Resume Next
     End If
     emsg = Err.Number & " " & Err.Description & " in ClassifyOneMail"
     trace.trace "ERROR", emsg
     MsgBox emsg, vbCritical
     Exit Function
     Resume
     Cancel = True
     Exit Function
proc:
     
     Dim sch As Outlook.Search
     Dim rsts As Outlook.results
     Dim i As Integer, maxFound As Integer
     'timeout is 5 seconds
     Const timeout As Single = 1 / 24 / 60 / 12
     
     
     Dim folders As Scripting.Dictionary
     Dim foldersCount As Scripting.Dictionary
     Set folders = New Dictionary
     Set foldersCount = New Dictionary
     Dim FolderName As String
     Dim fld As Outlook.Folder
     Dim fldParent As Outlook.Folder
     Dim startTime As Date
     Dim Msg As String
     Dim folderId As String
     Dim folderIdKey As String
     Dim fldInbox As Folder: Set fldInbox = Application.Session.GetDefaultFolder(olFolderInbox)
     
     
     Dim foundMi As MailItem
     Dim omi As Object
     Dim copyMi As MailItem
     Dim tEntry As TopicEntry
     
     Dim topic As String
     Dim topicDic As Scripting.Dictionary
     
     topic = mi.ConversationTopic
     Msg = ""
     If MailIndex Is Nothing Then
          Set MailIndex = New Scripting.Dictionary
          reindexRequired = True
     End If
          
     If MailIndex.Exists(topic) Then
          Set topicDic = MailIndex(topic)
          maxFound = 0
          
          For i = 0 To topicDic.Count - 1
               If Not IsEmpty(topicDic.Items(i)) Then
                    Set tEntry = topicDic.Items(i)
                    If tEntry.Count > maxFound Then
                         maxFound = tEntry.Count
                         folderId = topicDic.Keys(i)
                    End If
               End If
          Next i
          On Error Resume Next
          Set fld = Application.Session.GetFolderFromID(folderId)
          If Err.Number = 0 Then
               On Error GoTo proc_err
               Msg = "Save to folder «" & fld.folderPath & "»?"
               If unAttended Then
                    trace.trace "ClassifyOneMail", Msg
               ElseIf safeEntryId(fld) = safeEntryId(suggestedFld) Then
                    trace.trace "ClassifyOneMail", Msg
               Else
                    Select Case MsgBox(Msg, vbQuestion Or vbYesNoCancel)
                         Case vbYes
                         Case vbCancel
                              Cancel = (vbYes = MsgBox("Cancel message sending?", vbYesNo))
                              Exit Function
                         Case vbNo
                              Set fld = Nothing
                              Set fld = Application.Session.PickFolder
                              If Not fld Is Nothing Then folderId = fld.EntryID
                    End Select
               End If
          ElseIf Err.Number = -2147220991 Then '     The messaging interface has returned an unknown error. If the problem persists, restart Outlook.
               On Error GoTo proc_err
               If unAttended Then
                    trace.trace "ClassifyOneMail", "folder «" & folderId & "» not found!!!"
                    Exit Function '---------------------------------------------------->>>>>>>>>>>
               End If
               Set fld = Application.Session.PickFolder
               folderId = fld.EntryID
               reindexRequired = True
          Else
               emsg = Err.Number & " " & Err.Description & " in ClassifyOneMail"
               trace.trace "ERROR", emsg
               On Error GoTo proc_err
          End If
     ElseIf unAttended Then
          Exit Function '---------------------------------------------------->>>>>>>>>>>
     ElseIf Not (suggestedFld Is Nothing) Then
          Set fld = suggestedFld
     End If
     If fld Is Nothing Then 'pick manually...
          Set topicDic = Nothing
          Set fld = Application.Session.PickFolder
          If fld Is Nothing Then
               If (Not fromEvent) Or unAttended Then
                    Cancel = True
               Else
                    Cancel = (vbYes = MsgBox("Cancel message sending?", vbYesNo))
               End If
               If Cancel Then Exit Function
          End If
     End If
     If Not fld Is Nothing Then
          If fromEvent Then
               mailIndexAppend mi.ConversationTopic, fld.EntryID, fld.folderPath
               Set mi.SaveSentMessageFolder = fld
               'DemoItemProperties mi
               'mi.PropertyAccessor
               'If mi.ItemProperties.item("Keywords") Is Nothing Then
               '     mi.ItemProperties.Add("Keywords", olKeywords).Value = "Classified"
               'Else
               '     mi.ItemProperties.item("keywords").Value = "Classified"
               'End If
               'Debug.Print
               
          Else
               mailIndexAppend mi.ConversationTopic, fld.EntryID, fld.folderPath
               If Not mi.UnRead And Not unAttended And Not (mi.parent Is fldInbox) Then
                    ' prevent automatic archiving of unread mails in the default inbox folder
                    Utilities.moveItem mi, fld, "ClassifyOneMail, Not fld Is Nothing"
               End If
          End If
          Msg = ""
          If Not topicDic Is Nothing Then
               For i = 0 To topicDic.Count - 1
                    If folderId <> topicDic.Keys(i) Then
                         If Not IsEmpty(topicDic.Items(i)) Then
                              Set tEntry = topicDic.Items(i)
                              Msg = vbCrLf & tEntry.folderPath & vbTab & " " & tEntry.Count '     Application.Session.GetFolderFromID(topicDic.Keys(i)).folderPath & vbTab & " " & topicDic.Items(i)
                         End If
                    End If
               Next i
          End If
          If Not unAttended Then
               If Msg <> "" Then
                    Msg = "Some items matching this are stored elsewhere:" & Msg & _
                         vbCrLf & "Do you want to move them also to " & FolderName & "?"
                    If vbYes = MsgBox(Msg, vbQuestion Or vbYesNo) Then
                         For i = 0 To topicDic.Count - 1
                              trace.trace "ClassifyOneMail", Application.Session.GetFolderFromID(topicDic.Keys(i)).folderPath
                              If folderId <> topicDic.Keys(i) Then
                                   For Each omi In Application.Session.GetFolderFromID(topicDic.Keys(i)).Items
                                        If TypeName(omi) = "MailItem" Then
                                             Set foundMi = omi
                                             'trace.trace "ClassifyOneMail", foundMi.subject
                                             If foundMi.ConversationTopic = mi.ConversationTopic Then
                                                  Utilities.moveItem foundMi, fld, "ClassifyOneMail, items matching are stored elsewhere"
                                             End If
                                        End If
                                        DoEvents
                                   Next omi
                              End If ' <> folder
                         Next i 'folder
                    End If ' vbYes
               End If ' msg <> ""
          End If 'Not unattended
     End If 'Not fld Is Nothing
     ' CheckOneMailUnicity mi, fromEvent, Nothing, fld
     If reindexRequired Then
          StartMakeMailIndex loadfile:=False
     End If
     If Now - LastIndexDump > 1 / 24 * 2 Then 'dump every 2 hours
          LastIndexDump = Now
          mailIndexDump
     End If
     Set ClassifyOneMail = fld
End Function
Private Function safeEntryId(fld As Folder) As String
     If Not (fld Is Nothing) Then safeEntryId = fld.EntryID
End Function
Public Sub ClassifyActiveMail()
     If TypeName(Application.ActiveWindow.CurrentItem) = "MailItem" Then
          ClassifyOneMail Application.ActiveWindow.CurrentItem, False
     End If
End Sub
Public Sub LocateActiveMail()
Dim expl As Explorer
Dim mi As MailItem
     If TypeName(Application.ActiveWindow.CurrentItem) = "MailItem" Then
          Set mi = Application.ActiveWindow.CurrentItem
     Else
          Exit Sub
     End If
     'Debug.Print mi.Parent.folderPath & "\" & mi.Subject
     If Application.Explorers.Count > 0 Then
          Set expl = Application.Explorers(1)
          expl.Activate
          Set expl.CurrentFolder = mi.parent
     Else
          MsgBox "No suitable explorer found to locate the mail", vbCritical
     End If
End Sub

Public Sub ClassifySelectedMail()
Dim sel As Outlook.Selection
Dim fld As Outlook.Folder
Dim origFld As Outlook.Folder
Dim convtopics As Scripting.Dictionary
Dim o As Object
Dim mi As MailItem
On Error Resume Next
     trace.trace "ClassifySelectedMail", TypeName(Application.ActiveWindow)
     trace.trace "ClassifySelectedMail", Err.Number, Err.Description
     Set convtopics = New Scripting.Dictionary
     'convtopics.CompareMode = TextCompare
     trace.trace "ClassifySelectedMail", TypeName(Application.ActiveWindow)
     
     
     For Each o In Application.ActiveExplorer.Selection
          If TypeName(o) = "MailItem" Then
               Set mi = o
               Set origFld = mi.parent
               If convtopics.Exists(mi.ConversationTopic) Then
                    Utilities.moveItem mi, convtopics(mi.ConversationTopic), "ClassifySelectedMail"
               Else
                    Set fld = ClassifyOneMail(mi, False, , , fld)
                    If Not fld Is Nothing Then
                         convtopics.add mi.ConversationTopic, fld
                    End If
               End If
          End If
          DoEvents
     Next o
     If vbNo = MsgBox("Want to search for similar mails in " & origFld.folderPath & "?", vbQuestion Or vbYesNo) Then
          Exit Sub
     End If
     Dim Msg As String, msglist As Collection
     Set msglist = New Collection
     For Each o In origFld.Items
          If TypeName(o) = "MailItem" Then
               Set mi = o
               If convtopics.Exists(mi.ConversationTopic) Then
                    msglist.add mi
                    Msg = Msg & vbCrLf & mi.subject & vbTab & " from " & mi.SenderName & " on " & mi.ReceivedTime
               End If
          End If
          DoEvents
     Next o
     If Msg <> "" Then
          If vbYes = MsgBox(Msg, vbYesNo, "Move also to " & fld.folderPath) Then
               For Each mi In msglist
                    Utilities.moveItem mi, convtopics(mi.ConversationTopic), "ClassifySelectedMail, confirmed"
                    DoEvents
               Next mi
          End If
     End If
End Sub
Public Sub CheckFolderMailsUnicity()
Dim sel As Outlook.Selection
Dim selectedObject As Object, exp As Explorer
On Error GoTo proc_err
GoTo proc
proc_err:
     trace.trace "ERROR", Err.Number & " " & Err.Description & " in CheckFolderMailsUnicity"
     MsgBox Err.Number & " " & Err.Description & " in CheckFolderMailsUnicity", vbCritical
     Exit Sub
     Resume
proc:
     FolderUnicityLog.log.Text = ""
     
     Set exp = Application.ActiveExplorer
     For Each selectedObject In exp.Selection
          If TypeName(selectedObject) = "MailItem" Then
               Set FolderUnicityLog.fld = selectedObject.parent
               FolderUnicityLog.Show
               Exit Sub
          End If
     Next selectedObject
End Sub

Public Sub CheckMailUnicity()
Dim sel As Outlook.Selection
Dim o As Object
     For Each o In Application.Explorers(1).Selection
          If TypeName(o) = "MailItem" Then
               CheckOneMailUnicity o, False
          End If
     Next o
End Sub
Public Sub CheckOneMailUnicity(mi As MailItem, fromEvent As Boolean, Optional log As TextBox = Nothing, Optional fld As Outlook.Folder = Nothing)
     Const timeout As Single = 1 / 24 / 60 / 12
     Dim startTime As Date
     Dim schs() As Outlook.Search
     Dim sch As Outlook.Search
     Dim rsts As Outlook.results
     Dim searchFoldername As String
     Dim refMail As MailItem
     Dim foundObj As Object
     Dim foundMi As MailItem
     Dim refMailSign As MailSignature
     Dim foundMailSign As MailSignature
     Dim i As Integer, si As Integer
     Dim parentFlds() As Outlook.Folder
     Dim parentFld As Outlook.Folder
     Dim parentFldVar As Variant
     Dim ToBeDeleted As Collection
     Dim Msg As String
     Dim DoIt As Boolean
     
     Dim strF As String
On Error GoTo proc_err
GoTo proc
proc_err:
     trace.trace "ERROR", Err.Number & " " & Err.Description & " in CheckOneMailUnicity"
     If log Is Nothing Then
          MsgBox Err.Number & " " & Err.Description & " in CheckOneMailUnicity", vbCritical
     Else
          log.Text = log.Text & vbCrLf & "ERROR " & Err.Number & " " & Err.Description & " in CheckOneMailUnicity"
     End If
     Exit Sub
     Resume
proc:
     
     Set refMail = mi
     Set refMailSign = New MailSignature
     refMailSign.init refMail
     
 If Not log Is Nothing Then
          log.Text = log.Text & vbCrLf & "Checking " & refMailSign.ExtendedSubject
     End If
     If fld Is Nothing Or fld Is mi.parent Then
          ReDim parentFlds(0)
          Set parentFlds(0) = mi.parent
     Else
          ReDim parentFlds(1)
          Set parentFlds(0) = mi.parent
          Set parentFlds(1) = fld
'
     End If
     
     Set ToBeDeleted = New Collection
     Msg = ""
     For si = 0 To UBound(parentFlds)
          'If CheckUnicityMailSearchComp(si) Then
               'Set rsts = schs(si).Results
               For Each foundObj In parentFlds(si).Items
                    If TypeName(foundObj) = "MailItem" Then
                         Set foundMi = foundObj
                         If foundMi.EntryID = refMail.EntryID Then
                              ' ignore
                         ElseIf foundMi.parent Is Application.Session.GetDefaultFolder(olFolderSentMail) Then
                              ' ignore
                         ElseIf foundMi.ConversationTopic = mi.ConversationTopic Then                     ' now use the signature
                              Set foundMailSign = New MailSignature
                              foundMailSign.init foundMi
                              If refMailSign.contains(foundMailSign) Then
                                   Msg = Msg & vbCrLf & refMailSign.ExtendedSubject & " contains " & foundMailSign.ExtendedSubject
                                   ToBeDeleted.add foundMailSign
                              ElseIf refMailSign.isContainedIn(foundMailSign) Then
                                   Msg = Msg & vbCrLf & refMailSign.ExtendedSubject & " is contained into " & foundMailSign.ExtendedSubject
                                   ToBeDeleted.add refMailSign
                                   Set refMail = foundMi
                                   Set refMailSign = foundMailSign
                              Else
                                   'msg = msg & vbCrLf & foundMailSign.ExtendedSubject & " does not fit"
                              End If
                         End If
                    End If
               Next foundObj
          'End If
     Next si
     For Each refMailSign In ToBeDeleted
          Msg = Msg & vbCrLf & "==>Delete: " & refMailSign.ExtendedSubject
     Next refMailSign
     If ToBeDeleted.Count > 0 Then
          DoIt = False
          If log Is Nothing Then
          Else
               log.Text = log.Text & vbCrLf & Msg
          End If
          DoIt = (vbYes = MsgBox(Msg, vbYesNo Or vbQuestion, "Confirm following actions"))
          If DoIt Then
               For Each refMailSign In ToBeDeleted
                    Utilities.moveItem refMailSign.mail, Nothing, "CheckOneMailUnicity"
               Next refMailSign
          End If
     End If
End Sub

Private Sub StartMakeMailIndex(loadfile As Boolean)
Dim FileName As String, filebk As String, filepath As String
Dim fs As New FileSystemObject
Dim ts As TextStream
Dim fileLine As String
Dim topic As String, folderId As String, folderPath As String, scount As String, Count As Long
Dim lineItems As Variant
Dim topicDic As Scripting.Dictionary
Dim cnt As Long, doindex As Boolean

On Error GoTo proc_err
GoTo proc
proc_err:
     Dim ErrN As Long: ErrN = Err.Number
     Dim ErrD As String: ErrD = Err.Description
     trace.trace "ERROR", ErrN & " " & ErrD & " in StartMakeMailIndex"
     MsgBox ErrN & " " & ErrD & " in StartMakeMailIndex", vbCritical
     Exit Sub
     Resume
proc:

On Error GoTo proc_err
     trace.trace "StartMakeMailIndex", "Start"
     IndexingStatus.Show False
     IndexingStatus.LabelPrefix = "Initializing..."
     IndexingStatus.value = 0
     DoEvents
     
     filepath = Environ("tmp") & "\OutlookIndex.txt"
     Set MailIndex = New Scripting.Dictionary
               
          If fs.FileExists(filepath) And loadfile Then
          Set ts = fs.OpenTextFile(filepath, ForReading)
          IndexingStatus.LabelPrefix = "Loading..."
          While Not ts.AtEndOfStream
               cnt = cnt + 1
               fileLine = ts.ReadLine
               lineItems = parseXmlLine(fileLine, "topic", "folderid", "folderpath", "count")
               topic = lineItems(0)
               folderId = lineItems(1)
               folderPath = lineItems(2)
               scount = lineItems(3)
               
               If IsNumeric(scount) Then
                    Count = CLng(scount)
               Else
                    Count = 0
               End If
               IndexingStatus.Setlabel topic
               If Not MailIndex.Exists(topic) Then
                    Set topicDic = New Scripting.Dictionary
                    Set MailIndex(topic) = topicDic
               Else
                    Set topicDic = MailIndex(topic)
               End If
               'folderId = Application.Session.folders(FolderPath).id
               Set topicDic(folderId) = MakeTopicEntry(Count, folderPath)
          Wend
          ts.Close
          IndexingStatus.Hide
          If cnt < 10 Then 'not reasonably possible: there has been an issue with the file.
               doindex = True
          End If
     Else
          doindex = True
     End If
     If doindex Then
          On Error GoTo proc_err
          Dim MailFolder As Outlook.Folder: Set MailFolder = Application.Session.folders("philippe_raemy@swissonline.ch")
          IndexingStatus.MaxValue = CountMailFolders(MailFolder)
          IndexingStatus.LabelPrefix = "Indexing..."
          MakeMailIndex MailFolder
          IndexingStatus.Hide
          mailIndexDump
          trace.trace "StartMakeMailIndex", "Application_Startup DONE"
     End If
     LastIndexDump = Now
End Sub
Private Function CountMailFolders(rootFl As Outlook.Folder) As Long
Dim fl As Outlook.Folder
     IndexingStatus.Setlabel rootFl.folderPath
     If IndexFolder(rootFl.folderPath) Then
          CountMailFolders = 1 + rootFl.Items.Count
     Else
          CountMailFolders = 1
     End If
     For Each fl In rootFl.folders
          CountMailFolders = CountMailFolders + CountMailFolders(fl)
          DoEvents
     Next fl
     
End Function

Private Sub MakeMailIndex(rootFl As Outlook.Folder)
Dim obj As Object
Dim mi As MailItem
Dim topic As String
Dim id As String
Dim fl As Outlook.Folder
Dim topicDic As Scripting.Dictionary
On Error GoTo proc_err
Dim activeCaption As String
GoTo proc
proc_err:
     trace.trace "ERROR", Err.Number & " " & Err.Description & " in MakeMailIndex"
     MsgBox Err.Number & " " & Err.Description & " in MakeMailIndex", vbCritical
     Exit Sub
     Resume
proc:
     For Each fl In rootFl.folders
          MakeMailIndex fl
          DoEvents
     Next fl
     IndexingStatus.Setlabel rootFl.folderPath
     If IndexFolder(rootFl.folderPath) Then
          For Each obj In rootFl.Items
               Set mi = Nothing
               On Error Resume Next
               Set mi = obj
               On Error GoTo proc_err
               If Not mi Is Nothing Then
                    DoEvents
                    topic = mi.ConversationTopic
                    DoEvents
                    id = rootFl.EntryID
                    mailIndexAppend topic, id, rootFl.folderPath
               End If
               IndexingStatus.value = IndexingStatus.value + 1
               DoEvents
          Next obj
     End If
     IndexingStatus.value = IndexingStatus.value + 1
End Sub

Sub mailIndexAppend(miTopic As String, flEntryId As String, folderPath As String)
Dim topicDic As Scripting.Dictionary
Dim tEntry As TopicEntry
'     trace.trace "mailIndexAppend", miTopic
     If MailIndex.Exists(miTopic) Then
          Set topicDic = MailIndex(miTopic)
          If topicDic.Exists(flEntryId) Then
               Set tEntry = topicDic(flEntryId)
               tEntry.increment
               Set topicDic(flEntryId) = tEntry
          Else
               Set topicDic(flEntryId) = MakeTopicEntry(1, folderPath)
          End If
     Else
          Set topicDic = New Scripting.Dictionary
          Set MailIndex(miTopic) = topicDic
          Set topicDic(flEntryId) = MakeTopicEntry(1, folderPath)
     End If
End Sub
Public Sub mailIndexDump()
Dim i As Long, j As Long
Dim FileName As String, filebkname As String, filepath As String, filebkpath As String
Dim fileLine As String
Dim fs As New FileSystemObject
Dim ts As TextStream
Dim topicDic As Scripting.Dictionary
Dim flgReindex As Boolean
Dim tEntry As TopicEntry
GoTo proc
proc_err:
     trace.trace "ERROR", Err.Number & " " & Err.Description & " in mailIndexDump"
     If Err.Number = -2147221238 Then 'The operation cannot be performed because the object has been deleted.
          Resume Next
     End If
     If Err.Number = -2147220991 Then 'The messaging interface has returned an unknown error. If the problem persists, restart Outlook.
          flgReindex = True
          StartMakeMailIndex loadfile:=False
          Exit Sub
     End If
     MsgBox Err.Number & " " & Err.Description & " in mailIndexDump", vbCritical
     On Error Resume Next
     ts.Close
     Exit Sub
     Resume
proc:
     
     On Error GoTo proc_err:
     If MailIndex Is Nothing Then
          'MsgBox "Mail index is empty: not saving", vbCritical
          Exit Sub
     End If
     IndexingStatus.Show False
     IndexingStatus.LabelPrefix = "Dumping..."
     IndexingStatus.value = 0
     
     FileName = "OutlookIndex.txt"
     filepath = Environ("tmp") & "\" & FileName
     filebkname = "OutlookIndex.bk.txt"
     filebkpath = Environ("tmp") & "\" & filebkname
     If fs.FileExists(filepath) Then
          If fs.FileExists(filebkpath) Then
               fs.DeleteFile filebkpath, True
          End If
          fs.GetFile(filepath).name = filebkname
     End If
     Set ts = fs.CreateTextFile(filepath, True, False)

     For i = 0 To MailIndex.Count - 1
          Set topicDic = MailIndex.Items(i)
          IndexingStatus.Setlabel CStr(MailIndex.Keys(i))
          DoEvents
          For j = 0 To topicDic.Count - 1
               If Not IsEmpty(topicDic.Items(j)) Then
                    Set tEntry = topicDic.item(topicDic.Keys(j))
                    fileLine = "<topic>" & MailIndex.Keys(i) & "</topic>" & _
                    "<folderid>" & topicDic.Keys(j) & "</folderid>" & _
                    "<folderpath>" & tEntry.folderPath & "</folderpath>" & _
                    "<count>" & tEntry.Count & "</count>"
                    On Error Resume Next
                    ts.WriteLine fileLine
                    If Err.Number <> 0 Then
                              trace.trace "mailIndexDump", Err.Number & " " & Err.Description & " in mailIndexDump, trying to write"
                              trace.trace "mailIndexDump", fileLine
                    End If
                    On Error GoTo proc_err
                    DoEvents
               End If
          Next j
     Next i
     ts.Close
     IndexingStatus.Hide
End Sub
Private Function IndexFolder(folderPath As String) As Boolean
     Static excluded(4) As String
     Static inited As Boolean
     Dim i As Integer
     If Not inited Then 'initialize
          inited = True
          excluded(0) = "\Deleted Items"
          excluded(1) = "\Inbox"
          excluded(2) = "\Convergence\datafiles"
          excluded(3) = "\Convergence reports"
          excluded(4) = "\Calendar"
     End If
     For i = 0 To UBound(excluded)
          If Right(folderPath, Len(excluded(i))) = excluded(i) Then
               IndexFolder = False
               Exit Function
          End If
     Next i
     IndexFolder = True
End Function

Private Function splitstr(str As String, delim As String _
     , Optional ByRef str0 As String _
     , Optional ByRef str1 As String _
     , Optional ByRef str2 As String _
     , Optional ByRef str3 As String _
     , Optional ByRef str4 As String _
     )
Dim a As Variant
     a = Split(str, delim)
     If UBound(a) >= 0 Then str0 = a(0)
     If UBound(a) >= 1 Then str1 = a(1)
     If UBound(a) >= 2 Then str2 = a(2)
     If UBound(a) >= 3 Then str3 = a(3)
     If UBound(a) >= 4 Then str4 = a(4)
End Function

Public Function parseXmlLine(line As String, ParamArray tags() As Variant) As Variant
Dim s() As String
Dim i As Integer
Dim iStart As Long, iEnd As Long
     On Error Resume Next
     ReDim s(UBound(tags))
     For i = 0 To UBound(tags)
          iStart = InStr(1, line, "<" & tags(i) & ">", vbTextCompare) + Len(tags(i)) + 2
          iEnd = InStr(1, line, "</" & tags(i) & ">", vbTextCompare)
          s(i) = Mid(line, iStart, iEnd - iStart)
     Next i
     parseXmlLine = s
End Function
Public Function TestParseXmlLine() As String
Dim a As Variant, i As Integer
a = parseXmlLine("<abcd>1234</abcd><xml>023987509823745</xml>", "xxx", "xml", "abcd", "ghj")
For i = 0 To UBound(a)
     TestParseXmlLine = TestParseXmlLine & i & "] " & a(i) & vbCrLf
Next i
End Function
Public Sub BackupPst()
     Exit Sub
     Shell "J:\outlook\backupAndRestart.bat"
     Application.Quit
End Sub
Public Sub BackupPst_Old()
     Const LocalPath = "C:\outlook\"
     Const SecuredPath = "J:\outlook\"
     Const FileName = "Personal Folders.pst"
     Const WaitLockSeconds = 60
     Dim myNameSpace As Outlook.NameSpace
     Dim fs As New FileSystemObject
     Dim WaitLockStart As Date
     
On Error GoTo proc_err
GoTo proc
proc_err:
     trace.trace "BackupPst_Old", "ERROR in BackupPst()", Err.Number, Err.Description
     Select Case Err.Number
          Case -2147024863 'Automation error The process cannot access the file because another process has locked a portion of the file.
               If WaitLockStart = 0 Then
                    WaitLockStart = Now
               End If
               If 86400 * (Now - WaitLockStart) > WaitLockSeconds Then
                    MsgBox Err.Number & " " & Err.Description & " in BackupPst", vbCritical
               Else
                    DoEvents
                    Resume
               End If
          Case -2147221233 '     The operation failed. An object could not be found.
               trace.trace "BackupPst_Old", "resume"
               Resume Next
          Case Else
               MsgBox Err.Number & " " & Err.Description & " in BackupPst", vbCritical
     End Select
     Exit Sub
     Resume
proc:
     WaitLockStart = 0
     Set myNameSpace = Application.GetNamespace("MAPI")
     myNameSpace.RemoveStore Application.Session.folders("Archives")
     fs.CopyFile LocalPath & FileName, LocalPath & FileName & ".bk", OverWriteFiles:=True
     Shell "robocopy " & LocalPath & " " & SecuredPath & " *.bk"
     myNameSpace.AddStore LocalPath & FileName
End Sub
Public Sub MaintenanceTasks()
     StartMakeMailIndex loadfile:=False
     BackupPst
End Sub
Private Function MakeTopicEntry(cnt As Long, folderPath As String) As TopicEntry
     Set MakeTopicEntry = New TopicEntry
     MakeTopicEntry.Count = cnt
     MakeTopicEntry.folderPath = folderPath
End Function

'Public Sub JunkSelectedMail()
'Dim exp As Outlook.Explorer
'Dim o As Object, mi As MailItem, fwdMi As MailItem
'     Set exp = Application.ActiveWindow
'     For Each o In exp.Selection
'          If TypeName(o) = "MailItem" Then
'               Set mi = o
'               trace.trace "JunkSelectedMail, ", mi.SenderEmailAddress
'               MakeSpamMail mi
'          End If
'     Next o
'End Sub

'Public Sub JunkSelectedMail_oldversion()
'Dim sel As Outlook.Selection
'Dim exp As Outlook.Explorer
'Dim rul As Outlook.Rule
'Dim o As Object, mi As MailItem, fwdMi As MailItem
'Dim i As Integer
'Dim aList As Variant
'Dim flgAdd As Boolean
'Dim anyAdded As Boolean
'
'     trace.trace "JunkSelectedMail start"
'
'     If TypeName(Application.ActiveWindow) = "Explorer" Then
'
'          Set exp = Application.ActiveWindow
'          Set rul = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Store.GetRules.Item("extra spam")
'
'          For Each o In exp.Selection
'               If TypeName(o) = "MailItem" Then
'                    Set mi = o
'                    trace.trace "JunkSelectedMail, ", mi.SenderEmailAddress
'                    flgAdd = True
'                    aList = rul.Conditions.SenderAddress.Address
'                    If IsArray(aList) Then
'                         For i = 0 To UBound(aList)
'                              If UCase(aList(i)) = UCase(mi.SenderEmailAddress) Then
'                                   MsgBox "Address already in the list of the rule:" & vbCrLf _
'                                        & mi.SenderEmailAddress & vbCrLf _
'                                        & "Check if rule is enabled!", vbCritical
'                                        flgAdd = False
'                                   Exit For
'                              End If
'                         Next i
'                         If flgAdd Then
'                              ReDim Preserve aList(UBound(aList) + 1)
'                              aList(UBound(aList)) = mi.SenderEmailAddress
'                              anyAdded = True
'                         End If
'                    Else
'                         aList = Array(rul.Conditions.SenderAddress.Address, mi.SenderEmailAddress)
'                    End If
'                    Set fwdMi = mi.Forward
'                    trace.trace "JunkSelectedMail, fwd,", mi.SenderEmailAddress
'                    DoEvents
'                    fwdMi.Recipients.Add "Spam, Collection - Collection_Spam@cargill.com"
'                    fwdMi.Send
'                    trace.trace "JunkSelectedMail, sent,", mi.SenderEmailAddress
'                    DoEvents
'                    mi.Delete
'               End If
'          Next o
'          If anyAdded Then
'               rul.Conditions.SenderAddress.Address = aList
'               rul.Parent.save
'               trace.trace "JunkSelectedMail, rule saved"
'               'rul.Execute
'          End If
'     End If
'     trace.trace "JunkSelectedMail, done"
'End Sub
Public Sub exportRules()
Dim rul As Outlook.Rule
Dim ruleC As Outlook.RuleCondition
Dim ruleA As Outlook.RuleAction
Dim s As String
Dim fn As Integer
fn = FreeFile
Open "c:\temp\Outlook rules.txt" For Output As fn
     For Each rul In Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).store.GetRules
          DispText fn, 0, ""
          DispText fn, 0, rul.name
          DispText fn, 0, "-------------------"
          DispText fn, 0, "Conditions:"
          For Each ruleC In rul.Conditions
               If ruleC.Enabled Then
                    DispText fn, 1, TypeName(ruleC)
                    Select Case TypeName(ruleC)
                         Case "AddressRuleCondition"
                              dispArray fn, 2, ruleC.address
                         Case "TextRuleCondition"
                              dispArray fn, 2, ruleC.Text
                         Case "ToOrFromRuleCondition"
                              dispCollection fn, 2, ruleC.Recipients
                         Case "RuleCondition"
                              DispText fn, 2, "No special condition"
                         Case Else
                              DispText fn, 2, "unknown conditions type " & TypeName(ruleC)
                    End Select
               End If
          Next ruleC
          DispText fn, 0, "Actions:"
          For Each ruleA In rul.Actions
               If ruleA.Enabled Then
                    DispText fn, 1, TypeName(ruleA)
                    Select Case TypeName(ruleA)
                         Case "RuleAction"
                              Select Case ruleA.ActionType
                              Case 0
                                   DispText fn, 2, "Unknown"
                              Case 3
                                   DispText fn, 2, "Delete"
                              Case 19
                                   DispText fn, 2, "Mark read"
                              Case 24
                                   DispText fn, 2, "Desktop alert"
                              Case Else
                                   DispText fn, 2, "unknown action type " & TypeName(ruleA) & " type:" & ruleA.ActionType
                              End Select
                         Case "SendRuleAction"
                              dispCollection fn, 2, ruleA.Recipients
                         Case "NewItemAlertRuleAction"
                              DispText fn, 2, ruleA.Text
                         Case "MoveOrCopyRuleAction"
                              DispText fn, 2, ruleA.Folder.folderPath
                         Case "AssignToCategoryRuleAction"
                         
                              dispArray fn, 2, ruleA.Categories
                         Case Else
                              trace.trace "exportRules", "unknown action type ", TypeName(ruleA)
                    End Select
               End If
          Next ruleA
          DispText fn, 0, "-------------------"
     Next rul
     Close #fn
End Sub
Sub DispText(fn As Integer, indent As Integer, t As String)
     Print #fn, t
     trace.trace String(indent, vbTab), t
End Sub
Sub dispArray(fn As Integer, indent As Integer, a As Variant)
Dim i As Integer
          For i = LBound(a) To UBound(a)
               DispText fn, indent, CStr(a(i))
          Next i
End Sub
Sub dispCollection(fn As Integer, indent As Integer, a As Variant)
Dim i As Integer
          For i = 1 To a.Count
               DispText fn, indent, CStr(a(i))
          Next i
End Sub


Public Sub ClassifyReadReceipts()
Dim mapi As NameSpace
Dim fldMail As Outlook.Folder
Dim fldInbox As Outlook.Folder
Dim fldReceipts As Outlook.Folder
Dim fld As Outlook.Folder
Dim mi As MailItem
Dim ri As ReportItem
Dim mei As MeetingItem
Dim obj As Object
Dim word As Variant
Dim dic As Scripting.Dictionary
Dim i As Integer
Set dic = New Scripting.Dictionary
dic.CompareMode = TextCompare

On Error GoTo proc_err
GoTo proc
proc_err:
     trace.trace "ERROR", Err.Number & " " & Err.Description & " in ClassifyReadReceipts"
     MsgBox Err.Number & " " & Err.Description & " in ClassifyReadReceipts", vbCritical
     Exit Sub
     Resume
     Exit Sub
proc:

     Set mapi = Application.GetNamespace("Mapi")
     For Each fld In mapi.folders
          If LCase(fld.name) Like "mailbox*" _
               And Not LCase(fld.name) Like "*shared*" _
               And Not LCase(fld.name) Like "*releasemanagement*" Then
               Set fldMail = fld
               Exit For
          End If
     Next fld
     If fldMail Is Nothing Then
          MsgBox "Could not find mailbox folder.", vbCritical
     End If
     Set fldInbox = fldMail.folders("Inbox")
     Set fldReceipts = fldMail.folders("_Receipts")
     
     For i = fldInbox.Items.Count To 1 Step -1
          Set obj = fldInbox.Items(i)
          Select Case TypeName(obj)
               Case "MailItem"
'                    Set mi = obj
'                    'trace.trace mi.SenderName, mi.subject
'                    For Each word In Split(mi.subject)
'                         If dic.Exists(word) Then
'                              dic(word) = dic(word) + 1
'                         Else
'                              dic(word) = 1
'                         End If
'                    Next word
               Case "ReportItem"
                    Set ri = obj
                    Utilities.moveItem ri, fldReceipts, "ClassifyReadReceipts, ReportItem"
               Case "MeetingItem"
                    Set mei = obj
                    If Not mei.UnRead Then
                         Utilities.moveItem mei, fldReceipts, "ClassifyReadReceipts, unread meeting item"
                    End If
               Case Else
                     trace.trace "ClassifyReadReceipts", "unkown type", TypeName(obj)
          End Select
          DoEvents
     Next i
     'trace.trace "ClassifyReadReceipts", "-----------------------------------------"
     'For i = 0 To dic.Count - 1
     '     trace.trace "ClassifyReadReceipts", dic.Key(i), dic.Item(i)
     'Next i
End Sub

Sub ClassifyFromSearchFoldersProc(destFld As Outlook.Folder)
          Dim colStores As Outlook.Stores
          Dim oStore As Outlook.store
          Dim oSearchFolders As Outlook.folders
          Dim oFolder As Outlook.Folder
          Dim DestSubFld As Outlook.Folder
          Dim DestSubFldName As String
          Dim itm As Outlook.MailItem
'          Dim bs As ButtonStatus
'          Dim folderNames As Variant
          Dim Items() As Variant, i As Integer
          
On Error GoTo proc_err
GoTo proc
proc_err:
    Dim Msg As String
    Msg = Err.Number & " " & Err.Description & " in ClassifyFromSearchFoldersProc"
    trace.trace "ERROR", Msg
    MsgBox Msg, vbCritical
    Exit Sub
    Resume
    Exit Sub
proc:
    
    ' ClassifyReadReceipts
    
    Set colStores = Application.Session.Stores
    Dim s As Integer
    For s = 1 To colStores.Count
        Set oStore = Nothing
        On Error Resume Next
        Set oStore = colStores(s)
        On Error GoTo proc_err
        Dim folders() As String
        ReDim folders(0)
        If Not (oStore Is Nothing) Then
            If oStore.ExchangeStoreType = olPrimaryExchangeMailbox Or oStore.DisplayName Like "*@*" Then
                Set oSearchFolders = oStore.GetSearchFolders
                For Each oFolder In oSearchFolders
                    If Not oFolder.name Like "Search Folder*" _
                    And Not oFolder.name Like "MS-OLK*" _
                    And Not Left(oFolder.name, 1) = "*" _
                    Then
                        ReDim Preserve folders(UBound(folders) + 1)
                        folders(UBound(folders)) = oFolder.name
                    End If
                Next oFolder
                Dim f As Integer
                Utilities.BubbleSort folders
                For f = 1 To UBound(folders)
                    Set oFolder = oSearchFolders.item(folders(f))
                    trace.trace "ClassifyFromSearchFoldersProc", oFolder.folderPath
                    DestSubFldName = oFolder.name
                    Dim ReadItemsOnly As Boolean
                    ReadItemsOnly = LCase(DestSubFldName) Like "*{read}*"
                    If InStr(1, DestSubFldName, "{", vbBinaryCompare) > 0 Then
                        DestSubFldName = Mid(DestSubFldName, 1, InStr(1, DestSubFldName, "{", vbBinaryCompare) - 1)
                    End If
                    Set DestSubFld = Utilities.EnsureFolderExists(destFld, DestSubFldName)
                    If oFolder.Items.Count > 0 Then
                        ReDim Items(1 To oFolder.Items.Count)
                        For i = 1 To oFolder.Items.Count: Set Items(i) = oFolder.Items(i): Next i
                        For i = LBound(Items) To UBound(Items)
                            Dim DoIt As Boolean
                            If TypeName(Items(i)) = "MailItem" Then
                                Set itm = Items(i)
                                DoIt = Not ReadItemsOnly Or Not itm.UnRead
                            Else
                                DoIt = True
                            End If
                            If DoIt Then
                                Utilities.moveItem Items(i), DestSubFld, "ClassifyFromSearchFoldersProc"
                            End If
                            DoEvents
                        Next i
                    End If
                Next f
            End If
        End If
    Next s
End Sub
Public Sub HistorizeMails()
    HistorizeMailsProc "all"
End Sub
Public Sub ClassifyFromSearchFolders()
    HistorizeMailsProc "fromsearch"
End Sub
Public Sub resetButton()
     Dim caption As String
     Dim ncaption As String
     Dim bs As New ButtonStatus
     caption = InputBox("Please type in the caption of the button to be reset")
     ncaption = InputBox("Please type in the new or default caption")
     If caption = "" Then Exit Sub
     bs.init "customs", caption, exactMatch:=False
     bs.setDefaultCaption ncaption
     Set bs = Nothing
End Sub
Public Sub HistorizeMailsProc(action As String)

'inbox older than 1 mth
'inbox subfolders older than 8 months
'Various older than 15 days
'ShortTerm older than 15 days, then deleted
'VeryShortTerm older than 7 days, then deleted
'Some obsolete folders right now.


Dim fldInbox As Outlook.Folder
Dim fldPers As Outlook.Folder
Dim fldPersRoot As Outlook.Folder
Dim fldArch As Outlook.Folder
Dim mapi As NameSpace
Dim obj As Object
Dim InboxRetentionDate As Date
Dim CalendarRetentionDate As Date
Dim FoldersRetentionDate As Date
Dim ShortTermRetentionDate As Date
Dim VeryShortTermRetentionDate As Date
Dim ArchivesRetentionDate As Date
Dim ObsoleteFoldersRetentionDate As Date
Dim ObsoleteFolders As Variant
Dim subIbFld As Outlook.Folder
Dim subPerFld As Outlook.Folder
Dim fld As Outlook.Folder
Dim i As Integer
Dim Msg As String
Dim bs As ButtonStatus
Dim flgArchiveUnread As Boolean

On Error GoTo proc_err
GoTo proc
proc_err:
     Msg = Err.Number & " " & Err.Description & " in HistorizeMailsProc"
     trace.trace "ERROR", Msg
     MsgBox Msg, vbCritical
     Exit Sub
     Resume
     Exit Sub
proc:

'     Set bs = New ButtonStatus
'     bs.init "customs", "Historize"

     trace.trace "HistorizeMailsProc(""" & action & """)"
     
     
     Set mapi = Application.GetNamespace("Mapi")
     Select Case action
          Case Else
               InboxRetentionDate = DateAdd("d", -1000, Now())
               FoldersRetentionDate = DateAdd("m", 0, Now())
               ShortTermRetentionDate = DateAdd("d", -15, Now())
               VeryShortTermRetentionDate = DateAdd("d", -7, Now())
               ArchivesRetentionDate = DateAdd("d", -180, Now())
               CalendarRetentionDate = DateAdd("d", -90, Now())
          
               ObsoleteFoldersRetentionDate = Now()
               ObsoleteFolders = Array() '"Audit")
               Set fldPersRoot = mapi.folders("philippe_raemy@swissonline.ch")
               Set fldPers = mapi.folders("philippe_raemy@swissonline.ch").folders("Archives")
               Set fldInbox = Application.Session.GetDefaultFolder(olFolderInbox)
               
               flgArchiveUnread = False
     End Select
     
     Select Case action
          Case "all", "fromsearch"
               ' bs.SetCaption "Classify From Search Folders"
               ClassifyFromSearchFoldersProc fldPers '     Application.Session.GetDefaultFolder(olFolderInbox).Parent
     End Select
     Select Case action
          Case "all", "historize", "Cargill Archives"
          ' bs.SetCaption "Make mail index"
          StartMakeMailIndex loadfile:=False
     
          ' bs.SetCaption "Classify mails in Inbox"
          For i = fldInbox.Items.Count To 1 Step -1
               Set obj = fldInbox.Items(i)
               If TypeName(obj) = "MailItem" Then
                    Set fld = ClassifyOneMail(obj, False, unAttended:=True)
                    'If fld Is Nothing Then HandleIncomingMails obj
               End If
               DoEvents
          Next i
          ' bs.SetCaption "Historize mails in Inbox"
          'For i = fldInbox.Items.Count To 1 Step -1
          '     Set obj = fldInbox.Items(i)
          '     Select Case TypeName(obj)
          '          Case "MailItem", "MeetingItem"
          '               If obj.SentOn < InboxRetentionDate And (Not obj.UnRead Or flgArchiveUnread) Then
          '                    trace.trace "HistorizeMailsProc", obj.SentOn & "," & obj.subject & "-->" & fldPers.folderPath
          '                    Utilities.moveItem obj, fldPers, "HistorizeMailsProc(" & action & ")"
          '                    DoEvents
          '               End If
          '     End Select
          'Next i
          For Each subIbFld In fldInbox.parent.folders
               ' bs.SetCaption "Historize mails in " & subIbFld.name
               Set subPerFld = Utilities.EnsureFolderExists(fldPers, subIbFld.name)
               Select Case LCase(subIbFld.name)
                    Case "shortterm"
                         RecurseHistorizeMails subIbFld, Nothing, ShortTermRetentionDate, bs
                    Case "veryshortterm"
                         RecurseHistorizeMails subIbFld, Nothing, VeryShortTermRetentionDate, bs
                    'Case "archivesretentiondate"
                    '     RecurseHistorizeMails subIbFld, subPerFld, InboxRetentionDate, bs
                    'Case Else
                    '     RecurseHistorizeMails subIbFld, subPerFld, FoldersRetentionDate, bs
               End Select
          Next subIbFld
          Select Case action
               Case "Cargill Archives"
               Case Else
               ' RecurseHistorizeMails Application.Session.GetDefaultFolder(olFolderInbox).Parent.folders("_Receipts"), Nothing, ShortTermRetentionDate _
                    , bs, deleteEmptyFolders:=True
               RecurseHistorizeMails Application.Session.GetDefaultFolder(olFolderCalendar), Utilities.EnsureFolderExists(fldPersRoot, "Calendar"), CalendarRetentionDate _
                    , bs, deleteEmptyFolders:=True
               'RecurseHistorizeMails mapi.folders("Archives"), mapi.folders("Personal Zarchives"), ArchivesRetentionDate _
                , bs, deleteEmptyFolders:=True, specialFolders:=ObsoleteFolders, specialFoldersRetentionDate:=ObsoleteFoldersRetentionDate
               ' RecurseHistorizeMails fldInbox.folders("Deleted Items"), Nothing, ShortTermRetentionDate _
                    , bs, deleteEmptyFolders:=True
               RecurseHistorizeMails Application.Session.GetDefaultFolder(olFolderDeletedItems), Nothing, ShortTermRetentionDate _
                    , bs, deleteEmptyFolders:=True
               'RecurseHistorizeMails mapi.folders("Archives").folders("Deleted Items"), Nothing, Now _
               '     , bs, deleteEmptyFolders:=True
               ' bs.SetCaption "Make mail index"
               StartMakeMailIndex loadfile:=False
          End Select
     End Select
     trace.trace "HistorizeMailsProc(""" & action & """): done"
     
End Sub
Private Sub RecurseHistorizeMails( _
     fldInbox As Outlook.Folder, _
     fldPers As Outlook.Folder, _
     retentionDate As Date, _
     bs As ButtonStatus, _
     Optional deleteEmptyFolders As Boolean, _
     Optional specialFolders As Variant, _
     Optional specialFoldersRetentionDate As Date = CDate("1900-01-01"), _
     Optional MinMoveDate As Date = CDate("1900-01-01") _
)
Dim subIbFld As Outlook.Folder
Dim subPerFld As Outlook.Folder
Dim moveToFld As Boolean
Dim i As Integer, j As Integer
Dim recursed As Boolean
Dim obj As Object
Dim Msg As String

On Error GoTo proc_err
GoTo proc
proc_err:
     Msg = Err.Number & " " & Err.Description & " in RecurseHistorizeMails"
     trace.trace "ERROR", Msg
     MsgBox Msg, vbCritical
     Exit Sub
     Resume
     Exit Sub
proc:

     Msg = "RecurseHistorizeMails(" & fldInbox.folderPath
     If fldPers Is Nothing Then
          Msg = Msg & ", Delete"
     Else
          Msg = Msg & "," & fldPers.folderPath
     End If
     Msg = Msg & "," & retentionDate & ")"
     bs.SetCaption Msg
     trace.trace Msg
     Debug.Assert fldInbox.name <> "Commodities Daily"
     
     For i = fldInbox.Items.Count To 1 Step -1
          Set obj = fldInbox.Items(i)
          moveToFld = False
          Select Case TypeName(obj)
               Case "MailItem", "MeetingItem", "PostItem", "SharingItem"
                    If (obj.SentOn < retentionDate Or obj.SentOn > DateAdd("y", 2000, Now)) _
                    And obj.SentOn > MinMoveDate _
                    Then
                         moveToFld = True
                    ElseIf obj.SentOn > Now Then
                         trace.trace "RecurseHistorizeMails", obj.SentOn & "," & obj.subject & "<--" & fldInbox.folderPath
                         Debug.Assert False
                    End If
               Case "AppointmentItem"
                    If Not obj.IsRecurring _
                    And (obj.Start < retentionDate Or obj.Start > DateAdd("y", 2000, Now)) _
                    And obj.Attachments.Count > 0 _
                    And obj.Start > MinMoveDate _
                    Then
                         moveToFld = True
                    End If
               Case "ReportItem"
                         moveToFld = True
               Case Else
                    Debug.Print TypeName(obj)
                    'Debug.Assert False
          End Select
          If moveToFld Then
               Utilities.moveItem obj, fldPers, "RecurseHistorizeMails(retentionDate=" & retentionDate & ")"
          End If
          DoEvents
     Next i
     For i = fldInbox.folders.Count To 1 Step -1
          Set subIbFld = fldInbox.folders(i)
          If Not fldPers Is Nothing Then
               Set subPerFld = Nothing
               On Error Resume Next
               Set subPerFld = fldPers.folders(subIbFld.name)
               On Error GoTo 0
               If subPerFld Is Nothing Then
                    Set subPerFld = fldPers.folders.add(subIbFld.name)
               End If
          End If
          recursed = False
          If IsArray(specialFolders) Then
               For j = LBound(specialFolders) To UBound(specialFolders)
                    If Not recursed And LCase(subIbFld.folderPath) Like LCase("*" & specialFolders(j)) Then
                         RecurseHistorizeMails subIbFld, subPerFld, specialFoldersRetentionDate, bs, deleteEmptyFolders
                         recursed = True
                    End If
               Next j
          End If
          If Not recursed Then
               RecurseHistorizeMails subIbFld, subPerFld, retentionDate, bs, deleteEmptyFolders _
               , specialFolders, specialFoldersRetentionDate
          End If
          DoEvents
     Next i
     If deleteEmptyFolders And fldInbox.folders.Count = 0 And fldInbox.Items.Count = 0 And Not fldInbox.name = "Deleted Items" Then
          trace.trace "RecurseHistorizeMails", "Deleting folder «" & fldInbox.folderPath & "»"
          On Error Resume Next
          fldInbox.delete
          On Error GoTo proc_err
     End If
     trace.trace "RecurseHistorizeMails", "(" & fldInbox.folderPath & "): <=== done"
End Sub

Public Sub oneshotrecoverfromarchive()
     RecurseHistorizeMails _
          fldInbox:=Application.GetNamespace("Mapi").folders("philippe_raemy@swissonline.ch"), _
          fldPers:=Application.GetNamespace("Mapi").folders("Archives"), _
          retentionDate:=CDate("1.1.2012"), _
          bs:=New ButtonStatus
'      RecurseHistorizeMails _
          fldInbox:=Application.GetNamespace("Mapi").folders("Archives"), _
          fldPers:=Application.GetNamespace("Mapi").folders("philippe_raemy@swissonline.ch"), _
          retentionDate:=Now + 1, _
          bs:=New ButtonStatus, _
          MinMoveDate:=CDate("1.1.2012")
End Sub

Sub RemoveAllItemsAndFoldersInFolder(fld As Outlook.Folder)
    Dim oFolders As Outlook.folders
    Dim oItems As Outlook.Items
    Dim i As Long
    'Obtain a reference to deleted items folder
    Set oItems = fld.Items
    For i = oItems.Count To 1 Step -1
        oItems.item(i).delete
    Next
    Set oFolders = fld.folders
    For i = oFolders.Count To 1 Step -1
        oFolders.item(i).delete
    Next
End Sub


Sub EnumerateSearchFoldersInStores()
    Dim colStores As Outlook.Stores
    Dim oStore As Outlook.store
    Dim oSearchFolders As Outlook.folders
    Dim oFolder As Outlook.Folder
    
    On Error Resume Next
    Set colStores = Application.Session.Stores
    For Each oStore In colStores
        Set oSearchFolders = oStore.GetSearchFolders
        For Each oFolder In oSearchFolders
            trace.trace "EnumerateSearchFoldersInStores", oFolder.folderPath
        Next
    Next
End Sub

Sub ForwardAllMeetingItemsFromInbox()
Dim fldInbox As Folder: Set fldInbox = Application.Session.GetDefaultFolder(olFolderInbox)
Dim mObj As Object
For Each mObj In fldInbox.Items
     Debug.Print TypeName(mObj),
     Select Case TypeName(mObj)
          Case "MailItem": Debug.Print mObj.subject
          Case "MeetingItem":
               ForwardMeetingItem mObj
          Case Else
               Debug.Print
     End Select
Next mObj
End Sub

Sub ForwardMeetingItem(mi As Outlook.MeetingItem)
     Dim fmi As Outlook.MeetingItem
     Debug.Print mi.subject
     Select Case mi.Class
          Case OlObjectClass.olMeetingRequest, OlObjectClass.olMeetingCancellation
               Set fmi = mi.Forward
               fmi.Recipients.add "philippe_raemy@swissonline.ch"
               fmi.Send
     End Select
End Sub


Sub ForwardAllMeetingItemsFromCalendar()
Dim fldInbox As Folder: Set fldInbox = Application.Session.GetDefaultFolder(olFolderCalendar)
Dim mObj As Object
For Each mObj In fldInbox.Items
     Select Case TypeName(mObj)
          Case "AppointmentItem"
               ForwardAppointmentItem mObj
          Case Else
     End Select
Next mObj
End Sub

Sub ForwardAppointmentItem(ai As Outlook.AppointmentItem)
     Dim fai As Outlook.AppointmentItem
     
     Select Case ai.Class
          Case OlObjectClass.olAppointment
               If ai.Start >= Now Then
                    Set fai = ai.Copy
                    While fai.Recipients.Count > 0: fai.Recipients.Remove (1): Wend
                    
                    fai.Recipients.add "praemy61@gmail.com"
                    Debug.Print "Sent: " & fai.subject
                    disableAutoClassify = True
                    fai.Send
                    fai.delete
                    disableAutoClassify = False
               End If
     End Select
End Sub

Private Sub GlobFldInbox_BeforeItemMove(ByVal item As Object, ByVal MoveTo As Outlook.Folder, Cancel As Boolean)
     Dim subject As String
     Select Case TypeName(item)
          Case "MailItem", "AppointmentItem": subject = item.subject
          Case Else: subject = TypeName(item)
     End Select
     ' Debug.Print "GlobFldInbox_BeforeItemMove", TypeName(item), MoveTo.name, subject
End Sub

Public Function FindMailsNotForMe() As Variant
Dim mails() As MailItem, idx As Integer
Dim miObj As Object
Dim mi As MailItem, re As Recipient, forMe As Boolean
Const mailAddress = "philippe_raemy@swissonline.ch"
ReDim mails(50)
     
     For Each miObj In Application.Session.GetDefaultFolder(olFolderInbox).Items
          If TypeName(miObj) = "MailItem" Then
               Set mi = miObj
               
               forMe = (mi.SenderEmailAddress = mailAddress)
               If Not forMe Then
                    For Each re In mi.Recipients
                         If LCase(re.address) = mailAddress Then
                              forMe = True
                              Exit For
                         End If
                    Next re
               End If
               If Not forMe Then
                    Debug.Print "----------------" & vbCrLf & "From:" & mi.sender.address & vbCrLf & "to:" & mi.To & vbCrLf & "cc:" & mi.CC & vbCrLf & "bcc:" & mi.BCC & vbCrLf & mi.subject
                    If idx > UBound(mails) Then
                         ReDim Preserve mails(idx + 50)
                    End If
                    Set mails(idx) = mi
                    idx = idx + 1
               End If
          End If
     Next miObj
     If idx > 0 Then
          ReDim Preserve mails(idx - 1)
          FindMailsNotForMe = mails
     Else
          FindMailsNotForMe = Array()
     End If
End Function


Public Sub HandleMailsNotForMe()
Dim mails
mails = FindMailsNotForMe()
Debug.Print UBound(mails)
End Sub

Public Sub BlockSender()
    MarkSenderImpl "Indésirables", 0
End Sub
Public Sub BlockSenderAnyDomain()
    MarkSenderImpl "Indésirables", 1
End Sub
Public Sub BlockSenderDomain()
    MarkSenderImpl "Indésirables", 2
End Sub
Private Sub MarkSenderImpl(Category As String, WholeDomain As Integer)
Dim o As Object, mi As MailItem, address As String
    Dim fso As New FileSystemObject, ts As TextStream
    Dim senders As Variant, sender As Variant, skip As Boolean
    Set ts = fso.OpenTextFile(Environ("LOCALAPPDATA") & "\email\" & Category & ".txt", ForReading)
    senders = Split(ts.ReadAll, vbCrLf)
    ts.Close
    Set ts = fso.OpenTextFile(Environ("LOCALAPPDATA") & "\email\" & Category & ".txt", ForAppending, True)
    For Each o In Application.ActiveExplorer.Selection
        If TypeName(o) = "MailItem" Then
            Set mi = o
            Select Case WholeDomain
                Case 0: address = mi.SenderEmailAddress
                Case 1: address = Split(mi.SenderEmailAddress, "@")(0) & "*"
                Case 2: address = "*" & Split(mi.SenderEmailAddress, "@")(1)
            End Select
            address = InputBox("Do you want to block this sender?", "Sender blocking", address)
            If address = "" Or address = "*" Then
                MsgBox "No sender blocked", vbCritical
                Exit Sub
            End If
            skip = False
            
            For Each sender In senders
                If address Like sender Or sender Like address Then
                    skip = True
                End If
            Next sender
            If Not skip Then
                ReDim Preserve senders(UBound(senders) + 1)
                senders(UBound(senders)) = address
                ts.WriteLine address
            End If
            AutoCategorizeSenderMails address
        End If
    Next o
    ts.Close
End Sub
Public Function AutoCategorizeSenderMails(Optional addresses As String, Optional rmi As MailItem) As Integer
' returns the # of mails marked as undesirable
    
On Error GoTo proc_err
GoTo proc
proc_err:
    Dim emsg As String
    emsg = Err.Number & " " & Err.Description & " in ClassifyOneMail"
    trace.trace "ERROR", emsg
    MsgBox emsg, vbCritical
    Exit Function
    Resume
    Exit Function
proc:
    
    Dim fso As New FileSystemObject, ts As TextStream
    
    Const Categories = "Indésirables"
    Const NotForMeCategory = "Not For Me"
    Const ScumbagCategory = "Scumbag"
    If addresses = "" Then
        Set ts = fso.OpenTextFile(Environ("LOCALAPPDATA") & "\email\Indésirables.txt", ForReading)
        addresses = ts.ReadAll
        ts.Close
    End If
    
    Dim senders As Variant, sender As Variant
    senders = Split(addresses, vbCrLf)
    
    Dim omi As Object
  
    If Not rmi Is Nothing Then
        AutoCategorizeSenderMails = AutoCategorizeSenderOneMail(senders, rmi, Categories, NotForMeCategory, ScumbagCategory)
    Else
        For Each omi In Application.Session.folders("philippe_raemy@swissonline.ch").folders("Inbox").Items
            If TypeName(omi) = "MailItem" Then
                AutoCategorizeSenderMails = AutoCategorizeSenderMails + AutoCategorizeSenderOneMail(senders, omi, Categories, NotForMeCategory, ScumbagCategory)
            End If
        Next omi
    End If
End Function
Private Function AutoCategorizeSenderOneMail(senders As Variant, mi As MailItem, Category As String, NotForMeCategory As String, ScumbagCategory As String) As Integer
Dim sender As Variant
Dim i As Integer
Dim forMe As Boolean
Dim modified As Boolean

      
On Error GoTo proc_err
GoTo proc
proc_err:
    Dim emsg As String
    emsg = Err.Number & " " & Err.Description & " in ClassifyOneMail"
    trace.trace "ERROR", emsg
    MsgBox emsg, vbCritical
    Exit Function
    Resume
    Exit Function
proc:
    Dim principalName As String
    
    ' Debug.Print mi.SenderEmailAddress, mi.subject,
    If mi.SenderEmailAddress = PrincipalMail Then
    ElseIf mi.SenderName = PrincipalMail And Not mi.SenderEmailAddress = PrincipalMail Then
        modified = AddCategory(mi, ScumbagCategory)
    Else
        principalName = Split(PrincipalMail, "@")(0)
        If mi.SenderEmailAddress Like principalName & "*" Then
            modified = AddCategory(mi, ScumbagCategory)
        End If
    End If
    If Not forMe Then
        For i = 1 To mi.Recipients.Count
            If LCase(mi.Recipients(i).address) = PrincipalMail Then
                forMe = True
                Exit For
            End If
        Next i
    End If
    If Not forMe Then
        modified = AddCategory(mi, NotForMeCategory)
    End If
    
    For Each sender In senders
        Debug.Print sender
        sender = UCase(sender)
        Debug.Assert Not sender Like "*qoqa*"
        If UCase(mi.SenderEmailAddress) Like sender _
        Or UCase(mi.SenderName) Like sender _
        Then
            modified = AddCategory(mi, Category)
            AutoCategorizeSenderOneMail = 1
            Exit For
        End If
    Next sender
    If modified Then mi.save
    ' Debug.Print forMe, mi.Categories
End Function
Private Function AddCategory(mi As MailItem, Category As String) As Boolean
    If Not mi.Categories Like "*" & Category & "*" Then
        mi.Categories = mi.Categories & ", " & Category
        AddCategory = True
    End If
End Function

Public Sub AutoCategorizeSenderMailsFromInbox()
    AutoCategorizeSenderMails
End Sub

Public Function GetFolder(FolderName As String) As Outlook.Folder
Dim sto As store
Dim s As Integer
    For s = 1 To Application.Session.Stores.Count
        On Error Resume Next
        Set sto = Application.Session.Stores(s)
        If Err.Number = 0 Then
            On Error GoTo 0
            If sto.DisplayName = FolderName Then
                Set GetFolder = sto.GetRootFolder
                Exit Function
            End If
        End If
    Next s
End Function

Public Sub CmdDuplexScan_Click()
    CmdScan_Click "Duplex"
End Sub

Public Sub CmdGlassScan_Click()
    CmdScan_Click "Glass"
End Sub

Public Sub CmdSimplexScan_Click()
    CmdScan_Click "Simplex"
End Sub

Private Sub CmdScan_Click(profile As String)
    On Error GoTo proc_err
    GoTo proc:

    Dim defaultValues As String

proc_err:
    MsgBox Err.Description
    Exit Sub
    Resume
proc:

    Dim FileName As String
    FileName = Scanner.Scan(profile, _
        DestinationFolder:=Environ("temp"), _
        NamePattern:="outlook.scan.pdf", _
        OpenFileWhenDone:=True, _
        ReplaceIfExist:=False, _
        defaults:="", _
        EditFileName:=True _
    )
    Debug.Print FileName
    Dim mi As MailItem, harry As Inspector
    Set harry = Application.ActiveInspector
    If harry Is Nothing Then
        Err.Raise vbObjectError, "vba", "No active inspector object"
    Else
        Set mi = harry.CurrentItem
        If mi.Sent Then
            Err.Raise vbObjectError, "vba", "The current mail item cannot be modified"
        Else
            mi.Attachments.add FileName
        End If
    End If
End Sub


Private Sub Items_FolderAdd(ByVal Folder As MAPIFolder)

End Sub


