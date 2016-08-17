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
Public ClassifyMailSearchComp As Boolean
Private CheckUnicityMailSearchComp(1) As Boolean
'mailIndex is a dictionnary of dictionnaries.
'1st level indexed with conversation topics
'2nd level indexed with folder-ids and containing TopicEntry objects (folderpath + count of mails)
Private MailIndex As Scripting.Dictionary
Private LastIndexDump As Date


Private Sub Application_Startup()
    StartMakeMailIndex True
End Sub
Private Sub Application_Quit()
    mailIndexDump
End Sub

Private Sub Application_ItemSend(ByVal item As Object, cancel As Boolean)
Dim mi As MailItem
    Select Case TypeName(item)
        Case "MailItem"
            ClassifyMail item, cancel
    End Select
End Sub

Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Search)
    trace.trace "AdvancedSearchComplete Event", "tag=" & SearchObject.Tag & ", search=" & SearchObject.Filter
    Select Case SearchObject.Tag
        Case "AutoClassifyMail": ClassifyMailSearchComp = True
        Case "CheckMailUnicity_0": CheckUnicityMailSearchComp(0) = True
        Case "CheckMailUnicity_1": CheckUnicityMailSearchComp(1) = True
    End Select
End Sub

Public Function GetInboxFolder() As MAPIFolder
    Set GetInboxFolder = Application.Session.GetDefaultFolder(olFolderInbox)
End Function

Public Function GetArchivesFolder() As MAPIFolder
    Set GetArchivesFolder = Application.Session.folders("Philippe.Raemy@kochind.com").folders("Archives")
End Function

Private Sub ClassifyMail(ByRef mi As MailItem, ByRef cancel As Boolean)
    If mi.To = "Spam, Collection - Collection_Spam@cargill.com" Then
    Else
        ClassifyOneMail mi, True, cancel
    End If
End Sub
Private Function ClassifyOneMail(ByRef mi As MailItem, ByVal fromEvent As Boolean, Optional ByRef cancel As Boolean = False, Optional ByVal unAttended As Boolean = False, Optional ByVal suggestedFld As Folder = Nothing) As MAPIFolder
    Dim reindexRequired As Boolean
    Dim emsg    As String
On Error GoTo proc_err
GoTo proc
proc_err:
    If Err.Number = -2147220991 Then ' The messaging interface has returned an unknown error. If the problem persists, restart Outlook.
        reindexRequired = True
        Resume Next
    End If
    MsgBox trace.trace("Error", Err.Number & " " & Err.Description & " in ClassifyOneMail"), vbCritical
    Exit Function
    Resume
    cancel = True
    Exit Function
proc:
    
    Dim sch As Outlook.Search
    Dim rsts As Outlook.Results
    Dim i As Integer, maxFound As Integer
    'timeout is 5 seconds
    Const timeout As Single = 1 / 24 / 60 / 12
    
    
    Dim folders As Scripting.Dictionary
    Dim foldersCount As Scripting.Dictionary
    Set folders = New Dictionary
    Set foldersCount = New Dictionary
    Dim folderName As String
    Dim fld As MAPIFolder
    Dim fldParent As MAPIFolder
    Dim startTime As Date
    Dim msg As String
    Dim folderId As String
    Dim folderIdKey As String
    Dim fldInbox As Folder: Set fldInbox = GetInboxFolder()
    
    
    Dim foundMi As MailItem
    Dim oMi As Object
    Dim copyMi As MailItem
    Dim tEntry As TopicEntry
    
    Dim topic As String
    Dim topicDic As Scripting.Dictionary
    
    topic = mi.ConversationTopic
    msg = ""
    If MailIndex Is Nothing Then
        Set MailIndex = New Scripting.Dictionary
        StartMakeMailIndex loadfile:=True
    End If
        
    If MailIndex.Exists(topic) Then
        Set topicDic = MailIndex(topic)
        maxFound = 0
        
        For i = 0 To topicDic.Count - 1
            If Not IsEmpty(topicDic.items(i)) Then
                Set tEntry = topicDic.items(i)
                If tEntry.Count > maxFound Then
                    maxFound = tEntry.Count
                    folderId = topicDic.Keys(i)
                End If
            End If
        Next i
        On Error Resume Next
        Set fld = Application.Session.GetFolderFromID(folderId)
        If Err.Number = -2147220991 Then '    The messaging interface has returned an unknown error. If the problem persists, restart Outlook.
            On Error GoTo proc_err
            If unAttended Then
                trace.trace "ClassifyOneMail", "folder «" & folderId & "» not found!!!"
                Exit Function '---------------------------------------------------->>>>>>>>>>>
            End If
            Set fld = Application.Session.PickFolder
            folderId = fld.EntryID
            reindexRequired = True
        ElseIf Err.Number = 0 Then
            On Error GoTo proc_err
            msg = "Save to folder «" & fld.folderPath & "»?"
            If unAttended Then
                trace.trace "ClassifyOneMail", msg
            ElseIf safeEntryId(fld) = safeEntryId(suggestedFld) Then
                trace.trace "ClassifyOneMail", msg
            Else
                Select Case MsgBox(msg, vbQuestion Or vbYesNoCancel)
                    Case vbYes
                    Case vbCancel
                        cancel = (vbYes = MsgBox("Cancel message sending?", vbYesNo))
                        Set fld = Application.Session.GetDefaultFolder(olFolderSentMail)
                        Exit Function
                    Case vbNo
                        Set fld = Nothing
                        Set fld = Application.Session.PickFolder
                        If Not fld Is Nothing Then folderId = fld.EntryID
                End Select
            End If
        Else
            GoTo proc_err
        End If
    ElseIf unAttended Then
        Exit Function '---------------------------------------------------->>>>>>>>>>>
    ElseIf Not (suggestedFld Is Nothing) Then
        Set fld = suggestedFld
    Else 'pick manually...
        Set topicDic = Nothing
        Set fld = Application.Session.PickFolder
        If fld Is Nothing Then
            If (Not fromEvent) Or unAttended Then
                cancel = True
            Else
                cancel = (vbYes = MsgBox("Cancel message sending?", vbYesNo))
            End If
            If cancel Then Exit Function
            Set fld = Application.Session.GetDefaultFolder(olFolderSentMail)
        End If
    End If
    If Not fld Is Nothing Then
        If fromEvent Then
            mailIndexAppend mi.ConversationTopic, fld.EntryID, fld.folderPath
            Set mi.SaveSentMessageFolder = fld
            'DemoItemProperties mi
            'mi.PropertyAccessor
            'If mi.ItemProperties.item("Keywords") Is Nothing Then
            '    mi.ItemProperties.Add("Keywords", olKeywords).Value = "Classified"
            'Else
            '    mi.ItemProperties.item("keywords").Value = "Classified"
            'End If
            'Debug.Print
            
        Else
            mailIndexAppend mi.ConversationTopic, fld.EntryID, fld.folderPath
            If Not mi.UnRead And Not unAttended And Not (mi.parent Is fldInbox) Then
                ' prevent automatic archiving of unread mails in the default inbox folder
                Utilities.moveItem mi, fld, "ClassifyOneMail, Not fld Is Nothing"
            End If
        End If
        msg = ""
        If Not topicDic Is Nothing Then
            For i = 0 To topicDic.Count - 1
                If folderId <> topicDic.Keys(i) Then
                    If Not IsEmpty(topicDic.items(i)) Then
                        Set tEntry = topicDic.items(i)
                        msg = vbCrLf & tEntry.folderPath & vbTab & " " & tEntry.Count '    Application.Session.GetFolderFromID(topicDic.Keys(i)).folderPath & vbTab & " " & topicDic.Items(i)
                    End If
                End If
            Next i
        End If
        If Not unAttended Then
            If msg <> "" Then
                msg = "Some items matching this are stored elsewhere:" & msg & _
                    vbCrLf & "Do you want to move them also to " & folderName & "?"
                Select Case MsgBox(msg, vbQuestion Or vbYesNoCancel)
                    Case vbCancel
                        cancel = True
                        Exit Function
                    Case vbYes
                        For i = 0 To topicDic.Count - 1
                            trace.trace "ClassifyOneMail", Application.Session.GetFolderFromID(topicDic.Keys(i)).folderPath
                            If folderId <> topicDic.Keys(i) Then
                                For Each oMi In Application.Session.GetFolderFromID(topicDic.Keys(i)).items
                                    If TypeName(oMi) = "MailItem" Then
                                        Set foundMi = oMi
                                        Debug.Print foundMi.ConversationTopic
                                        If foundMi.ConversationTopic = mi.ConversationTopic Then
                                            Utilities.moveItem foundMi, fld, "ClassifyOneMail, items matching are stored elsewhere"
                                            reindexRequired = True
                                        End If
                                    End If
                                    DoEvents
                                Next oMi
                            End If ' <> folder
                        Next i 'folder
                                    
                End Select ' MsgBox
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
Private Function safeEntryId(ByVal fld As Folder) As String
    If Not (fld Is Nothing) Then safeEntryId = fld.EntryID
End Function
Public Sub ClassifyActiveMail()
    If TypeName(Application.ActiveWindow.CurrentItem) = "MailItem" Then
        ClassifyOneMail Application.ActiveWindow.CurrentItem, False
    End If
End Sub
Private Sub LocateActiveMail()
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

Private Sub ClassifySelectedMail()
Dim sel As Outlook.Selection
Dim fld As MAPIFolder
Dim origFld As MAPIFolder
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
    Dim msg As String, msglist As Collection
    Set msglist = New Collection
    For Each o In origFld.items
        If TypeName(o) = "MailItem" Then
            Set mi = o
            If convtopics.Exists(mi.ConversationTopic) Then
                msglist.add mi
                msg = msg & vbCrLf & mi.Subject & vbTab & " from " & mi.SenderName & " on " & mi.ReceivedTime
            End If
        End If
        DoEvents
    Next o
    If msg <> "" Then
        If vbYes = MsgBox(msg, vbYesNo, "Move also to " & fld.folderPath) Then
            For Each mi In msglist
                Utilities.moveItem mi, convtopics(mi.ConversationTopic), "ClassifySelectedMail, confirmed"
                DoEvents
            Next mi
        End If
    End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Sub CheckFolderMailsUnicity()
'Dim sel As Outlook.Selection
'Dim selectedObject As Object, exp As Explorer
'On Error GoTo proc_err
'GoTo proc
'proc_err:
'    MsgBox trace.trace("Error", Err.Number & " " & Err.Description & " in ClassifySelectedMail"), vbCritical
'    Exit Sub
'    Resume
'proc:
'    FolderUnicityLog.log.Text = ""
'
'    Set exp = Application.ActiveExplorer
'    For Each selectedObject In exp.Selection
'        If TypeName(selectedObject) = "MailItem" Then
'            Set FolderUnicityLog.fld = selectedObject.Parent
'            FolderUnicityLog.Show
'            Exit Sub
'        End If
'    Next selectedObject
'End Sub

Public Sub ArchiveAllAndCheckUnicity()
    ArchiveAndCheckUnicity GetArchivesFolder()
End Sub

Private Sub ArchiveAndCheckUnicity(root As MAPIFolder)
Dim fld As MAPIFolder
Dim o As Object
Dim items() As MailItem, i As Integer
ReDim items(100)
    
    For Each fld In root.folders
        ArchiveAndCheckUnicity fld
    Next fld
    Debug.Print root.folderPath
    For Each o In root.items
        If TypeName(o) = "MailItem" Then
            If i > UBound(items) Then ReDim Preserve items(UBound(items) + 100)
            Set items(i) = o
            i = i + 1
        End If
    Next o
    
    If i > 0 Then
        ReDim Preserve items(i - 1)
        CheckMailUnicityArray items
    End If
    Debug.Print root.folderPath & ": Check mail unicity done."

End Sub

Sub CheckMailUnicityArray(items() As MailItem)
Dim inbox As MAPIFolder:      Set inbox = GetInboxFolder()
Dim topics As Scripting.Dictionary: Set topics = New Scripting.Dictionary
Dim mi As MailItem, varItem As Variant
        
        For Each varItem In items
            Set mi = varItem
            If Not topics.Exists(mi.ConversationTopic) Then
                topics.add mi.ConversationTopic, 0
                CheckOneMailUnicity mi, False, fld:=inbox, verbose:=False
            End If
        Next varItem
End Sub


Public Sub CheckMailUnicity()
Dim sel As Outlook.Selection
Dim items() As MailItem, i As Integer
Dim item As MailItem, varItem As Variant
ReDim items(100)
Dim o As Object
    For Each o In Application.Explorers(1).Selection
        If TypeName(o) = "MailItem" Then
            If i > UBound(items) Then ReDim Preserve items(UBound(items) + 100)
            Set items(i) = o
            i = i + 1
        End If
    Next o
    If i > 0 Then
        ReDim Preserve items(i - 1)
        CheckMailUnicityArray items
    End If
    MsgBox "Check mail unicity done.", vbInformation
End Sub

Public Sub CheckOneMailUnicity(ByRef mi As MailItem, ByVal fromEvent As Boolean, Optional ByVal log As TextBox = Nothing, Optional ByVal fld As MAPIFolder = Nothing, Optional ByVal verbose As Boolean = True)
    Const timeout As Single = 1 / 24 / 60 / 12
    Dim startTime As Date
    Dim schs() As Outlook.Search
    Dim sch As Outlook.Search
    Dim rsts As Outlook.Results
    Dim searchFoldername As String
    Dim refMail As MailItem
    Dim foundObj As Object
    Dim foundMi As MailItem
    Dim refMailSign As MailSignature
    Dim foundMailSign As MailSignature
    Dim i As Integer, si As Integer
    Dim parentFlds() As MAPIFolder
    Dim parentFld As MAPIFolder
    Dim parentFldVar As Variant
    Dim ToBeDeleted As Collection
    Dim ToBeMoved As Collection
    Dim msg As String
    Dim doit As Boolean
    
    Dim strF As String
On Error GoTo proc_err
GoTo proc
proc_err:
    If verbose Then
        MsgBox trace.trace("Error", Err.Number & " " & Err.Description & " in CheckOneMailUnicity"), vbCritical
    Else
        trace.trace "Error", Err.Number & " " & Err.Description & " in CheckOneMailUnicity"
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
    End If
    
    Set ToBeDeleted = New Collection
    Set ToBeMoved = New Collection
    msg = ""
    Debug.Print mi.ConversationTopic
    For si = 0 To UBound(parentFlds)
        'If CheckUnicityMailSearchComp(si) Then
            'Set rsts = schs(si).Results
            For Each foundObj In parentFlds(si).items
                If TypeName(foundObj) = "MailItem" Then
                    Set foundMi = foundObj
                    If foundMi.EntryID = refMail.EntryID Then
                        ' ignore
                    ElseIf foundMi.parent Is Application.Session.GetDefaultFolder(olFolderSentMail) Then
                        ' ignore
                    ElseIf foundMi.ConversationTopic = mi.ConversationTopic Then                 ' now use the signature
                        Set foundMailSign = New MailSignature
                        foundMailSign.init foundMi
                        If refMailSign.contains(foundMailSign) Then
                            msg = msg & vbCrLf & refMailSign.ExtendedSubject & " contains " & foundMailSign.ExtendedSubject
                            ToBeDeleted.add foundMailSign
                        ElseIf refMailSign.isContainedIn(foundMailSign) Then
                            msg = msg & vbCrLf & refMailSign.ExtendedSubject & " is contained into " & foundMailSign.ExtendedSubject
                            ToBeDeleted.add refMailSign
                            Set refMail = foundMi
                            Set refMailSign = foundMailSign
                            If refMail.parent.folderPath <> mi.parent.folderPath Then
                                ToBeMoved.add refMailSign
                            End If
                        ElseIf foundMi.parent.folderPath <> mi.parent.folderPath Then
                            ToBeMoved.add foundMailSign
                        End If
                    End If
                    ' Debug.Print foundMi.ConversationTopic
                    DoEvents
                End If
            Next foundObj
        'End If
    Next si
    For Each refMailSign In ToBeDeleted
        msg = msg & vbCrLf & "==>Delete: " & refMailSign.ExtendedSubject
    Next refMailSign
    For Each refMailSign In ToBeMoved
        msg = msg & vbCrLf & "==>Move: " & refMailSign.ExtendedSubject & " from " & refMailSign.mail.parent.name & " to " & mi.parent.name
    Next refMailSign
    If ToBeDeleted.Count > 0 Or ToBeMoved.Count > 0 Then
        doit = False
        If log Is Nothing Then
        Else
            log.Text = log.Text & vbCrLf & msg
        End If
        If verbose Then
            doit = (vbYes = MsgBox(msg, vbYesNo Or vbQuestion, "Confirm following actions"))
        Else
            doit = True
        End If
        If doit Then
            For Each refMailSign In ToBeMoved
                Utilities.moveItem refMailSign.mail, mi.parent, "CheckOneMailUnicity"
            Next refMailSign
            For Each refMailSign In ToBeDeleted
                Utilities.moveItem refMailSign.mail, Nothing, "CheckOneMailUnicity"
            Next refMailSign
        End If
    End If
End Sub

Private Sub StartMakeMailIndex(ByVal loadfile As Boolean)
Dim filename As String, filebk As String, filepath As String
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
    MsgBox trace.trace("Error", Err.Number & " " & Err.Description & " in StartMakeMailIndex"), vbCritical
    Exit Sub
    Resume
proc:

On Error GoTo proc_err
    trace.trace "StartMakeMailIndex", "Start"
    ProgressBar.Caption = "Initial Indexing Status"
    ProgressBar.Show False
    
    ProgressBar.LabelPrefix = "Initializing..."
    ProgressBar.Value = 0
    DoEvents
    
    filepath = Environ("tmp") & "\OutlookIndex.txt"
    Set MailIndex = New Scripting.Dictionary
            
    If fs.FileExists(filepath) And loadfile Then
        Set ts = fs.OpenTextFile(filepath, ForReading)
        ProgressBar.LabelPrefix = "Loading..."
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
            ProgressBar.Setlabel topic
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
        ProgressBar.Hide
        If cnt < 10 Then 'not reasonably possible: there has been an issue with the file.
            doindex = True
        End If
    Else
        doindex = True
    End If
    If doindex Then
        On Error GoTo proc_err
        ProgressBar.MaxValue = CountMailFolders(GetArchivesFolder())
        ProgressBar.LabelPrefix = "Indexing..."
        MakeMailIndex GetArchivesFolder()
        ProgressBar.Hide
        mailIndexDump
        trace.trace "StartMakeMailIndex", "Application_Startup DONE"
    End If
    LastIndexDump = Now
End Sub
Public Sub ManualMakeMailIndex()
    StartMakeMailIndex False
End Sub

Private Function CountMailFolders(ByVal rootFl As MAPIFolder) As Long
Dim fl As MAPIFolder
    ProgressBar.Setlabel rootFl.folderPath
    If IndexFolder(rootFl.folderPath) Then
        CountMailFolders = 1 + rootFl.items.Count
    Else
        CountMailFolders = 1
    End If
    For Each fl In rootFl.folders
        CountMailFolders = CountMailFolders + CountMailFolders(fl)
        DoEvents
    Next fl
    
End Function

Private Sub MakeMailIndex(ByVal rootFl As MAPIFolder)
Dim obj As Object
Dim mi As MailItem
Dim topic As String
Dim id As String
Dim fl As MAPIFolder
Dim topicDic As Scripting.Dictionary
On Error GoTo proc_err
Dim activeCaption As String
GoTo proc
proc_err:
    MsgBox trace.trace("Error", Err.Number & " " & Err.Description & " in MakeMailIndex"), vbCritical
    Exit Sub
    Resume
proc:
    For Each fl In rootFl.folders
        MakeMailIndex fl
        DoEvents
    Next fl
    ProgressBar.Setlabel rootFl.folderPath
    If IndexFolder(rootFl.folderPath) Then
        For Each obj In rootFl.items
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
            ProgressBar.Value = ProgressBar.Value + 1
            DoEvents
        Next obj
    End If
    ProgressBar.Value = ProgressBar.Value + 1
End Sub

Sub mailIndexAppend(ByVal miTopic As String, ByVal flEntryId As String, ByVal folderPath As String)
Dim topicDic As Scripting.Dictionary
Dim tEntry As TopicEntry
'    trace.trace "mailIndexAppend", miTopic
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
Private Sub mailIndexDump()
Dim i As Long, j As Long
Dim filename As String, filebkname As String, filepath As String, filebkpath As String
Dim fileLine As String
Dim fs As New FileSystemObject
Dim ts As TextStream
Dim topicDic As Scripting.Dictionary
Dim flgReindex As Boolean
Dim tEntry As TopicEntry
GoTo proc
proc_err:
    If Err.Number = -2147221238 Then 'The operation cannot be performed because the object has been deleted.
        trace.trace "ERROR", Err.Number & " " & Err.Description & " in mailIndexDump"
        Resume Next
    End If
    If Err.Number = -2147220991 Then 'The messaging interface has returned an unknown error. If the problem persists, restart Outlook.
        flgReindex = True
        StartMakeMailIndex loadfile:=False
        Exit Sub
    End If
    MsgBox trace.trace("Error", Err.Number & " " & Err.Description & " in mailIndexDump"), vbCritical
    On Error Resume Next
    ts.Close
    Exit Sub
    Resume
proc:
    
    On Error GoTo proc_err:
    If MailIndex Is Nothing Then
        ' MsgBox "Mail index is empty: not saving", vbCritical
        Exit Sub
    End If
    ProgressBar.Caption = "Initial Indexing Status"
    ProgressBar.Show False
    ProgressBar.LabelPrefix = "Dumping..."
    ProgressBar.Value = 0
    
    filename = "OutlookIndex.txt"
    filepath = Environ("tmp") & "\" & filename
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
        Set topicDic = MailIndex.items(i)
        ProgressBar.Setlabel CStr(MailIndex.Keys(i))
        DoEvents
        For j = 0 To topicDic.Count - 1
            If Not IsEmpty(topicDic.items(j)) Then
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
    ProgressBar.Hide
End Sub
Private Function IndexFolder(ByVal folderPath As String) As Boolean
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

Private Function splitstr(ByVal str As String, delim As String _
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

Public Function parseXmlLine(ByVal line As String, ParamArray tags() As Variant) As Variant
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
Private Function MakeTopicEntry(cnt As Long, folderPath As String) As TopicEntry
    Set MakeTopicEntry = New TopicEntry
    MakeTopicEntry.Count = cnt
    MakeTopicEntry.folderPath = folderPath
End Function

'Private Sub JunkSelectedMail()
'Dim exp As Outlook.Explorer
'Dim o As Object, mi As MailItem, fwdMi As MailItem
'    Set exp = Application.ActiveWindow
'    For Each o In exp.Selection
'        If TypeName(o) = "MailItem" Then
'            Set mi = o
'            trace.trace "JunkSelectedMail, ", mi.SenderEmailAddress
'            MakeSpamMail mi
'        End If
'    Next o
'End Sub

'Private Sub JunkSelectedMail_oldversion()
'Dim sel As Outlook.Selection
'Dim exp As Outlook.Explorer
'Dim rul As Outlook.Rule
'Dim o As Object, mi As MailItem, fwdMi As MailItem
'Dim i As Integer
'Dim aList As Variant
'Dim flgAdd As Boolean
'Dim anyAdded As Boolean
'
'    trace.trace "JunkSelectedMail start"
'
'    If TypeName(Application.ActiveWindow) = "Explorer" Then
'
'        Set exp = Application.ActiveWindow
'        Set rul = GetInboxFolder().Store.GetRules.Item("extra spam")
'
'        For Each o In exp.Selection
'            If TypeName(o) = "MailItem" Then
'                Set mi = o
'                trace.trace "JunkSelectedMail, ", mi.SenderEmailAddress
'                flgAdd = True
'                aList = rul.Conditions.SenderAddress.Address
'                If IsArray(aList) Then
'                    For i = 0 To UBound(aList)
'                        If UCase(aList(i)) = UCase(mi.SenderEmailAddress) Then
'                            MsgBox "Address already in the list of the rule:" & vbCrLf _
'                                & mi.SenderEmailAddress & vbCrLf _
'                                & "Check if rule is enabled!", vbCritical
'                                flgAdd = False
'                            Exit For
'                        End If
'                    Next i
'                    If flgAdd Then
'                        ReDim Preserve aList(UBound(aList) + 1)
'                        aList(UBound(aList)) = mi.SenderEmailAddress
'                        anyAdded = True
'                    End If
'                Else
'                    aList = Array(rul.Conditions.SenderAddress.Address, mi.SenderEmailAddress)
'                End If
'                Set fwdMi = mi.Forward
'                trace.trace "JunkSelectedMail, fwd,", mi.SenderEmailAddress
'                DoEvents
'                fwdMi.Recipients.Add "Spam, Collection - Collection_Spam@cargill.com"
'                fwdMi.Send
'                trace.trace "JunkSelectedMail, sent,", mi.SenderEmailAddress
'                DoEvents
'                mi.Delete
'            End If
'        Next o
'        If anyAdded Then
'            rul.Conditions.SenderAddress.Address = aList
'            rul.Parent.save
'            trace.trace "JunkSelectedMail, rule saved"
'            'rul.Execute
'        End If
'    End If
'    trace.trace "JunkSelectedMail, done"
'End Sub
Private Sub exportRules()
Dim rul As Outlook.Rule
Dim ruleC As Outlook.RuleCondition
Dim ruleA As Outlook.RuleAction
Dim s As String
Dim fn As Integer
fn = FreeFile
Open "c:\temp\Outlook rules.txt" For Output As fn
    For Each rul In GetInboxFolder().Store.GetRules
        DispText fn, 0, ""
        DispText fn, 0, rul.name
        DispText fn, 0, "-------------------"
        DispText fn, 0, "Conditions:"
        For Each ruleC In rul.Conditions
            If ruleC.Enabled Then
                DispText fn, 1, TypeName(ruleC)
                Select Case TypeName(ruleC)
                    Case "AddressRuleCondition"
                        dispArray fn, 2, ruleC.Address
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
Sub DispText(ByVal fn As Integer, ByVal indent As Integer, ByVal t As String)
    Print #fn, t
    trace.trace String(indent, vbTab), t
End Sub
Sub dispArray(ByVal fn As Integer, ByVal indent As Integer, ByVal a As Variant)
Dim i As Integer
        For i = LBound(a) To UBound(a)
            DispText fn, indent, CStr(a(i))
        Next i
End Sub
Sub dispCollection(ByVal fn As Integer, ByVal indent As Integer, ByVal a As Variant)
Dim i As Integer
        For i = 1 To a.Count
            DispText fn, indent, CStr(a(i))
        Next i
End Sub


Private Sub ClassifyReadReceipts()
Dim mapi As NameSpace
Dim fldMail As MAPIFolder
Dim fldInbox As MAPIFolder
Dim fldReceipts As MAPIFolder
Dim fld As MAPIFolder
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
    MsgBox trace.trace("Error", Err.Number & " " & Err.Description & " in ClassifyReadReceipts"), vbCritical
    Exit Sub
    Resume
    Exit Sub
proc:

    Set mapi = Application.GetNamespace("Mapi")
    For Each fld In mapi.folders
        If LCase(fld.name) Like "philippe.raemy@*" _
            And Not LCase(fld.name) Like "*shared*" _
            And Not LCase(fld.name) Like "*releasemanagement*" Then
            Set fldMail = fld
            Exit For
        End If
    Next fld
    If fldMail Is Nothing Then
        MsgBox "Could not find mailbox folder.", vbCritical
    End If
    Set fldInbox = GetInboxFolder()
    Set fldReceipts = fldMail.folders("_Receipts")
    
    For i = fldInbox.items.Count To 1 Step -1
        Set obj = fldInbox.items(i)
        Select Case TypeName(obj)
            Case "MailItem"
'                Set mi = obj
'                'trace.trace mi.SenderName, mi.subject
'                For Each word In Split(mi.subject)
'                    If dic.Exists(word) Then
'                        dic(word) = dic(word) + 1
'                    Else
'                        dic(word) = 1
'                    End If
'                Next word
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
    '    trace.trace "ClassifyReadReceipts", dic.Key(i), dic.Item(i)
    'Next i
End Sub

Sub ClassifyFromSearchFoldersProc(ByVal destFld As MAPIFolder)
        Dim colStores As Outlook.Stores
        Dim oStore As Outlook.Store
        Dim oSearchFolders As Outlook.folders
        Dim oFolder As Outlook.Folder
        Dim DestSubFld As Outlook.Folder
        Dim DestSubFldName As String
        Dim itm As Outlook.MailItem
        Dim i As Integer ' , j As Integer
'        Dim folderNames As Variant
        
        Dim SearchFolderNameRoots As Variant
        Dim NameRoot As Variant
        Dim context As String
        Dim fDebug As Boolean

On Error GoTo proc_err
GoTo proc
proc_err:
        If Err.Number = 440 And context = "looping messages" Then
                Debug.Print Err.Description
                ' index out of bound
                Err.Clear
                Resume Next
        End If

        MsgBox trace.trace("Error", Err.Number & " " & Err.Description & " in ClassifyFromSearchFoldersProc"), vbCritical
        Exit Sub
        Resume
        Exit Sub
proc:
    
    SearchFolderNameRoots = Array("LongTerm*", "ShortTerm*")
    ProgressBar.init "customs", "Store searches"
    ProgressBar.SetCaption "Classify From Search Folders"
    
    ClassifyReadReceipts
    
    Set colStores = Application.Session.Stores
    For Each oStore In colStores
        If oStore.ExchangeStoreType = olPrimaryExchangeMailbox Then
            Set oSearchFolders = oStore.GetSearchFolders
            For Each oFolder In oSearchFolders
                trace.trace "ClassifyFromSearchFoldersProc", oFolder.folderPath
                If fDebug Then Debug.Print oFolder.folderPath,
                If Left(oFolder.folderPath, 2) = "\\" _
                And Not oFolder.name Like "MS-OLK*" Then
                    For Each NameRoot In SearchFolderNameRoots
                        If oFolder.name Like NameRoot Then
                            If fDebug Then Debug.Print "match " & NameRoot
                            ProgressBar.SetCaption oFolder.name
                            DestSubFldName = oFolder.name
                            If InStr(1, DestSubFldName, "{", vbBinaryCompare) > 0 Then
                                DestSubFldName = Mid(DestSubFldName, 1, InStr(1, DestSubFldName, "{", vbBinaryCompare) - 1)
                            End If
                            Set DestSubFld = EnsureFolderExists(destFld, DestSubFldName)
                            ' this complex loop approach because of multiple threading issues.
                            ' just doing "While oFolder.Items.Count >= 1" works but is much slower than the for, hence a combination fo the two
                            While oFolder.items.Count >= 1
                                For i = oFolder.items.Count To 1 Step -1
                                        If i <= oFolder.items.Count Then
                                                context = "looping messages"
                                                Utilities.moveItem oFolder.items(i), DestSubFld, "ClassifyFromSearchFoldersProc"
                                                context = ""
                                                DoEvents
                                        End If
                                Next i
                            Wend
                        End If
                    Next NameRoot
                End If
                If fDebug Then Debug.Print
            Next
            Exit For ' once the main store was found there is not need to further search
        End If
    Next
    ProgressBar.Hide
End Sub
Public Function EnsureFolderExists(ByVal parentFolder As Outlook.Folder, ByVal folderPath As String) As Outlook.Folder
Dim a As Variant
Dim i As Integer
Dim parentPath As String

Set EnsureFolderExists = parentFolder

a = Split(folderPath, "\")
For i = 0 To UBound(a)
    If a(i) <> "" Then
        On Error Resume Next
        Set EnsureFolderExists = EnsureFolderExists.folders(a(i))
        If Err.Number <> 0 Then
            On Error GoTo 0
            Set EnsureFolderExists = EnsureFolderExists.folders.add(a(i))
        End If
    End If
Next i
End Function
Public Sub HistorizeMails()
        HistorizeMailsProc "all"
End Sub
Public Sub ClassifyFromSearchFolders()
        HistorizeMailsProc "fromsearch"
End Sub

Public Sub HistorizeMailsProc(ByVal action As String)

'inbox older than 1 mth
'inbox subfolders older than 8 months
'Various older than 15 days
'ShortTerm older than 15 days, then deleted
'Some obsolete folders right now.


Dim fldInbox As MAPIFolder
Dim fldPers As MAPIFolder
Dim fldPersRoot As MAPIFolder
Dim mapi As NameSpace
Dim obj As Object
Dim InboxRetentionDate As Date
Dim CalendarRetentionDate As Date
Dim FoldersRetentionDate As Date
Dim ShortTermRetentionDate As Date
Dim ArchivesRetentionDate As Date
Dim ObsoleteFoldersRetentionDate As Date
Dim ObsoleteFolders As Variant
Dim subIbFld As MAPIFolder
Dim subPerFld As MAPIFolder
Dim fld As MAPIFolder
Dim i As Integer
Dim msg As String

On Error GoTo proc_err
GoTo proc
proc_err:
    MsgBox trace.trace(Err.Number & " " & Err.Description & " in HistorizeMailsProc"), vbCritical
    Exit Sub
    Resume
    Exit Sub
proc:

    Dim bs As New ProgressBar
    ProgressBar.init "customs", "Historize"

    trace.trace "HistorizeMailsProc(""" & action & """)"
    InboxRetentionDate = DateAdd("d", -15, Now())
    FoldersRetentionDate = DateAdd("m", 0, Now())
    ShortTermRetentionDate = DateAdd("d", -15, Now())
    ArchivesRetentionDate = DateAdd("d", -180, Now())
    CalendarRetentionDate = DateAdd("d", -90, Now())

    ObsoleteFoldersRetentionDate = Now()
    ObsoleteFolders = Array() '"Audit")
    
    
    Set mapi = Application.GetNamespace("Mapi")
    Set fldPersRoot = GetArchivesFolder() ' FindFolder(mapi, "Archives")
    Set fldPers = fldPersRoot ' mapi.folders("archives")
    Set fldInbox = GetInboxFolder()
    
    Select Case action
        Case "all", "fromsearch"
            ProgressBar.SetCaption "Classify From Search Folders"
            ClassifyFromSearchFoldersProc fldPers ' GetInboxFolder().Parent
            For Each fld In fldPers.folders
                If LCase(fld.name) Like "shortterm*" Then
                    RecurseHistorizeMails fld, Nothing, GetRetentionDateFromFolderName(fld.name, ShortTermRetentionDate)
                End If
            Next fld
            RecurseHistorizeMails Application.Session.GetDefaultFolder(olFolderDeletedItems), Nothing, ShortTermRetentionDate _
                , deleteEmptyFolders:=True
    End Select
    Select Case action
        Case "all", "historize"
        ProgressBar.SetCaption "Make mail index"
        StartMakeMailIndex loadfile:=False
    
        ProgressBar.SetCaption "Classify mails in Inbox"
        For i = fldInbox.items.Count To 1 Step -1
            Set obj = fldInbox.items(i)
            If TypeName(obj) = "MailItem" Then
                Set fld = ClassifyOneMail(obj, False, unAttended:=True)
                'If fld Is Nothing Then HandleIncomingMails obj
            End If
            DoEvents
        Next i
        For Each subIbFld In fldInbox.folders
            ProgressBar.SetCaption "Historize mails in " & subIbFld.name
            Set subPerFld = EnsureFolderExists(fldPers, subIbFld.name)
            Select Case LCase(subIbFld.name)
                Case "shortterm"
                    RecurseHistorizeMails subIbFld, Nothing, ShortTermRetentionDate
                Case "archivesretentiondate"
                    RecurseHistorizeMails subIbFld, subPerFld, InboxRetentionDate
                Case Else
                    RecurseHistorizeMails subIbFld, subPerFld, FoldersRetentionDate
            End Select
        Next subIbFld
        RecurseHistorizeMails GetInboxFolder().parent.folders("_Receipts"), Nothing, ShortTermRetentionDate _
            , deleteEmptyFolders:=True
        'RecurseHistorizeMails Application.Session.GetDefaultFolder(olFolderCalendar), EnsureFolderExists(fldPersRoot, "Calendar"), CalendarRetentionDate _
        '    , deleteEmptyFolders:=True
        RecurseHistorizeMails GetArchivesFolder(), mapi.folders("Personal Zarchives"), ArchivesRetentionDate _
            , deleteEmptyFolders:=True, specialFolders:=ObsoleteFolders, specialFoldersRetentionDate:=ObsoleteFoldersRetentionDate
        ' RecurseHistorizeMails fldInbox.folders("Deleted Items"), Nothing, ShortTermRetentionDate _
            , deleteEmptyFolders:=True
        RecurseHistorizeMails Application.Session.GetDefaultFolder(olFolderDeletedItems), Nothing, ShortTermRetentionDate _
            , deleteEmptyFolders:=True
        RecurseHistorizeMails GetArchivesFolder().folders("Deleted Items"), Nothing, Now _
            , deleteEmptyFolders:=True
        ProgressBar.SetCaption "Make mail index"
        StartMakeMailIndex loadfile:=False
    End Select
    trace.trace "HistorizeMailsProc(""" & action & """): done"
    
End Sub
Public Function GetRetentionDateFromFolderName(ByVal fldName As String, ByVal defltDate As String) As Date
    Dim NumDaysCandidate As String
    GetRetentionDateFromFolderName = defltDate
    On Error Resume Next
    NumDaysCandidate = Mid(fldName, InStr(fldName, "[") + 1, InStr(fldName, "]") - InStr(fldName, "[") - 1)
    If Not IsNumeric(NumDaysCandidate) Then Exit Function
    GetRetentionDateFromFolderName = DateAdd("d", -CInt(NumDaysCandidate), Now())
End Function
Private Function FindFolder(ByVal root As Variant, ByVal folderName As String) As Outlook.Folder
On Error GoTo err_proc
GoTo proc:
err_proc:
    If TypeName(root) = "NameSpace" Then
        Debug.Print "Error searching for folder '" & folderName & "' in  namespace '" & root & "':" & Err.Description
    Else
        Debug.Print "Error searching for folder '" & folderName & "' in '" & root.name & "':" & Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume
proc:

    Dim fld As Outlook.Folder
    For Each fld In root.folders
        If UCase(fld.name) = UCase(folderName) Then
            Set FindFolder = fld
            Exit Function
        End If
    Next fld
    For Each fld In root.folders
        If Not fld.folderPath Like "*All Public Folders*" _
        And Not fld.folderPath Like "*SharePoint*" _
        Then
            Set FindFolder = FindFolder(fld, folderName)
            If Not FindFolder Is Nothing Then Exit Function
        End If
    Next fld
End Function

Private Sub RecurseHistorizeMails( _
    ByVal fldInbox As MAPIFolder, _
    ByVal fldPers As MAPIFolder, _
    ByVal retentionDate As Date, _
    Optional ByVal deleteEmptyFolders As Boolean, _
    Optional ByVal specialFolders As Variant, _
    Optional ByVal specialFoldersRetentionDate As Date _
)
Dim subIbFld As MAPIFolder
Dim subPerFld As MAPIFolder
Dim moveToFld As Boolean
Dim i As Integer, j As Integer
Dim recursed As Boolean
Dim obj As Object
Dim msg As String

On Error GoTo proc_err
GoTo proc
proc_err:
    MsgBox trace.trace("Error", Err.Number & " " & Err.Description & " in HistorizeMailsProc"), vbCritical
    Exit Sub
    Resume
    Exit Sub
proc:

    msg = "RecurseHistorizeMails(" & fldInbox.folderPath
    If fldPers Is Nothing Then
        msg = msg & ", Delete"
    Else
        msg = msg & "," & fldPers.folderPath
    End If
    msg = msg & "," & retentionDate & ")"
    ProgressBar.SetCaption msg
    trace.trace msg
    Debug.Assert fldInbox.name <> "Commodities Daily"
    
    For i = fldInbox.items.Count To 1 Step -1
        Set obj = fldInbox.items(i)
        moveToFld = False
        Select Case TypeName(obj)
            Case "MailItem", "MeetingItem", "PostItem", "SharingItem"
                If obj.SentOn < retentionDate Or obj.SentOn > DateAdd("y", 2000, Now) Then
                    moveToFld = True
                ElseIf obj.SentOn > Now Then
                    trace.trace "RecurseHistorizeMails", obj.SentOn & "," & obj.Subject & "<--" & fldInbox.folderPath
                    'Debug.Assert False
                End If
            Case "AppointmentItem"
                If Not obj.IsRecurring And (obj.Start < retentionDate Or obj.Start > DateAdd("y", 2000, Now)) And obj.Attachments.Count > 0 Then
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
                    RecurseHistorizeMails subIbFld, subPerFld, specialFoldersRetentionDate, deleteEmptyFolders
                    recursed = True
                End If
            Next j
        End If
        If Not recursed Then
            RecurseHistorizeMails subIbFld, subPerFld, retentionDate, deleteEmptyFolders _
            , specialFolders, specialFoldersRetentionDate
        End If
        DoEvents
    Next i
    If deleteEmptyFolders And fldInbox.folders.Count = 0 And fldInbox.items.Count = 0 And Not fldInbox.name = "Deleted Items" Then
        trace.trace "RecurseHistorizeMails", "Deleting folder «" & fldInbox.folderPath & "»"
        fldInbox.delete
    End If
    trace.trace "RecurseHistorizeMails", "(" & fldInbox.folderPath & "): <=== done"
End Sub
Sub RemoveAllItemsAndFoldersInFolder(ByVal fld As Outlook.Folder)
        Dim oFolders As Outlook.folders
        Dim oItems As Outlook.items
        Dim i As Long
        'Obtain a reference to deleted items folder
        Set oItems = fld.items
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
        Dim oStore As Outlook.Store
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




