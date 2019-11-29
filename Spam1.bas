Attribute VB_Name = "Spam1"
Dim spammers As Scripting.Dictionary
Const JunkFolderRootName = "Inbox\shortTerm\Junk E-mail{%1}"

Private Sub initSpammersList()
Dim fs As FileSystemObject
Dim ts As TextStream
On Error GoTo proc_err
GoTo proc
proc_err:
  MsgBox Err.Number & " " & Err.Description & " in initSpammersList", vbCritical
  Exit Sub
  Resume
proc:

  If spammers Is Nothing Then
    Set spammers = New Scripting.Dictionary
    spammers.CompareMode = TextCompare
    Set fs = New FileSystemObject
    Set ts = fs.OpenTextFile(Environ("USERPROFILE") & "\Local Settings\Application Data\Microsoft\Outlook\spammers.txt", ForReading)
    While Not ts.AtEndOfStream
      spammers(ts.ReadLine) = Empty
    Wend
    ts.Close
    Set ts = Nothing
    Set fs = Nothing
  End If
End Sub
Private Sub SaveSpammersList()
Dim fs As FileSystemObject
Dim ts As TextStream
Dim k As Variant
  If spammers Is Nothing Then Exit Sub
  Set fs = New FileSystemObject
  Set ts = fs.OpenTextFile(Environ("USERPROFILE") & "\Local Settings\Application Data\Microsoft\Outlook\spammers.txt", ForWriting, Create:=True)
  ts.WriteLine "/* file created on " & Format(Now, "yyyy-mm-dd hh:mm:ss") & " */"
  For Each k In spammers.Keys
    ts.WriteLine k
  Next k
  ts.WriteLine "/* end of file */"
  ts.Close
  Set ts = Nothing
  Set fs = Nothing
End Sub

Public Sub MakeSpamMail(item As MailItem)
On Error GoTo proc_err
GoTo proc
proc_err:
  MsgBox Err.Number & " " & Err.Description & " in MakeSpamMail", vbCritical
  Exit Sub
  Resume
proc:
AddSearchFolder item.SenderEmailAddress, item.SenderEmailAddress
Dim AlreadyFound As Boolean
  initSpammersList
  For i = 0 To spammers.Count - 1
    If InStr(1, item.SenderEmailAddress, spammers.Keys(i), vbTextCompare) > 0 Then
      AlreadyFound = True
      Exit For
    End If
  Next i
  If Not AlreadyFound Then
    spammers.add item.SenderEmailAddress, Empty
    SaveSpammersList
  End If
  HandleIncomingMails item
End Sub

Public Sub HandleIncomingMails(item As MailItem)
Dim obj As Object
Dim rObj As ReportItem
Dim mObj As MailItem
Dim i As Integer
Dim fld As Outlook.Folder
Dim fs As New FileSystemObject
Dim ts As TextStream
Dim Msg As String

On Error GoTo proc_err
GoTo proc
proc_err:
  If Err.Number = -2147352567 Then
    Resume Next
  Else
    MsgBox Err.Number & " " & Err.Description & " in HandleIncomingMails", vbCritical
    Exit Sub
  End If
  Resume
proc:


  trace.trace "From «" & item.SenderEmailAddress & "»: " & item.subject
  initSpammersList
  For i = 0 To spammers.Count - 1
    If InStr(1, item.SenderEmailAddress, spammers.Keys(i), vbTextCompare) > 0 Then
      If InStr(1, item.SenderEmailAddress, "cargill.com", vbTextCompare) > 0 Then
        Set fld = Application.Session.GetDefaultFolder(olFolderInbox).folders("Various")
      Else
        Set fld = Application.Session.GetDefaultFolder(olFolderJunk)
      End If
      item.Move fld
      Msg = Format(Now, "YYYY-MM-DD hh:mm:ss") & " Move «" & item.subject & "»" & _
        " from «" & item.parent.folderPath & "»" & _
        " to " & fld.folderPath
      Set ts = fs.OpenTextFile(Environ("USERPROFILE") & "\Local Settings\Application Data\Microsoft\Outlook\spammers.log", ForAppending, True)
      ts.WriteLine Msg
      ts.Close
      trace.trace Msg
      Exit For
    End If
  Next i
End Sub
Public Function AddSearchFolder(mailAddress As String, Optional searchName As String) As Boolean
Dim oStore As store
Dim primaryStore As store
Dim SearchFld As Folder
Dim scope As String
Dim searchresult As Search
On Error GoTo proc_err
GoTo proc
proc_err:
    MsgBox Err.Number & " " & Err.Description & " in AddSearchFolder", vbCritical
    Exit Function
    Resume
proc:
    
    If searchName = "" Then searchName = mailAddress

    For Each oStore In Application.Session.Stores
    
        If oStore.ExchangeStoreType = olPrimaryExchangeMailbox Then
            Set primaryStore = oStore
            Set oSearchFolders = oStore.GetSearchFolders
            For Each SearchFld In oSearchFolders
                If SearchFld.name Like JunkFolderName & "*" Then
                    AddSearchFolder = False
                    Exit Function
                End If
            Next
        End If
    Next
    'If arrived there, we've not found the search folder: create on the main store
    scope = "'" & Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).folderPath & "'"
    Dim filter As String: filter = "(""urn:schemas:httpmail:from"" ci_phrasematch '" & searchName & "')" _
      & " OR (""urn:schemas:httpmail:to"" ci_phrasematch '" & searchName & "')"
    Set searchresult = Application.AdvancedSearch(scope, filter, False)
    searchresult.save searchName
    AddSearchFolder = True
  
End Function
Sub DisplayAvailableScopes()

    'Declare a variable that references a
    'SearchScope object.
    Dim ss As SearchScope
    Dim sss As SearchScopes

        'Loop through the SearchScopes collection.
        For Each ss In sss
            Select Case ss.Type
                Case msoSearchInMyComputer
                    MsgBox "My Computer is an available search scope."
                Case msoSearchInMyNetworkPlaces
                    MsgBox "My Network Places is an available search scope."
                Case msoSearchInOutlook
                    MsgBox "Outlook is an available search scope."
                Case msoSearchInCustom
                    MsgBox "A custom search scope is available."
                Case Else
                    MsgBox "Can't determine search scope."
            End Select
        Next ss

End Sub
Sub initJunkSearchFolders()
AddSearchFolder "%qoqa.ch%", "ShortTerm\Qoqa"
AddSearchFolder "adobe.com"
AddSearchFolder "airdefense.net"
AddSearchFolder "altigenweb-mail.info"
AddSearchFolder "altigenwebmail.info"
AddSearchFolder "angel.com"
AddSearchFolder "angel.com"
AddSearchFolder "announcements.informatica-news.com"
AddSearchFolder "ArchitectureSummit.net"
AddSearchFolder "asaaaa.com"
AddSearchFolder "ashley.taylor@shunra.com"
AddSearchFolder "castsoftware.com"
AddSearchFolder "cavisualdesign-mail.info"
AddSearchFolder "ccpguides-mails.info"
AddSearchFolder "centrifugesystems.com"
AddSearchFolder "communicatevisually.com"
AddSearchFolder "communicatevisually.com"
AddSearchFolder "connect.vmware.com"
AddSearchFolder "creditcardprocessguides.info"
AddSearchFolder "cybercartes-mail.com"
AddSearchFolder "db.nl00.net"
AddSearchFolder "defensepactom.com"
AddSearchFolder "dkpromo-mail.info"
AddSearchFolder "docucrunch.com"
AddSearchFolder "DocuCrunch.com"
AddSearchFolder "eiqnetworks.com"
AddSearchFolder "elastra.com"
AddSearchFolder "en25.com"
AddSearchFolder "FinanceTechNews.com"
AddSearchFolder "FinanceTechNews.com"
AddSearchFolder "FinanceTechNews.com"
AddSearchFolder "focus-erpmail.info"
AddSearchFolder "focuscrmmail.info"
AddSearchFolder "focusvoipguides.info"
AddSearchFolder "hardwarecity-mail.info"
AddSearchFolder "i-speak-mail.info"
AddSearchFolder "info.newscale.com"
AddSearchFolder "infosys.com"
AddSearchFolder "interwoven.com"
AddSearchFolder "jgs-dom-notification.com"
AddSearchFolder "mail.communications.sun.com"
AddSearchFolder "mail.vresp.com"
AddSearchFolder "mail.vresp.com"
AddSearchFolder "messagelabs.com"
AddSearchFolder "mindtree.com"
AddSearchFolder "mindtree.com"
AddSearchFolder "morecrm-mails.info"
AddSearchFolder "netapp.com"
AddSearchFolder "nonewsletter.resaplus.ch"
AddSearchFolder "nosonicwall.com"
AddSearchFolder "noverizonwireless.com"
AddSearchFolder "offers.ztfsg.com"
AddSearchFolder "omniture.com"
AddSearchFolder "omniture.com"
AddSearchFolder "omniture.com"
AddSearchFolder "onhold-companymail.info"
AddSearchFolder "onholdco-mail.info"
AddSearchFolder "optier.marketbright.com"
AddSearchFolder "osibusinessmail.info"
AddSearchFolder "owireless-mails.info"
AddSearchFolder "pbp-executivereports.net"
AddSearchFolder "pbpmedia.com"
AddSearchFolder "pbpmedia.com"
AddSearchFolder "pbtechnologytraining.com"
AddSearchFolder "pbtechnologytraining.com"
AddSearchFolder "pdb33.info"
AddSearchFolder "polaris.co.in"
AddSearchFolder "polaris.com"
AddSearchFolder "progressivebusinesstechnologytraining.com"
AddSearchFolder "rapidresponsemarketinginc.com"
AddSearchFolder "reply.informatica-news.com"
AddSearchFolder "reply.mb00.net"
AddSearchFolder "sgi.com"
AddSearchFolder "shunra.com"
AddSearchFolder "smartdraw.com"
AddSearchFolder "smartdrawcommunity.com"
AddSearchFolder "smartdrawcommunity.com"
AddSearchFolder "smartdrawinfo.com"
AddSearchFolder "smartdrawinfo.com"
AddSearchFolder "spl03.net"
AddSearchFolder "ssimpson@layer7tech.com"
AddSearchFolder "systemsinmotion.com"
AddSearchFolder "targetedconferences.com"
AddSearchFolder "targetedconferences.com"
AddSearchFolder "tp.omnichannel.net"
AddSearchFolder "trendmicro.rsys1.com"
AddSearchFolder "trythenewsilktest@microfocus.com"
AddSearchFolder "verizonwireless.com"
AddSearchFolder "vietnamam.com"
AddSearchFolder "vinmails.info"
AddSearchFolder "voipguidemail.info"
End Sub
