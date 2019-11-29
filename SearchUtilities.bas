Attribute VB_Name = "SearchUtilities"
Option Explicit

Sub SearchFolderForSender()

On Error GoTo Err_SearchFolderForSender

Dim strFrom As String
Dim strTo As String

' get the name & email address from a selected message
Dim oMail As Outlook.MailItem
Set oMail = ActiveExplorer.Selection.item(1)

strFrom = oMail.SenderEmailAddress
strTo = oMail.SenderName

If strFrom = "" Then Exit Sub

Dim strDASLFilter As String

' From & To fields
Const From1 As String = "http://schemas.microsoft.com/mapi/proptag/0x0065001f"
Const From2 As String = "http://schemas.microsoft.com/mapi/proptag/0x0042001f"
Const To1 As String = "http://schemas.microsoft.com/mapi/proptag/0x0e04001f"
Const To2 As String = "http://schemas.microsoft.com/mapi/proptag/0x0e03001f"

strDASLFilter = "((""" & From1 & """ CI_STARTSWITH '" & strFrom & "' OR """ & From2 & """ CI_STARTSWITH '" & strFrom & "')" & _
" OR (""" & To1 & """ CI_STARTSWITH '" & strFrom & "' OR """ & To2 & """ CI_STARTSWITH '" & strFrom & "' OR """ & To1 & """ CI_STARTSWITH '" & strTo & "' OR """ & To2 & """ CI_STARTSWITH '" & strTo & "' ))"

Debug.Print strDASLFilter

Dim strScope As String
strScope = "'Inbox', 'Sent Items'"
    
Dim objSearch As Search
Set objSearch = Application.AdvancedSearch(scope:=strScope, filter:=strDASLFilter, SearchSubFolders:=True, Tag:="SearchFolder")

'Save the search results to a searchfolder
objSearch.save (strTo)

Set objSearch = Nothing

Exit Sub

Err_SearchFolderForSender:
MsgBox "Error # " & Err & " : " & Error(Err)

End Sub

