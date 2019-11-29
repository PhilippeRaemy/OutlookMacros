Attribute VB_Name = "test"
Sub EnumerateFoldersInStores()
    Dim colStores As Outlook.Stores
    Dim oStore As Outlook.store
    Dim oRoot As Outlook.Folder
    Dim colRule As Outlook.Rules
    Dim oRule As Outlook.Rule
    Dim oCondition As Outlook.RuleCondition
    
    On Error Resume Next
    Set colStores = Application.GetNamespace("MAPI").Stores
    For Each oStore In colStores
      Set oRoot = oStore.GetRootFolder
      Set colRule = oStore.GetRules()
      Debug.Print oRoot.name
      For Each oRule In oStore.GetRules()
        Debug.Print , oRule.name
        For Each oCondition In oRule.Conditions
          If oCondition.Enabled Then
            Debug.Print , , TypeName(oCondition)
          End If
        Next
      Next oRule
    Next
End Sub

Sub enumerateButtons()
Dim cmdb As CommandBar
Dim cmdbb As CommandBarButton
Dim obj As Object
For Each cmdb In Application.ActiveExplorer.CommandBars
  Debug.Print cmdb.name
  For Each obj In cmdb.Controls
    On Error Resume Next
    Set cmdbb = obj
    If Err.Number = 0 Then
        Debug.Print cmdb.name & "." & cmdbb.caption
        If cmdbb.caption = "Store searches" Then
        'customs.Historize
          Err.Clear
          On Error GoTo 0
          cmdbb.Execute
          
        End If
    Else
      Debug.Print cmdb.name & ".[" & TypeName(obj) & "]"
    End If
    Err.Clear
    On Error GoTo 0
  Next obj
  
Next cmdb
End Sub

Sub DemoPropertyAccessorGetProperty()
    Dim PropName, Header As String
    Dim oMail As Object
    Dim oPA As Outlook.PropertyAccessor
    'Get first item in the inbox
    Set oMail = _
        Application.Session.GetDefaultFolder(olFolderInbox).Items(1)
    'PR_TRANSPORT_MESSAGE_HEADERS
    PropName = "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/keywords" '"http://schemas.microsoft.com/mapi/proptag/0x007D001E"
    'Obtain an instance of PropertyAccessor class
    Set oPA = oMail.PropertyAccessor
    'Call GetProperty
    Header = oPA.GetProperty(PropName)
    Debug.Print (Header)
End Sub
Sub DemoItemProperties(mi As MailItem)
    Dim PropName, Header As String
    Dim i As Integer
    'Get first item in the inbox
    'Set mi = Application.Session.GetDefaultFolder(olFolderInbox).Items(1)
    'PR_TRANSPORT_MESSAGE_HEADERS
    Debug.Print TypeName(mi.ItemProperties)
    For i = 0 To mi.ItemProperties.Count - 1
        Debug.Print i; "]", mi.ItemProperties(i).name, mi.ItemProperties(i).Type
    Next i
End Sub

Sub DemoPropertyAccessorSetProperties()
    Dim PropNames(), myValues() As Variant
    Dim arrErrors As Variant
    Dim prop1, prop2, prop3, prop4 As String
    Dim i As Integer
    Dim oMail As Outlook.MailItem
    Dim oPA As Outlook.PropertyAccessor
    'Get first item in the inbox
    Set oMail = _
        Application.Session.GetDefaultFolder(olFolderInbox).Items(3)
    'Names for properties using the MAPI string namespace
    prop1 = "http://schemas.microsoft.com/mapi/string/" & _
        "{FFF40745-D92F-4C11-9E14-92701F001EB3}/mylongprop"
    prop2 = "http://schemas.microsoft.com/mapi/string/" & _
        "{FFF40745-D92F-4C11-9E14-92701F001EB3}/mystringprop"
    prop3 = "http://schemas.microsoft.com/mapi/string/" & _
        "{FFF40745-D92F-4C11-9E14-92701F001EB3}/mydateprop"
    prop4 = "http://schemas.microsoft.com/mapi/string/" & _
        "{FFF40745-D92F-4C11-9E14-92701F001EB3}/myboolprop"
    PropNames = Array(prop1, prop2, prop3, prop4)
    myValues = Array(1020, "111-222-Kudo", Now(), False)
    'Set values with SetProperties call
    'If the properties do not exist, then SetProperties
    'adds the properties to the object when saved.
    'The type of the property is the type of the element
    'passed in myValues array.
    Set oPA = oMail.PropertyAccessor
    arrErrors = oPA.SetProperties(PropNames, myValues)
    If Not (IsEmpty(arrErrors)) Then
        'Examine the arrErrors array to determine if any
        'elements contain errors
        For i = LBound(arrErrors) To UBound(arrErrors)
            'Examine the type of the element
            If IsError(arrErrors(i)) Then
                Debug.Print (CVErr(arrErrors(i)))
            End If
        Next
    End If
    'Save the item
    oMail.save
End Sub

Sub DemoPropertyAccessorGetProperties()
    Dim PropNames(), myValues() As Variant
    Dim arrErrors As Variant
    Dim prop1, prop2, prop3, prop4 As String
    Dim i As Integer
    Dim oMail As Outlook.MailItem
    Dim oPA As Outlook.PropertyAccessor
    'Get first item in the inbox
    Set oMail = _
        Application.Session.GetDefaultFolder(olFolderInbox).Items(3)
    'Names for properties using the MAPI string namespace
    prop1 = "http://schemas.microsoft.com/mapi/string/" & _
        "{FFF40745-D92F-4C11-9E14-92701F001EB3}/mylongprop"
    prop2 = "http://schemas.microsoft.com/mapi/string/" & _
        "{FFF40745-D92F-4C11-9E14-92701F001EB3}/mystringprop"
    prop3 = "http://schemas.microsoft.com/mapi/string/" & _
        "{FFF40745-D92F-4C11-9E14-92701F001EB3}/mydateprop"
    prop4 = "http://schemas.microsoft.com/mapi/string/" & _
        "{FFF40745-D92F-4C11-9E14-92701F001EB3}/myboolprop"
    PropNames = Array(prop1, prop2, prop3, prop4)
    myValues = Array(1020, "111-222-Kudo", Now(), False)
    'Set values with SetProperties call
    'If the properties do not exist, then SetProperties
    'adds the properties to the object when saved.
    'The type of the property is the type of the element
    'passed in myValues array.
    Set oPA = oMail.PropertyAccessor
    arrErrors = oPA.FetProperties(PropNames)
    If Not (IsEmpty(arrErrors)) Then
        'Examine the arrErrors array to determine if any
        'elements contain errors
        For i = LBound(arrErrors) To UBound(arrErrors)
            'Examine the type of the element
            If IsError(arrErrors(i)) Then
                Debug.Print (CVErr(arrErrors(i)))
            End If
        Next
    End If
    'Save the item
    'oMail.save
End Sub
Sub testdic()
    Dim a As Variant
    a = Array("z", "a", "d")
    BubbleSort a
    Dim k As Variant
    For Each k In a
        Debug.Print (k)
    Next k
    Dim folders() As String
    Debug.Print LBound(folders)
End Sub

