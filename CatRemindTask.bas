Attribute VB_Name = "CatRemindTask"

Public Sub SetCustomFlag()
' Assign selected mail item to a category
' Create a task for follow up

Dim objMsg As Object
Dim objTask As Outlook.TaskItem

' GetCurrent Item function is at http://slipstick.me/e8mio
' Set objItem = objApp.ActiveInspector.CurrentItem

Set objMsg = GetCurrentItem()

If TypeOf objMsg Is MailItem Then

    With objMsg

' due this week flag
        .MarkAsTask olMarkThisWeek
' sets a specific due date
        .TaskDueDate = Now + 5

        .FlagRequest = UniConst(CheckOK) & " " & UniConst(XMark) & " " & objMsg.SenderName
        .ReminderSet = True
        .ReminderTime = Now + 4
' Now show Categories
        .ShowCategoriesDialog
'    .Display
        .Save
    End With
    
'Create a corresponding task for reminders & follow up
        Set objTask = Application.CreateItem(olTaskItem)

        With objTask
            .Subject = UniConst(RightArrow) & " " & UniConst(Clock2) & " " & objMsg.Subject
            .StartDate = objMsg.ReceivedTime
            .DueDate = objMsg.ReceivedTime + 5
            .Body = objMsg.Body
            .Categories = objMsg.Categories
            .ReminderSet = True
            .ReminderTime = Now + 4
            .Attachments.Add objMsg
            .Save
        
        End With
        
        Set objTask = Nothing
End If

Set objMsg = Nothing

End Sub

Function GetCurrentItem() As Object
    Dim objApp As Outlook.Application
           
    Set objApp = Application
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
    End Select
       
    Set objApp = Nothing
End Function



'---------------------------------------------------------------------------------------
' Procedure : UniConst
' Author    : Adam Waller
' Date      : 7/7/2020
' Purpose   : Search for characters: https://emojipedia.org/
'           : Look up UTF-16 Decimal value(s) from the following site:
'           : http://www.fileformat.info/info/unicode/char/search.htm
' Felix Reta added usage: Uniconst(LeftArrow,RightArrow,Clock2) as a string (2021)
'---------------------------------------------------------------------------------------
'
Public Function UniConst(Text As eUnicodeConst) As String

    Select Case Text
        Case LeftArrow:     UniConst = ChrW(8592)
        Case RightArrow:    UniConst = ChrW(8594)
        Case Clock2:        UniConst = ChrW(55357) & ChrW(56657)
        Case CheckOK:       UniConst = ChrW(&H2705)
        Case Ghost:         UniConst = ChrW(&HD83D) & ChrW(&HDC7B)
        Case MapleLeaf:     UniConst = ChrW(&HD83C) & ChrW(&HDF41)
        Case Calaca:        UniConst = ChrW(&H2620)
        Case Corazon:       UniConst = ChrW(&H2764)
        Case Finger1:       UniConst = ChrW(&H261D)
        Case HandBro:       UniConst = ChrW(&H270A)
        Case SailBoat:      UniConst = ChrW(&H26F5)
        Case Airplane:      UniConst = ChrW(&H2708)
        Case Watch:         UniConst = ChrW(&H231A)
        Case XMark:         UniConst = ChrW(10060)
        
    End Select

End Function

Public Sub CreateNewMessage()
Dim objMsg As MailItem

Set objMsg = Application.CreateItem(olMailItem)

 With objMsg
  .To = "Alias@domain.com"
  .CC = "Alias2@domain.com"
  .BCC = "Alias3@domain.com"
  .Subject = UniConst(Finger1) & " " & UniConst(XMark) & " " & UniConst(CheckOK) & " " & "El texto"
  .Categories = "Office"
  .VotingOptions = "Yes;No;Maybe;"
  .BodyFormat = olFormatHTML ' send plain text message
  .Importance = olImportanceHigh
  .Sensitivity = olConfidential
'  .Attachments.Add ("path-to-file.docx")

' Calculate a date using DateAdd or enter an explicit date
'  .ExpiryTime = DateAdd("m", 6, Now) '6 months from now
' .DeferredDeliveryTime = #8/1/2012 6:00:00 PM#
  
  .Display
End With

Set objMsg = Nothing
End Sub

Sub CopyItem()
 
 Dim myNameSpace As Outlook.NameSpace
 Dim myFolder As Outlook.Folder
 Dim myNewFolder As Outlook.Folder
 Dim myItem As Outlook.MailItem
 Dim myCopiedItem As Outlook.MailItem
 
 Set myNameSpace = Application.GetNamespace("MAPI")
 Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
 Set myNewFolder = myFolder.Folders.Add("Saved Mail", olFolderDrafts)
 Set myItem = Application.CreateItem(olMailItem)
 myItem.Subject = "Speeches"
 Set myCopiedItem = myItem.Copy
 myCopiedItem.Move myNewFolder
 
End Sub
Sub pruebita()

    Subject = "Coco" & " " & UniConst(Chile) & " " & "El texto"
    Debug.Print Subject
    
End Sub