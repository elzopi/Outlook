VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Provides option to categorize and automatically create a task from an email just before sending it
' Code has to reside at ThisOutlookSession
' Requires class built to take advantage of vba autofill-enumerate when coding:
'
' List friendly names of Unicode characters
' Public Enum eUnicodeConst
'     RightArrow
'     LeftArrow
'     Clock2
' End Enum
' Added category selection prompt
' Added task creation with 4 days due date
' Added msgBox to indicate if email is confidential, defaulted to No
' Modified for Office 365, old application_startup + application_itemsend are no longer working

' Felix Reta June 2021, please credit my work if you use my code or part of it, it's just professional courtesy

Option Explicit

Private WithEvents olSentItems As Items
Attribute olSentItems.VB_VarHelpID = -1
Private WithEvents olCalendarItems As Items
Attribute olCalendarItems.VB_VarHelpID = -1
Public WithEvents myOlApp As Outlook.Application
Attribute myOlApp.VB_VarHelpID = -1

Sub Application_Startup()

  Initialize_handler

End Sub

Public Sub Initialize_handler()
' Added for Office 365 October 2021 Felix Reta

 Set myOlApp = Outlook.Application
 
End Sub

Private Sub myOlApp_ItemSend(ByVal Item As Object, Cancel As Boolean)
' Added for Office 365 October 2021 Felix Reta as per Microsoft sample:
' https://docs.microsoft.com/en-us/office/vba/api/outlook.application.itemsend

Dim xPrompt As String
Dim xOkOrCancel As Integer
Dim NewEmail As MailItem
Dim objTask As Outlook.TaskItem
Dim xYesNoorCancel As Integer
Dim xYesorNo As Integer

On Error Resume Next

If TypeOf Item Is MailItem Then

' If Item.Class = olMail Then
    Set NewEmail = Item

    xPrompt = "Do you want to process before sending the email?"
    xYesNoorCancel = MsgBox(xPrompt, vbYesNoCancel + vbExclamation + vbDefaultButton1)

    Select Case xYesNoorCancel

    Case vbYes
         xPrompt = "Is this Confidential?"
         xYesorNo = MsgBox(xPrompt, vbYesNo + vbExclamation + vbDefaultButton2)
         If xYesorNo = vbYes Then
            NewEmail.Sensitivity = olConfidential
         Else
            NewEmail.Sensitivity = olNormal
         End If
         
         NewEmail.ShowCategoriesDialog
'         NewEmail.MarkAsTask olMarkNextWeek 'or olmarkthisweek
'         NewEmail.TaskDueDate = Now + 5
'         NewEmail.FlagRequest = "Follow Up " & NewEmail.SenderName
'         NewEmail.ReminderSet = True
'         NewEmail.ReminderTime = Now + 4
    
'Create a corresponding task for reminders & follow up
        Set objTask = Application.CreateItem(olTaskItem)

        With objTask
            .Subject = UniConst(PeaceLove) & " " & UniConst(SailBoat) & " " & Item.Subject
            .StartDate = Item.ReceivedTime
            .DueDate = Item.ReceivedTime + 5
            .Body = Item.Body
            .Categories = Item.Categories
            .ReminderSet = True
            .ReminderTime = Now + 4
            .Attachments.Add NewEmail
            
            .Save
        
        End With
        
        Set objTask = Nothing

    Case vbNo
        Cancel = False
        
    Case vbCancel
        Cancel = True
        
End Select

    Set NewEmail = Nothing

End If

 
' Original sample code used to test for Office 365
' Dim prompt As String
' prompt = "Are you sure you want to send " & Item.Subject & "?"
' If MsgBox(prompt, vbYesNo + vbQuestion, "Sample") = vbNo Then
' Cancel = True
' End If
 
End Sub