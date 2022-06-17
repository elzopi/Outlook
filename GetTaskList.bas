Attribute VB_Name = "GetTaskList"
Public Sub GetListOfTasks()
' VBA Script that gets list of Outlook Tasks
' Use Tools->Macro->Security to allow Macros to run, then restart Outlook
' Run Outlook, Press Alt+F11 to open VBA
' Programming by Greg Thatcher, http://www.GregThatcher.com
    
    
    On Error GoTo On_Error
    Dim OLApp As Outlook.Application
    Dim Session As Outlook.Namespace
    Dim Report As String
    Dim TaskFolder As Outlook.Folder
    Dim currentItem As Object
    Dim currentTask As TaskItem
    
    Set OLApp = New Outlook.Application
    
    Set Session = OLApp.GetNamespace("MAPI")
    
    Set TaskFolder = Session.GetDefaultFolder(olFolderTasks)
    
    For Each currentItem In TaskFolder.Items
        If (currentItem.Class = olTask) Then
            Set currentTask = currentItem
            
            
            Report = Report & AddToReportIfNotBlank("ConversationTopic", currentTask.ConversationTopic)
            Report = Report & AddToReportIfNotBlank("ActualWork", currentTask.ActualWork)
            Report = Report & AddToReportIfNotBlank("AutoResolvedWinner", currentTask.AutoResolvedWinner)
            Report = Report & AddToReportIfNotBlank("BillingInformation", currentTask.BillingInformation)
            Report = Report & AddToReportIfNotBlank("Body", currentTask.Body)
            Report = Report & AddToReportIfNotBlank("CardData", currentTask.CardData)
            Report = Report & AddToReportIfNotBlank("Categories", currentTask.Categories)
            Report = Report & AddToReportIfNotBlank("Companies", currentTask.Companies)
            Report = Report & AddToReportIfNotBlank("Complete", currentTask.Complete)
            Report = Report & AddToReportIfNotBlank("ContactNames", currentTask.ContactNames)
            Report = Report & AddToReportIfNotBlank("ConversationIndex", currentTask.ConversationIndex)
            Report = Report & AddToReportIfNotBlank("CreationTime", currentTask.CreationTime)
            Report = Report & AddToReportIfNotBlank("DateCompleted", currentTask.DateCompleted)
            Report = Report & AddToReportIfNotBlank("DelegationState", currentTask.DelegationState)
            Report = Report & AddToReportIfNotBlank("Delegator", currentTask.Delegator)
            Report = Report & AddToReportIfNotBlank("DownloadState", currentTask.DownloadState)
            Report = Report & AddToReportIfNotBlank("DueDate", currentTask.DueDate)
            Report = Report & AddToReportIfNotBlank("EntryID", currentTask.EntryID)
            Report = Report & AddToReportIfNotBlank("Importance", currentTask.Importance)
            Report = Report & AddToReportIfNotBlank("InternetCodepage", currentTask.InternetCodepage)
            Report = Report & AddToReportIfNotBlank("IsConflict", currentTask.IsConflict)
            Report = Report & AddToReportIfNotBlank("IsRecurring", currentTask.IsRecurring)
            Report = Report & AddToReportIfNotBlank("LastModificationTime", currentTask.LastModificationTime)
            Report = Report & AddToReportIfNotBlank("MarkForDownload", currentTask.MarkForDownload)
            Report = Report & AddToReportIfNotBlank("MessageClass", currentTask.MessageClass)
            Report = Report & AddToReportIfNotBlank("Mileage", currentTask.Mileage)
            Report = Report & AddToReportIfNotBlank("NoAging", currentTask.NoAging)
            Report = Report & AddToReportIfNotBlank("Ordinal", currentTask.Ordinal)
            Report = Report & AddToReportIfNotBlank("OutlookInternalVersion", currentTask.OutlookInternalVersion)
            Report = Report & AddToReportIfNotBlank("OutlookVersion", currentTask.OutlookVersion)
            Report = Report & AddToReportIfNotBlank("Owner", currentTask.Owner)
            Report = Report & AddToReportIfNotBlank("Ownership", currentTask.Ownership)
            Report = Report & AddToReportIfNotBlank("PercentComplete", currentTask.PercentComplete)
            Report = Report & AddToReportIfNotBlank("ReminderOverrideDefault", currentTask.ReminderOverrideDefault)
            Report = Report & AddToReportIfNotBlank("ReminderPlaySound", currentTask.ReminderPlaySound)
            Report = Report & AddToReportIfNotBlank("ReminderSet", currentTask.ReminderSet)
            Report = Report & AddToReportIfNotBlank("ReminderSoundFile", currentTask.ReminderSoundFile)
            Report = Report & AddToReportIfNotBlank("ReminderTime", currentTask.ReminderTime)
            Report = Report & AddToReportIfNotBlank("ResponseState", currentTask.ResponseState)
            Report = Report & AddToReportIfNotBlank("Role", currentTask.Role)
            Report = Report & AddToReportIfNotBlank("Saved", currentTask.Saved)
            Report = Report & AddToReportIfNotBlank("SchedulePlusPriority", currentTask.SchedulePlusPriority)
            Report = Report & AddToReportIfNotBlank("SendUsingAccount", currentTask.SendUsingAccount)
            Report = Report & AddToReportIfNotBlank("Sensitivity", currentTask.Sensitivity)
            Report = Report & AddToReportIfNotBlank("Size", currentTask.Size)
            Report = Report & AddToReportIfNotBlank("StartDate", currentTask.StartDate)
            Report = Report & AddToReportIfNotBlank("Status", currentTask.Status)
            Report = Report & AddToReportIfNotBlank("StatusOnCompletionRecipients", currentTask.StatusOnCompletionRecipients)
            Report = Report & AddToReportIfNotBlank("StatusUpdateRecipients", currentTask.StatusUpdateRecipients)
            Report = Report & AddToReportIfNotBlank("Subject", currentTask.Subject)
            Report = Report & AddToReportIfNotBlank("TeamTask", currentTask.TeamTask)
            Report = Report & AddToReportIfNotBlank("ToDoTaskOrdinal", currentTask.ToDoTaskOrdinal)
            Report = Report & AddToReportIfNotBlank("TotalWork", currentTask.TotalWork)
            Report = Report & AddToReportIfNotBlank("UnRead", currentTask.UnRead)
            
            Report = Report & vbCrLf & vbCrLf
        End If
        
    Next
    
    
    Call CreateReportAsEmail("List of Tasks", Report)
    
Exiting:
        Exit Sub
On_Error:
    MsgBox "error=" & Err.Number & " " & Err.Description
    Resume Exiting
    
End Sub

Private Function AddToReportIfNotBlank(FieldName As String, FieldValue As String)
    AddToReportIfNotBlank = ""
    If (FieldValue <> "") Then
        AddToReportIfNotBlank = FieldName & " : " & FieldValue & vbCrLf
    End If
    
End Function

' VBA SubRoutine which displays a report inside an email
' Programming by Greg Thatcher, http://www.GregThatcher.com
Public Sub CreateReportAsEmail(Title As String, Report As String)
    On Error GoTo On_Error
    
    Dim Session As Outlook.Namespace
    Dim mail As MailItem
    Dim MyAddress As AddressEntry
    Dim Inbox As Outlook.Folder
    
    Set OLApp = New Outlook.Application
    
    Set Session = OLApp.GetNamespace("MAPI")
    Set Inbox = Session.GetDefaultFolder(olFolderInbox)
    Set mail = Inbox.Items.Add("IPM.Mail")
    
    Set MyAddress = Session.CurrentUser.AddressEntry
    mail.Recipients.Add (MyAddress.Address)
    mail.Recipients.ResolveAll
    
    mail.Subject = Title
    mail.Body = Report
    
'    mail.Save
    mail.Display
    
    
Exiting:
        Set Session = Nothing
        Exit Sub
On_Error:
    MsgBox "error=" & Err.Number & " " & Err.Description
    Resume Exiting

End Sub
