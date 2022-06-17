Attribute VB_Name = "eMailCalendar"
Public Sub SendPrettyAgenda()
' From: https://www.slipstick.com/outlook/calendar/email-tomorrows-agenda/

Dim oNamespace As NameSpace
Dim oFolder As Folder
Dim oCalendarSharing As CalendarSharing
Dim objMail As MailItem
Dim wd As Integer

Set oNamespace = Application.GetNamespace("MAPI")
Set oFolder = oNamespace.GetDefaultFolder(olFolderCalendar)
Set oCalendarSharing = oFolder.GetCalendarExporter

' get the day - send sat/sun/monday out Fri night
' Sun = 1, Mon = 2, Tue = 3, Wed = 4, Thu = 5, Fri = 6, Sat = 7
' none set Sat/Sun
' wd = Weekday(Date)
'If wd >= 2 And wd <= 7 Then
'    lDays = Date + 1
'ElseIf wd = 1 Then
'    lDays = Date + 7
'End If

ldays = AddBusinessDays(Date, 5)

With oCalendarSharing
' options are olFreeBusyAndSubject, olFullDetails, olFreeBusyOnly
    .CalendarDetail = olFreeBusyAndSubject
    .IncludeWholeCalendar = False
    .IncludeAttachments = False
    .IncludePrivateDetails = True
    .RestrictToWorkingHours = False
    .StartDate = Date + 1
    .EndDate = ldays
End With

' prepare as email
' options: olCalendarMailFormatEventList, olCalendarMailFormatDailySchedule
Set objMail = oCalendarSharing.ForwardAsICal(olCalendarMailFormatDailySchedule)
 
 ' Send the mail item to the specified recipient.
 With objMail
 .Recipients.Add "felix.reta@gmail.com"
 .Subject = "TC Calendar for WE: " & ldays
' Remove the attached ics
 .Importance = olImportanceLow
 .Sensitivity = olPersonal
 .Categories = "Internet"
 .Attachments.Remove (1)
 .Display 'for testing, change to .send
 End With

Set oCalendarSharing = Nothing
Set oFolder = Nothing
Set oNamespace = Nothing
End Sub

Function AddBusinessDays(StartDate As Date, numberOfDays As Integer) As Date

    Dim newDate As Date
    
    newDate = StartDate
' Possibly need to add Reference to MS Excel
    AddBusinessDays = WorksheetFunction.WorkDay(newDate, numberOfDays)

End Function


Sub CreateListofAppt()
   
   Dim CalFolder As Outlook.MAPIFolder
   Dim CalItems As Outlook.Items
   Dim ResItems As Outlook.Items
   Dim sFilter, strSubject, strAppt As String
   Dim iNumRestricted As Integer
   Dim itm, apptSnapshot As Object
   Dim tStart As Date, tEnd As Date, tFullWeek As Date
   Dim wd As Integer
  
   ' Use the default calendar folder
   Set CalFolder = Session.GetDefaultFolder(olFolderCalendar)
   Set CalItems = CalFolder.Items

   ' Sort all of the appointments based on the start time
   CalItems.Sort "[Start]"
   CalItems.IncludeRecurrences = True

   ' Set an end date
    tStart = Format(Date + 1, "Short Date")
    tEnd = Format(Date + 7, "Short Date")
    tFullWeek = Format(Date + 6, "Short Date")
 
    wd = Weekday(Date)
   ' Sun = 1, Mon = 2, Tues = 3, Wed = 4, Thu = 5, Fri = 6, Sat = 7
' get next day appt, do whole week on sunday
If wd >= 2 And wd <= 6 Then
   sFilter = "[Start] >= '" & tStart & "' AND [Start] <= '" & tEnd & "'"
ElseIf wd = 1 Then
   sFilter = "[Start] >= '" & tStart & "' AND [Start] <= '" & tFullWeek & "'"
End If

Debug.Print sFilter
   Set ResItems = CalItems.Restrict(sFilter)

   iNumRestricted = 0

   'Loop through the items in the collection.
   For Each itm In ResItems
   Debug.Print ResItems.Count
      iNumRestricted = iNumRestricted + 1
      
 ' Create list of appointments
  strAppt = strAppt & vbCrLf & itm.Subject & vbTab & " >> " & vbTab & itm.start & vbTab & " to: " & vbTab & Format(itm.End, "h:mm AM/PM")

   Next
   
' After the last occurrence is checked
' Open a new email message form and insert the list of dates
  Set apptSnapshot = Application.CreateItem(olMailItem)
  With apptSnapshot
    .Body = strAppt & vbCrLf & "Total appointments; " & iNumRestricted
    .To = "me@slipstick.com"
    .Subject = "Appointments for " & tStart
    .Display 'or .send
  End With

   Set itm = Nothing
   Set apptSnapshot = Nothing
   Set ResItems = Nothing
   Set CalItems = Nothing
   Set CalFolder = Nothing
   
End Sub
