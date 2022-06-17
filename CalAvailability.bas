Attribute VB_Name = "CalAvailability"
Public Sub AvailabilityICS()

' From: https://www.slipstick.com/outlook/calendar/email-tomorrows-agenda/
' Send availability for the next 3 days to current opened email

Dim oNamespace As NameSpace
Dim oFolder As Folder
Dim oCalendarSharing As CalendarSharing
Dim objMail As MailItem ' As Inspector
Dim wd As Integer
Dim lDate As Date
Dim sDtate As Date

Set oNamespace = Application.GetNamespace("MAPI")
Set oFolder = oNamespace.GetDefaultFolder(olFolderCalendar)
Set oCalendarSharing = oFolder.GetCalendarExporter

' start date tomorrow
sDtate = Date + 1

' end date is 3 business days
' Sun = 1, Mon = 2, Tue = 3, Wed = 4, Thu = 5, Fri = 6, Sat = 7
wd = Weekday(Date)
If wd >= 1 And wd <= 3 Then
    lDate = sDtate + 2
ElseIf wd >= 4 Then
    lDate = sDtate + 4
End If

With oCalendarSharing
' options are olFreeBusyAndSubject, olFullDetails, olFreeBusyOnly
    .CalendarDetail = olFreeBusyOnly
    .IncludeWholeCalendar = False
    .IncludeAttachments = False
    .IncludePrivateDetails = False
    .RestrictToWorkingHours = True
    .StartDate = sDtate
    .EndDate = lDate
End With

SaveAsPath = "C:\Users\freta\AppData\Local\Temp\Availability from " & Format(sDtate, "mmm dd - ") & Format(lDate, "mmm dd yyyy") & ".ics"
oCalendarSharing.SaveAsICal SaveAsPath

' Maybe reply or new email instead of adding to the open email
Set objMail = Application.ActiveInspector.currentItem
 
 ' Send the mail item to the specified recipient.
 With objMail
  .Attachments.Add SaveAsPath
  .Display
 End With

Set oCalendarSharing = Nothing
Set oFolder = Nothing
Set oNamespace = Nothing
End Sub
