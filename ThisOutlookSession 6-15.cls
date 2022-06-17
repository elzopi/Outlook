VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' This code must reside @ ThisOutlookSession
' (C) Designed and developed by Felix Reta, Epsitel Corporation (R), Feb-May 2014
' felix.reta@gmail.com
' Objective: Send an email with 2 calendar attachments (.ics & Outlook .msg)
' to a separate email address(es) to maintain calendars synchronized.
' Reason: In some corporate environments, you are procluded to share your calendar
'  for security settings.
' High level methodology:
'   1) On Application Startup, get default calendar pointer for processing
'   2) Main code (ItemAdd) is called whenever a calendar entry is added
'   3) Depending on the calendar entry status indicators an HTML email is built
'       with 2 attachments (VCal (.ics) and Outlook cal (.msg))
'   4) Email is sent to hard coded email addresses for manual update
'  Code is sparsely commented, please refer to it for details.
' Most of the data mapping was obtained by Googling it, most useful information was found at
' Outlookcode.com by Sue Mosher, great site for Outlook developers
' Loosely based on Googled information from the following sites:
' http://www.slipstick.com/developer/send-email-when-you-add-appointment-to-calendar/
' http://www.slipstick.com/developer/windows-environment-variables-outlook-macros/
' http://www.mrexcel.com/forum/excel-questions/414424-create-outlook-appointment-ics-file-excel.html
' If you use my code, be professional and credit me, Thanks
' Need following references:
'  Visual Basic for Applications
'  MS Outlook Object Library
'  OLE Automation
'  MS Office Object Library
'  Microsoft XML
'  Microsoft VBscript Regular Expressions
'  Microsoft Excel for WorkDay function

' Version 3
'        inserts an entry into Google Calendar using a subroutine
'        inserts Company Name from session information
'        obtains SMTP email address for sub and appointment info consumption
'        obtains Google account information via XML calls to config file
'       3.4
'       Included RC4 hashing password retrieval.
'       Via sendToGcal switch
' Requires TextHelpers, CryptoHelpers (for RC4 module) and XMLHelpers modules
' And clsCryptoFilterBox Class Module (in case M$ Crypto is used)
' The location of the XML file LasPws.xml needs to be manually updated at the
' GetAcctInfoFromXML Subroutine within XMLHelpers
' Ver 3.5
' Included XML code to read company paramenters (Name, To Email)
' Ver 3.6
' Corrected recurring event not saved as olVcal (.vcs), changed to olIcal (.ics) for built attachment
' Ver 3.7
' Exploring recurrence treatment
' Updated Organizer field as it appears to be empty when copied to main Calendar
' Ver 3.8
' Test for appointments in the past, don't process
' Normalized location of pws file @ XMLHelpers to My Documents
' Ver 3.9 Oct 2014
' Included test for copied items (e.g. from GCal feed), these weren't being processed as they came flagged
' as Not Responded thus were not processed originally
' Ver 4.0 Nov 2014
' Included Reminders firing capture
' Included Check to Ignore (M$ conferencing inserts bogus Calendar entry) - Amex specific
' Converted to Function at TextHelper = excludeText
' Included code for Inspectors to catch Follow Up reminders
' Requires clsInspectorHandler
' Corrected hardcoded To and CC for updated/changed messages
' Ver 4.1 Feb 2022
' Adapted to XML 6.0 variable names
' Adapted to Windows 10 user file structure from My Documents to Documents
' Converted original itemAdd code into a Sub and minimized:
' ItemAdd, ItemChange and DeletedItems_ItemAdd, the latter to try to catch deleted appointments/meetings
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
' Included regexp to determine 5 or 6 digits on the subject, specific to Alvaria project and customer numbers
'  used to assign category and move to project folder under projects
'  included addBusinessDays function
' Ignoring meeting responses received on Inbox, these caused duplicate outgoing emails sent for the same appointment
' Added AddWeekDays function

' Felix Reta March 2022, please credit my work if you use my code or part of it, it's just professional courtesy


Dim WithEvents curCal As Items
Attribute curCal.VB_VarHelpID = -1
Dim WithEvents DeletedItems As Items
Attribute DeletedItems.VB_VarHelpID = -1
Dim WithEvents myInbox As Items
Attribute myInbox.VB_VarHelpID = -1

' Dim WithEvents olSentItems As Items
Dim indInbox As Integer
Dim strTime As String


' Dim newCalFolder As Outlook.Folder
Public Sub Initialize_handler()
' Added for Office 365 October 2021 Felix Reta

' Set myOlApp = Outlook.Application
 
End Sub

Private Sub Application_Startup()
   Dim NS As Outlook.NameSpace
   Dim SentItems As Outlook.Folder

   Initialize_handler

   Set NS = Application.GetNamespace("MAPI")
   
' check inbox arrival for appointments processing, when using calendar folder, outlook would indicate changes on its own
   Set myInbox = NS.GetDefaultFolder(olFolderInbox).Items
' calendar to watch for new items
   Set curCal = NS.GetDefaultFolder(olFolderCalendar).Items
' watch deleted folder
Set DeletedItems = NS.GetDefaultFolder(olFolderDeletedItems).Items
' watch sent folder
' Set SentItems = NS.GetDefaultFolder(olFoldersentItems).Items
' calendar moving copy to
'   Set newCalFolder = GetFolderPath("data-file-name\calendar")
   Set NS = Nothing
End Sub
  
Private Sub myInbox_ItemAdd(ByVal Item As Object)

Dim sTypeName As String
Dim theAppointment As Outlook.AppointmentItem
Dim theMeeting As Outlook.MeetingItem
Dim inItem As Object
' Dim strTime As String
Dim ProjColl As MatchCollection
Dim ProjClient As Match
Dim bIsProject As Boolean
Dim theMClass As String
Dim strSubject As String
Dim strFrom As String
Dim bProcAppt As Boolean

' Setting the stage to automate Category and potentially folder (item copy to) if a 6 or 5 digit number is detected on the subject
' Using Regexp
' Adding regexp functionality to identify a project related email 6 digits and space, or customer number 5 digits and space
'    "\d{5,6}\x20"gm Caused an issue, simplified to \d{6} for Project Number, 8 digits for Case Numbers (WIP)
' Subject responses observed after a meeting was sent
' Accepted:
' New Time Proposed:
' Tentative:


Set inItem = Item
strSubject = Item.Subject
Set ProjColl = findProjects(strSubject)
bIsProject = False
bProcAppt = False
strTime = Now
sTypeName = TypeName(Item)

' WIP Check for resource assignments ProjectOfficeNA@Alvaria.com (Mary Avila as well)
' and subject Resource Allocation Approved
'    (Resource: lastname, firstname)
' and body contains Resource:   Project: nnnnnn    Hours:    Resource Type:

' account for replies for revised appointments, these are received back as MailItems with a link
' to outlook to accept/tentative/decline IT IS NOT AN APPOINTMENT
' When tentative, Item is added t   o the calendar curCa_ItemAdd
' also 3 other ItemChanges are fired with an AppointmentItem
' when a meeting is snoozed/dismissed, it also fires 3 Item changes
' when a meeting is declined, gets to DeletedItems ItemAdd as Meeting Received-Declined and also
' on ItemChange
' When an inbox invite is read, it fires curcal_itemchange which might mean it is ready to be written
' to the calendar, or actually written as "draft"
' sTypeName ReportItem bombs

If sTypeName <> "MeetingItem" Then
    Debug.Print strTime & " Type:" & sTypeName & " not an Meeting item, ignored"
    Exit Sub
Else


strFrom = Item.SenderEmailAddress

Debug.Print strTime & " Inbox ItemAdd Type: " & sTypeName & "-Subject: " & strSubject

' ReportItem

Select Case Left(strSubject, 9)

    Case "Accepted:"
        bProcAppt = False
    
End Select

' Collection has at least 1 item, process accordingly
If ProjColl.Count > 0 Then

    For Each ProjClient In ProjColl
        Debug.Print ProjClient
        Debug.Print strTime & " Project/Client found on subject: " & vbCrLf & ProjClient.value
    Next ProjClient
    bIsProject = True
Else
    Debug.Print strTime & " No project/client digits identified on this item"
    bIsProject = False
End If



   
    Debug.Print strTime & " Inbox-" & sTypeName & " request arrived "
    ' at this point a message arriving at inbox is an appointment
    Set theMeeting = Item
    ' Identify responses, potentially with no need to be processed
    
' Meeting Message classes:
' IPM.Schedule.Meeting.Resp.Neg for a decline response
' IPM.Schedule.Meeting.Resp.Pos for an acceptance response
' IPM.Schedule.Meeting.Resp.Ten for a tentative response
' IPM.Schedule.Meeting.Canceled for a cancelled meeting from organizer
'    counter proposal pending
' Ignore *.Resp.* items to ProcessAppt

    theMClass = theMeeting.MessageClass
    Debug.Print "Message Class: " & theMClass ' to identify responses and ignore, to avoid dupe emails out
        
    ' Do Not process responses, avoid dupe external emails
'    If Left(theMClass, 26) = "IPM.Schedule.Meeting.Resp." Then Exit Sub
' watch out for Cancellations, they show as Meeting.Canceled without .Resp
    Select Case theMClass
    
        Case InStr(1, theMClass, ".Resp")
            Exit Sub
        
        Case InStr(1, theMClass, ".Canceled") ' need to work on the data being sent to process
            Debug.Print strTime & " Meeting Subject is: " & strSubject & "-" & theMClass
            Debug.Print "Meeting Body is: " & theMeeting.Body
            Call ProcessAppt(theMeeting, "Canceled from Inbox")
'            Exit Sub
    
'    If InStr(1, theMClass, ".Resp") Then Exit Sub
        Case Else
    
            indInbox = 1 'indicator that meeting was seen on Inbox
    
    ' need to get the associated Appointment before calling add/change routines
            Debug.Print strTime & " Meeting Subject is: " & strSubject
            Debug.Print "Meeting Body is: " & theMeeting.Body
            Set theAppointment = theMeeting.GetAssociatedAppointment(True)
            Debug.Print "Appointment data: Start: " & theAppointment.start
            Debug.Print "End: " & theAppointment.End
            Debug.Print "Subject, Body: " & strSubject & " " & theAppointment.Body
    ' Success, at this point we have sufficient data for the appointment to build an outgoing email with it
    ' However, ALL INCOMING appointments will be processed, new, cancelled, proposed times
    ' need to account for MY issued appointments, resolved by analyzing message class containing "Resp" earlier in the code Feb 2022
            Debug.Print "Appointment status: " & theAppointment.ResponseStatus
    
            Call ProcessAppt(theAppointment, "from Inbox")
        
        End Select
 End If


' Application.NewMailEx event Occurs when a new item is received in the Inbox.
' Use the GetAssociatedAppointment method to return the AppointmentItem object associated with a MeetingItem object,
' and work directly with the AppointmentItem object to respond to the request.



End Sub

Private Sub curCal_ItemAdd(ByVal Item As Object)
Dim cAppt As AppointmentItem
Dim moveCal As AppointmentItem
' Dim strTime As String

strTime = Now

Debug.Print strTime & " curCal_ItemAdd " & TypeName(Item) & " Subject: " & Item.Subject

' TypeName(Item)="MeetingItem" instead of typeof item is MeetingItem, not working
If TypeName(Item) = "AppointmentItem" Then
   Set cAppt = Item
   Debug.Print "Organizer: " & cAppt.Organizer & " Class: " & cAppt.ItemProperties.Class
' Check if I am organizing it
   Set cAppt = Item
       If Right(cAppt.Organizer, 5) = "Felix" Then
          Call ProcessAppt(Item, "Added by Me")
       Else
          Exit Sub
       End If
End If
   
' On Error Resume Next

'remove to make a copy of all items
'If Item.BusyStatus = olBusy Then
'
'   Item.Body = Item.Body & "[" & GetGUID & "]"
'   Item.Save
   
' Set cAppt = Application.CreateItem(olAppointmentItem)

'With cAppt
'    .Subject = "Copied: " & Item.Subject
'    .Start = Item.Start
'    .Duration = Item.Duration
'    .Location = Item.Location
'    .Body = Item.Body
'End With
 
' set the category after it's moved to force EAS to sync changes
' Set moveCal = cAppt.Move(newCalFolder)
 'moveCal.Categories = "moved"
 'moveCal.Save
 
'    End If
 
 End Sub
 
 
Private Sub curCal_ItemChange(ByVal Item As Object)

' Be aware that this will be entered when anything changes in the calendar, including snooze

'Dim cAppt As AppointmentItem
Dim objAppointment As AppointmentItem
Dim sTypeName As String

strTime = Now

sTypeName = TypeName(Item)

If sTypeName <> "AppointmentItem" Then Exit Sub
    Debug.Print strTime & " ItemChange Type: " & sTypeName

' Call ProcessAppt(item, "Changed")

' On Error Resume Next
      
' use 2 + the length of the GUID
' strBody = Right(Item.Body, 38)
 
'For Each objAppointment In newCalFolder.Items
' If InStr(1, objAppointment.Body, strBody) Then
'         Set cAppt = objAppointment
'   End If
' Next

  
'With cAppt
'    .Subject = "Copied: " & Item.Subject
'    .Start = Item.Start
'    .Duration = Item.Duration
'    .Location = Item.Location
'    .Body = Item.Body
'    .Save
'End With
  
End Sub

Private Sub DeletedItems_ItemAdd(ByVal Item As Object)

Dim sTypeName As String
' Dim strTime As String
Dim bIsProject As Boolean
Dim mtgItem As MeetingItem
Dim apptItem As AppointmentItem

' only apply to appointments

sTypeName = TypeName(Item)
strTime = Now

Dim ProjColl As MatchCollection
Dim ProjClient As Match

' Setting the stage to automate Category and potentially folder (item copy to) if a 6 or 5 digit number is detected on the subject
' Using Regexp
' Adding regexp functionality to identify a project related email 6 digits and space, or customer number 5 digits and space
'    "\d{5,6}\x20"gm

bIsProject = False

Set ProjColl = findProjects(Item.Subject)

' Collection has at least 1 item, process accordingly
If ProjColl.Count > 0 Then
    Debug.Print strTime & " Projects or clients found in Subject" & vbCrLf
    For Each ProjClient In ProjColl
        Debug.Print ProjClient
    bIsProject = True
    Next ProjClient
End If


Debug.Print strTime & " DeletedItems ItemAdd type: " & sTypeName

' MeetingItem identified when proposing new time, makes sense, need to account for it
' "New Time Proposed:" [then old subject] is the left part of the subject. When email
' is received by the originator, has an outlook link to his/her calendar to review, edit
' and send the revised appointment
' MeetingItem is also the type for Deleting items from the Calendar (e.g. meeting cancellation received and pressing "Delete from Calendar"
' Need to account for different situations that make it to the Deleted folder:
' When accepting, proposing new time or rejecting

Select Case sTypeName

Case "MeetingItem"  ' watch out, if you delete in inbox will trigger an incorrect object, ProcessAppt bombs out
    Set mtgItem = Item
    Debug.Print "Meeting Sender: " & mtgItem.SenderName & " Msg Class: " & mtgItem.MessageClass
    Call ProcessAppt(mtgItem, "Deleted")
    
Case "AppointmentItem"
    Set apptItem = Item
    Debug.Print "Meeting Status: " & apptItem.MeetingStatus & "Organizer: " & apptItem.Organizer
    
    Call ProcessAppt(apptItem, "Deleted")

Case Else
    Debug.Print sTypeName & " Not processed"

End Select

' If sTypeName <> "AppointmentItem" Then Exit Sub


End Sub
 
 Private Sub application_ItemSend(ByVal Item As Object, Cancel As Boolean)
' Added for Office 365 October 2021 Felix Reta as per Microsoft sample:
' https://docs.microsoft.com/en-us/office/vba/api/outlook.application.itemsend
' We'll create a task asssociated with the email IF we decide to process it
' We are also checking for appointments set by me


Dim xPrompt As String
Dim xOkOrCancel As Integer
Dim NewEmail As MailItem
Dim objTask As Outlook.TaskItem
Dim oMeeting As Outlook.MeetingItem
Dim xYesNoorCancel As Integer
Dim xYesorNo As Integer
Dim apptItem As AppointmentItem
Dim bIsProject As Boolean
Dim strSubject As String
Dim strBody As String

' Dim strTime As String

On Error Resume Next

strTime = Now
bIsProject = False

Debug.Print strTime & " App ItemSend:"
Dim ProjColl As MatchCollection
Dim ProjClient As Match

' Setting the stage to automate Category and potentially folder (item copy to) if a 6 or 5 digit number is detected on the subject
' Using Regexp
' Subject is available for several types (email, appt, meeting)
' bIsProject switch to indicate a copy to the corresponding project bin

strSubject = Item.Subject

Set ProjColl = findProjects(Item.strSubject)

' Collection has at least 1 item, process accordingly
If ProjColl.Count > 0 Then
    Debug.Print strTime & "-" & ProjColl.Count & " Projects or clients found in Subject"
    For Each ProjClient In ProjColl
        Debug.Print ProjClient
    Next ProjClient
    Item.Categories = "Projects"
    bIsProject = True
    
End If

Select Case TypeName(Item)

Case "MailItem"
    Set NewEmail = Item

    strSubject = NewEmail.Subject
        
    xPrompt = "Do you want to process before sending the email?"
    xYesNoorCancel = MsgBox(xPrompt, vbYesNoCancel + vbExclamation + vbDefaultButton1)

    Select Case xYesNoorCancel

    Case vbYes
'         xPrompt = "Is this Confidential?"
'         xYesorNo = MsgBox(xPrompt, vbYesNo + vbExclamation + vbDefaultButton2)
'         If xYesorNo = vbYes Then
'            NewEmail.Sensitivity = olConfidential
'         Else
'            NewEmail.Sensitivity = olNormal
'         End If
    
         With NewEmail
            .MarkAsTask olMarkNextWeek
'            .FlagRequest = ""   ' might cause recipient to get a flagged email
            .ReminderSet = True
            .ReminderTime = Now + 4
            .ShowCategoriesDialog
            .Save
         End With
         
'         NewEmail.MarkAsTask (olMarkNextWeek)
'         NewEmail.ShowCategoriesDialog
'         NewEmail.MarkAsTask olMarkNextWeek 'or olmarkthisweek
'         NewEmail.TaskDueDate = Now + 5
'         NewEmail.FlagRequest = "Follow Up " & NewEmail.SenderName
'         NewEmail.ReminderSet = True
'         NewEmail.ReminderTime = Now + 4
'         NewEmail.Save

'Create a corresponding task for reminders & follow up
        Set objTask = Application.CreateItem(olTaskItem)

        With objTask
            .Subject = UniConst(PeaceLove) & " " & UniConst(SailBoat) & " " & strSubject & " " & UniConst(Airplane) & UniConst(Clock2)
'           .Subject = UniConst(Finger1) & " " & strSubject
'            .Subject = Item.Subject
            .StartDate = Item.ReceivedTime
            .DueDate = AddWeekDays(.StartDate, 5)
            strBody = "To(1): " & Item.Recipients(1) & " On: " & Now() & vbCrLf & String(20, UniConst(CheckOk)) & vbCrLf
            
            .Body = strBody & Item.Body
            .Categories = Item.Categories
            .ReminderSet = True
            .ReminderTime = AddWeekDays(.StartDate, 5)
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


Case "MeetingItem"
     Set oMeeting = Item
   
     Debug.Print strTime & " On send, type: " & TypeName(Item) & " Organized by: " & oMeeting.SenderEmailAddress
     Debug.Print "Subject: " & strSubject
    
     Call ProcessAppt(Item, "From ItemSend " & UniConst(PeaceLove))


Case "AppointmentItem"
     Set apptItem = Item
     Debug.Print strTime & " On send, type: " & TypeName(Item) & " Organized by: " & oMeeting.SenderEmailAddress
     Debug.Print "Subject: " & strSubject
     Call ProcessAppt(Item, "From ItemSend " & UniConst(Handbro))

Case Else
     Debug.Print strTime & " Not processed: " & TypeName(Item) & " Subject: " & Item.Subject
    
End Select
 
 
End Sub

Private Sub ProcessAppt(ByVal Item As Object, tFrom As String)

' (C) Designed and developed by Felix Reta, Epsitel Corporation (R), Feb-May 2014
' felix.reta@gmail.com
' Most of the data mapping was obtained by Googling it, most useful information was found at
' Outlookcode.com by Sue Mosher, great site for Outlook developers
' If you use my code, be professional and credit me, Thanks

  Dim strID As String
  Dim olNS As Outlook.NameSpace
  Dim oMeetingItem As Outlook.MeetingItem
  Dim oResponse As Outlook.MeetingItem
  Dim oAppointment As Outlook.AppointmentItem
  Dim iMeeting As Outlook.MeetingItem
  Dim objMsg As MailItem
  Dim enviro As String
  Dim strFilePath As String
  Dim iSensitivity As Integer
  Dim strBody As String
  Dim strStatus As String
  Dim strHTMLHeader As String
  Dim iStatus As Integer
  Dim mStatus As Integer
  Dim lProcess As Boolean
  Dim strHTMLcode As String
  Dim sHostName As String
  Dim strUser As String
  Dim strCompany As String
  Dim currentItem As Object
  Dim currentContact As ContactItem
  Dim oExUser As Outlook.ExchangeUser
  Dim TZOffset As String
  Dim sendToGCal As Boolean
  Dim elArchivo As String
  Dim LaCuenta As String
  Dim OID As String
  Dim OPW As String
  Dim LaPW As String
  Dim ElID As String
  Dim strTypeReq As String
'  Dim cCrypto As clsCryptoFilterBox ' Need Class Module clsCryptoFilterBox
  Dim ToEmail As String
  Dim ccEmail As String
  Dim strRecurr As String
  Dim Organisateur As String
  Dim strSubject As String
  Dim strLocation As String
  Dim dtStart As Date
  Dim dtEnd As Date
  Dim lAppt As Boolean
 ' Dim strTime As String
  
' By debugging, found out that for an attached .ics invite, the sequence for oAppointment.ResponseStatus
' is 5 (olResponseNotResponded) when opening the attachment
' and then 3 (olResponseAccepted) when clicking the Accept button
' appt_Write was not called under this condition (.ics opened)
' Feb 2014
  Set olNS = Application.Session
  Set oExUser = olNS.CurrentUser.AddressEntry.GetExchangeUser
  Email = oExUser.PrimarySmtpAddress
  
  strTime = Now
  strSubject = Item.Subject
  
  Debug.Print strTime & " at ProcessApt " & tFrom
  
  inProgress = True
  sendToGCal = False ' prep for eventual XML param read options
  
' Included for v2.2 to generate HTML emails to potentially use code from a file under My Documents
' default library

  strHTMLHeader = "<div class=WordSection1 STYLE=""font-family: Tahoma; font-size: 24px; color: pink ""><p class=MsoNormal><o:p>" & vbCrLf

' Sets up the temporary calendar attachment (.ics) file to be included in the generated email
' Original code forwarded the invite with disastrous results, as the organizer calendar entry
' included the external email used in the forward, ooopppsss!!!!

' Get Host Name / Get Computer Name Mar-2014
  sHostName = Environ$("computername")
' Get Temporary folder name
  enviro = CStr(Environ("TEMP"))
  strFilePath = enviro & "\Outlook Appointment"
  
' Sets up the potential use of an external file for HTML format
  enviro = CStr(Environ("UserProfile"))
  strHTMLcode = enviro & "\Documents\" & "HTML-OL-header.txt"
  
   strID = Item.EntryID
   strUser = Environ$("username")
   strUser = strUser & "-" & ThisOutlookSession.Session.CurrentUser.Name

' Build Company variables OID will contain the To email address and OPW the Company name
   enviro = CStr(Environ("UserProfile"))
   elArchivo = enviro & "\Documents\" & "LasPWs.xml"
   LaCuenta = "ALV"
   OID = ""
   OPW = ""
   strTypeReq = "Company"
   Call GetAcctInfoFromXML(strTypeReq, elArchivo, LaCuenta, OID, OPW)
   ToEmail = OID
' Modify for cc'd email if any
   ccEmail = "felix.reta@live.com"
' Get the Company name if available, or default it

Debug.Print "environ: " & enviro

Debug.Print "Datos de XML: " & OID & " " & OPW

strCompany = oExUser.CompanyName
If strCompany = "" Then
   strCompany = OPW
End If

' MeetingStatus constants:
' From: http://msdn.microsoft.com/en-us/library/office/ff869427.aspx
'   Name                            Value               Description
'   olMeeting                       1          The meeting has been scheduled.
'   olMeetingCanceled               5          The scheduled meeting has been cancelled.
'   olMeetingReceived               3          The meeting request has been received.
'   olMeetingReceivedAndCanceled    7          The scheduled meeting has been cancelled but still appears on the user's calendar.
'   olNonMeeting                    0          An Appointment item without attendees has been scheduled. This status can be used to set up holidays on a calendar.
    
' OlResponseStatus constants
' From: http://msdn.microsoft.com/en-us/library/office/ff868658.aspx
'   Name                    Value   Description
'   olResponseAccepted      3       Meeting accepted.
'   olResponseDeclined      4       Meeting declined.
'   olResponseNone          0       The appointment is a simple appointment and does not require a response.
'   olResponseNotResponded  5       Recipient has not responded.
'   olResponseOrganized     1       The AppointmentItem is on the Organizer's calendar or the recipient is the Organizer of the meeting.
'   olResponseTentative     2       Meeting tentatively accepted.

    
' Currently assuming received item is an appointment, need to account for meeting items sent by me from itemsend routine

lAppt = False

If TypeName(Item) = "AppointmentItem" Then
    lAppt = True
    Set oAppointment = Item
    iStatus = oAppointment.ResponseStatus
'    strSubject = oAppointment.Subject
    strLocation = oAppointment.Location
    dtStart = oAppointment.start
    dtEnd = oAppointment.End
    Organisateur = oAppointment.Organizer
    oAppointment.SaveAs strFilePath & ".ics", olICal 'or olVcal flag (vcal has issues with recurring appts.)
    oAppointment.SaveAs strFilePath & ".msg"

    MsgBox ("email sent for " & Organisateur & " meeting request" & "Action: " & tFrom)
    
' Meeting Status is updated whenever a meeting request is opened, empirically:
' 1 is present as received, if it is responded (e.g. ResponseStatus set) it is set accordingly.
' 5 is set when the meeting is opened (e.g. written to the calendar) and the organizer cancelled it
' 3 yet to be observed

' 0 & 0 (both flags) are set as soon as I start setting an appointment, and both set
' local appointments (e.g. penciled in)
' 1 & 1 (both flags) are set as soon as I save in my calendar without sending it
' An opportunity to grab the ThisOutlookSession.Session.CurrentUser.Name
' if so desired or .Address for X.400 address, not thoroughly tested,
' .Name works in Outlook 2010
'  oAppointment.Organizer = ThisOutlookSession.Session.CurrentUser.Name

' oAppointment.GetRecurrencePattern contains the recurrence pattern of the appointment
' For meetings still WIP as it bombs when tryng to save ics

    mStatus = oAppointment.MeetingStatus
    lProcess = True
    strStatus = "Appointment "
' Test for appointment in the past, don't process.
    If oAppointment.start < Now Then
        lProcess = False
    End If

    If oAppointment.RecurrenceState = olApptNotRecurring Then
        Debug.Print "Non recurring appointment"
        strRecurr = ""
' For v2.2 hope this makes it as part of the .ics
        oAppointment.Categories = "Internet"  'recurring appointment does not support
' Adjust field to your company name (I haven't seen this field making it to
' the calendar .msg type
' bombed on recurring appts
        oAppointment.Companies = strCompany
    Else
        Debug.Print "Recurring appointment identified"
        strRecurr = "Recurring "
    End If
    
Else

    ' account for cancelations, no associated appointment will be found
    ' Subject contains Canceled:
    ' Message Class is IPM.Schedule.Meeting.Canceled
     Debug.Print "Catching Canceled item, possibly: " & TypeName(Item)
        
    
    
    Set iMeeting = Item
    If iMeeting.CreationTime < Now Then
        lProcess = False
    End If
    
    If iMeeting.MessageClass = "IPM.Schedule.Meeting.Canceled" Then
        strLocation = "Loc TBD"
        dtStart = iMeeting.CreationTime
        dtEnd = dtStart + 4
        
    Else
    
        Set oAppointment = iMeeting.GetAssociatedAppointment(False)
    
        strLocation = oAppointment.Location
        dtStart = oAppointment.start
        dtEnd = oAppointment.End
    
    End If
    
    iStatus = 99
    mStatus = 99
    lProcess = True
    strStatus = "Meeting "
    iMeeting.Categories = "Internet"
    iMeeting.Companies = strCompany
    
    ' iMeeting.SenderEmailAddress = "Mili@stones.pw" ' read only attribute
    
    iMeeting.SaveAs strFilePath & ".doc", olRTF  ' olIcal bombs out for recurring appts
    iMeeting.SaveAs strFilePath & ".msg"
    
End If

Debug.Print strUser & "-" & sHostName & "-Item Entry ID: " & Right(strID, 12) & " Meeting Status at: " & tFrom & " is: " & mStatus

Select Case mStatus
    Case 0
        strStatus = strStatus & "Local save"
    Case 1
        strStatus = strStatus & "Scheduled"
    Case 3
        strStatus = strStatus & "Received"
    Case 5
        strStatus = strStatus & "Cancelled"
    Case 7
        strStatus = strStatus & "Received & Cancelled"
    Case 99
        strStatus = strStatus & "Meeting sent by me"
    
    Case Else
        strStatus = strStatus & "INVESTIGATING-" & oAppointment.MeetingStatus
        lProcess = False
End Select

Select Case iStatus
' This is set when you save an appointment locally (e.g. no invitees or blocking time)
    Case 0
        strStatus = strStatus & "-" & "Penciled in"
    Case 1
        strStatus = strStatus & "-" & "Organized"
    Case 2
        strStatus = strStatus & "-" & "Tentative"
        lProcess = False   ' Trying to avoid duplicate outgoing emails when the message is from an invitee
    Case 3
        strStatus = strStatus & "-" & "Accepted"
        lProcess = False   ' Trying to avoid duplicate outgoing emails when the message is from an invitee
    Case 4
        strStatus = strStatus & "-" & "Declined"
        lProcess = False   ' Trying to avoid duplicate outgoing emails when the message is from an invitee
    Case 5
' This is set right after an invite shows up in the Inbox....
        strStatus = strStatus & "-" & "Not responded"
    Case 99
        strStatus = strStatus & "-" & "Meeting request sent by me"
        
' Catch a copied appointment (e.g. from GCal using CopyToMyCal routine)
'        lProcess = excludeText(oAppointment.Subject)

'        If Left(oAppointment.Subject, 5) = "Copy:" Then
'           lProcess = True
'        Else
'           lProcess = False
'        End If
    Case Else
        strStatus = strStatus & "-" & "Unknown"
        lProcess = False
End Select


'   set oAppointment = Item.GetAssociatedAppointment(True)

Debug.Print strUser & "-" & sHostName & "-Item Entry ID: " & Right(strID, 12) & " Meeting Response Status at: " & tFrom & " is: " & strStatus

' Included If for v2 as 2 emails were generated when opening attached .ics to process only if
' appointment was specifically accepted
' Switch set to be tested depending on ResponseStatus flag Case statements

' lProcess = excludeText(oAppointment.Subject)

If lProcess Then

' Set of code to read a text file with HTML code, not tested
' Set fsi = CreateObject("Scripting.FileSystemObject")
' Set theFile = fsi.OpenTextFile(strHTMLCode, forreading, tristatefalse)
' Stationery located here:
' %appdata%\Microsoft\Stationery
' Signatures located here:
' %appdata%\Microsoft\Signatures

' Customize at string for your company name
' This builds an RTF (HTML maybe) invite for Outlook

        strBody = "<p>" & strRecurr & "Appointment " & strStatus & strCompany & sHostName & "</p>" & "<p>Organized by: " & Organisateur & " calendar.</p>" & _
        vbCrLf & "<p>" & tFrom & " Subject: " & " " & strSubject & "</p>  <p>Location: " & strLocation & _
        vbCrLf & "</p><p>Date and time: " & dtStart & "     Until: " & dtEnd & "</p>" & _
        vbCrLf & "<p>User: " & strUser & "</p>" & _
        vbCrLf & "<p>ID: " & strID & "</p>"

  txtemailAddress = Application.Session.CurrentUser.AddressEntry.GetExchangeUser.PrimarySmtpAddress
'  oAppointment.Companies = Application.Session.CurrentUser.AddressEntry.GetContact.CompanyName



  Set objMsg = Application.CreateItem(olMailItem)
' Next portion is used to generate HTML email
' adjust .To and .CC for your needs
  With objMsg
    .BodyFormat = olFormatHTML
    .To = ToEmail
    .CC = ccEmail
    .Subject = strSubject & " action: " & tFrom
    .HTMLBody = strHTMLHeader & vbCrLf & strBody
    .Attachments.Add strFilePath & ".ics", olByValue, , "ICS attachment"
    .Attachments.Add strFilePath & ".msg", olByValue, , "Outlook attachment"
' Hope this makes it to the email
    .Categories = "Internet"
'use .Display instead of Send if you want to add a note before sending
    .Display
    
  End With

' Uncomment for plain text email

'  objMsg.To = "felix.reta@alvaria.com"
'  objMsg.CC = "felix.reta@live.com"
'  objMsg.Subject = Item.Subject
  
'  objMsg.Body = strBody

'  objMsg.Attachments.Add strFilePath

'use Display instead of Send if you want to add a note before sending
'   objMsg.Send

   MsgBox ("email sent for " & Organisateur & " meeting request" & "Action: " & tFrom)
' Insert on my Google Calendar
' Previously, RTF/HTML (e.g. <p> statements would not be valid for v2 GCal

        strBody = strBody & "<p>" & tFrom & strRecurr & "Appointment " & strStatus & strCompany & sHostName & "" & "Organized by: " & Organisateur & " calendar." & _
        vbCrLf & tFrom & " Subject: " & strSubject & "  Location: " & strLocation & _
        vbCrLf & "Date and time: " & dtStart & "     Until: " & dtEnd & "" & _
        vbCrLf & "User: " & strUser & "" & _
        vbCrLf & "ID: " & strID & ""

strBody = "Original invite: " & strBody & vbCrLf & "Built information: " & strBody

' Time Zone Offset
TZOffset = "" ' to use computer's time zone, if it is "Z" is UTC 0

    If sendToGCal Then
        enviro = CStr(Environ("UserProfile"))
        elArchivo = enviro & "\Documents\" & "LasPWs.xml"

Debug.Print "environ: " & enviro
        
        LaCuenta = "Google"
        OID = ""
        OPW = ""
        strTypeReq = "Account"
        Call GetAcctInfoFromXML(strTypeReq, elArchivo, LaCuenta, OID, OPW)
        strTypeReq = "Frase"
        Call GetAcctInfoFromXML(strTypeReq, elArchivo, LaCuenta, ElID, LaPW)
' Decrypt
        For intX = 1 To Len(OPW) Step 2
            Mid(OPW, (intX + 1) / 2, 1) = _
                Chr(Val("&H" & Mid(OPW, intX, 2)))
        Next intX
        strUnCypher = RC4(Left(OPW, Len(OPW) / 2), ElID)
        Debug.Print "Uncyphered " & strUnCypher
'        Set cCrypto = New clsCryptoFilterBox ' To Crypto Object
'        cCrypto.Password = "Imagination is more important than knowledge-Albert Einstein"
'        cCrypto.InBuffer = OPW
'        cCrypto.Decrypt
'        strUnCypher = cCrypto.OutBuffer
'        Debug.Print strUnCypher

'        Call GoogleCalInsert(OID, strUnCypher, Organisateur, oAppointment.Subject, strBody, oAppointment.Location, TZOffset, "0", oAppointment.Start, oAppointment.End, Format(oAppointment.Start, "HH:MM:SS") & ".000", Format(oAppointment.End, "HH:MM:SS") & ".000")
'    Debug.Print "to GCal: " & OID & "-" & strUnCypher & "-" & Organisateur & "-" & oAppointment.Subject & "-" & strBody & "-" & oAppointment.Location & "-" & TZOffset & "-" & "0" & "-" & oAppointment.Start & "-" & oAppointment.End & "-" & Format(oAppointment.Start, "HH:MM:SS") & ".000" & "-" & Format(oAppointment.End, "HH:MM:SS") & ".000"
    
        strUnCypher = ""
    End If

End If

' Return original sensitivity (not sure this is necessary as the original appointment was not
' modified at all)
'   oAppointment.Sensitivity = iSensitivity
   
   Set objMsg = Nothing

'  If you'd like automatic response, activate following lines
'  Set oResponse = oAppointment.Respond(olMeetingAccepted)
'  oResponse.Send

'  oAppointment.Save
'  oMeetingoAppointment.Save

  Set oAppointment = Nothing
  Set oMeetingItem = Nothing

End Sub

Sub Pruebita()

Dim Str As String
Dim ProjColl As MatchCollection
Dim ProjClient As Match
Dim dteResult As Date
Dim dteTested As Date

dteTested = "4/17/2022"
' dteResult = WorksheetFunction.WorkDay(dteTested, 7)

dteResult = AddWeekDays(dteTested, 8)


Debug.Print dteResult


Str = "Roshan follow up [929219 87332 TEST]9898 965266 676111 7676 abcdef alicia-77732 55500 965674"
' Str = "Group Onboarding action: from Inbox"

Set ProjColl = findProjects(Str)

If ProjColl.Count > 0 Then

    Debug.Print strTime & "-" & ProjColl.Count & " Projects or clients found in Subject"

    For Each ProjClient In ProjColl
        Debug.Print ProjClient
    Next ProjClient

Else
    Debug.Print strtiime & "No project or client# present in Subject"
End If

End Sub

Function findProjects(strInput As String) As MatchCollection

Dim regexObject As RegExp
Dim Matches As MatchCollection
Dim Match As Match

Set regexObject = New RegExp


With regexObject
' Digit 5 included as 1st as NetSuite counter is currently at +500k
'    .Pattern = "\b\d{6}\b" 'Match 6 digits, between spaces or at the beginning and end of paragraph
    .Pattern = "\b50[5-6]\d{3}\b" ' Match exactly 50 then 5 to 6 then 3 digits between spaces or at the beginning and end of paragraph
    
    .Global = True 'use this to find all matches, not just the first match
End With

'Search string contains multiple versions of 'Hello'

'utilize the execute method and save the results in a new object that we call ‘matches’
Set Matches = regexObject.Execute(strInput)

'For Each Match In Matches
'  Debug.Print Match.value 'Result: all 5 or 6 digit numbers found in string
'Next Match

Set findProjects = Matches

End Function


Function GetFolderPath(ByVal FolderPath As String) As Outlook.Folder
' Not used for v0.0
    Dim oFolder As Outlook.Folder
    Dim FoldersArray As Variant
    Dim i As Integer
         
    On Error GoTo GetFolderPath_Error
    If Left(FolderPath, 2) = "\\" Then
        FolderPath = Right(FolderPath, Len(FolderPath) - 2)
    End If
    'Convert folderpath to array
    FoldersArray = Split(FolderPath, "\")
    Set oFolder = Application.Session.Folders.Item(FoldersArray(0))
    If Not oFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Outlook.Folders
            Set SubFolders = oFolder.Folders
            Set oFolder = SubFolders.Item(FoldersArray(i))
            If oFolder Is Nothing Then
                Set GetFolderPath = Nothing
            End If
        Next
    End If
    'Return the oFolder
    Set GetFolderPath = oFolder
    Exit Function
         
GetFolderPath_Error:
    Set GetFolderPath = Nothing
    Exit Function
End Function

Public Function GetGUID() As String
'Not used
GetGUID = Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36)
End Function
Function AddBusinessDays(StartDate As Date, numberOfDays As Integer) As Date

    Dim newDate As Date
    
    newDate = StartDate
' Possibly need to add Reference to MS Excel
    AddBusinessDays = WorksheetFunction.WorkDay(newDate, numberOfDays)

'Following didn't work when numdays >5

'    While numberOfDays > 0
'        newDate = newDate + 1
'
'        If Weekday(newDate, vbMonday) > 0 And Weekday(newDate) < 7 Then '2-6 is Mon-Fri then
'           numberOfDays = numberOfDays - 1
'        End If

'    Wend
    
'    AddBusinessDays = newDate


End Function

Function AddWeekDays(StartDate As Date, Days As Long) As Date
    Dim i As Long
    Dim d As Date
        
    d = StartDate
    i = 0
    
    While i < Days
        d = DateSerial(Year(d), Month(d), Day(d) + 1)
        If Weekday(d, vbMonday) < 6 Then
            i = i + 1
        End If
    Wend
 
    AddWeekDays = d
End Function
