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

Public WithEvents curCal As Outlook.Items
Attribute curCal.VB_VarHelpID = -1
' Added to catch when mail is flagged up as Follow Up
Public WithEvents objReminders As Outlook.Reminders
Attribute objReminders.VB_VarHelpID = -1
Public WithEvents colInspectors As Outlook.Inspectors
Attribute colInspectors.VB_VarHelpID = -1
Public gcolMyInspectors As Collection
Public Sub Initialize_handler()

    Set myOlCalItems = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderCalendar).Items

End Sub


Dim inProgress As Boolean
  
Private Sub Application_Startup()
   Dim NS As Outlook.NameSpace
   Set NS = Application.GetNamespace("MAPI")
   Set curCal = NS.GetDefaultFolder(olFolderCalendar).Items
   Set NS = Nothing
   inProgress = False
   Set objReminders = Application.Reminders
   
' Added to catch when mail is flagged up as Follow Up
   Set gcolMyInspectors = New Collection
   Set colInspectors = Application.Inspectors

End Sub

Private Sub Application_Quit()
 Set gcolMyInspectors = Nothing
 Set colInspectors = Nothing
End Sub

Private Sub objReminders_ReminderFire(ByVal ReminderObject As Reminder)
'    Call Send_Email_Using_VBA
'    MsgBox ("Reminder fired up!" & ReminderObject.Caption)
End Sub
Private Sub colInspectors_NewInspector(ByVal Inspector As Inspector)

 ' This will be called everytime we open
 ' a new Inspector, so check if this is
 ' one that we want to monitor
 
 Dim MyInspectorHandler As clsInspectorHandler

 es = Inspector.currentItem.Class

'    MsgBox ("At NewInspector Class: " & Es)

' If Inspector.currentItem.Class = olMail Then
'     If Not Inspector.currentItem.Sent Then
'         ' This is an unsent email so we want to
'         ' trap the buttons that can be clicked
'         Set MyInspectorHandler = New clsInspectorHandler
'         Call MyInspectorHandler.SetInspector(Inspector)
'         gcolMyInspectors.Add MyInspectorHandler
'     End If
' End If

End Sub
Private Sub myOlCalItems_ItemChange(ByVal Item As Object)

Dim prompt As String
Dim calItem As Outlook.MeetingItem
 
 Set calItem = Item
 
 Debug.Print "At myOlCalItems change sub, status: " & calItem.Status & " meeting status: " & calItem.MeetingStatus
 calItem.Display
 
End Sub
Private Sub curCal_ItemAdd(ByVal Item As Object)
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
  Dim cCrypto As clsCryptoFilterBox ' Need Class Module clsCryptoFilterBox
  Dim ToEmail As String
  Dim ccEmail As String
  Dim strRecurr As String
  Dim Organisateur As String

' By debugging, found out that for an attached .ics invite, the sequence for oAppointment.ResponseStatus
' is 5 (olResponseNotResponded) when opening the attachment
' and then 3 (olResponseAccepted) when clicking the Accept button
' appt_Write was not called under this condition (.ics opened)
' Feb 2014
  Set olNS = Application.Session
  Set oExUser = olNS.CurrentUser.AddressEntry.GetExchangeUser
  Email = oExUser.PrimarySmtpAddress
   
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
  strHTMLcode = enviro & "\My Documents\" & "HTML-OL-header.txt"
  
   strID = Item.EntryID
   strUser = Environ$("username")
   strUser = strUser & "-" & ThisOutlookSession.Session.CurrentUser.Name

' Build Company variables OID will contain the To email address and OPW the Company name
   elArchivo = "LasPWs.xml"
   LaCuenta = "ALV"
   OID = ""
   OPW = ""
   strTypeReq = "Company"
   Call GetAcctInfoFromXML(strTypeReq, elArchivo, LaCuenta, OID, OPW)
   ToEmail = OID
' Modify for cc'd email if any
   ccEmail = "felix.reta@live.com"
' Get the Company name if available, or default it

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

    
    Set oAppointment = Item

    iStatus = oAppointment.ResponseStatus
' Meeting Status is updated whenever a meeting request is opened, empirically:
' 1 is present as received, if it is responded (e.g. ResponseStatus set) it is set accordingly.
' 5 is set when the meeting is opened (e.g. written to the calendar) and the organizer cancelled it
' 3 yet to be observed

' 0 & 0 (both flags) are set as soon as I start setting an appointment, and both set
' local appointments (e.g. penciled in)
' 1 & 1 (both flags) are set as soon as I save in my calendar without sending it

    mStatus = oAppointment.MeetingStatus
Debug.Print strUser & "-" & sHostName & "-Item Entry ID: " & Right(strID, 12) & " Meeting Status at ItemAdd is: " & oAppointment.MeetingStatus

    lProcess = True
    strStatus = "Meeting "
Select Case oAppointment.MeetingStatus
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
    Case 3
        strStatus = strStatus & "-" & "Accepted"
    Case 4
        strStatus = strStatus & "-" & "Declined"
' This is set right after an invite shows up in the Inbox....
    Case 5
        strStatus = strStatus & "-" & "Not responded"
        
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

' Test for appointment in the past, don't process.
If oAppointment.Start < Now Then
    lProcess = False
End If

If oAppointment.RecurrenceState = olApptNotRecurring Then
    Debug.Print "Non recurring appointment"
    strRecurr = ""
Else
    Debug.Print "Recurring appointment identified"
    strRecurr = "Recurring "
End If

'   set oAppointment = Item.GetAssociatedAppointment(True)

Debug.Print strUser & "-" & sHostName & "-Item Entry ID: " & Right(strID, 12) & " Meeting Response Status at ItemAdd is: " & strStatus

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
        vbCrLf & "<p>Subject: " & oAppointment.Subject & "</p>  <p>Location: " & oAppointment.Location & _
        vbCrLf & "</p><p>Date and time: " & oAppointment.Start & "     Until: " & oAppointment.End & "</p>" & _
        vbCrLf & "<p>User: " & strUser & "</p>" & _
        vbCrLf & "<p>ID: " & strID & "</p>"

' For v2.2 hope this makes it as part of the .ics
  oAppointment.Categories = "Internet"

' Adjust field to your company name (I haven't seen this field making it to
' the calendar .msg type
  oAppointment.Companies = strCompany
  txtemailAddress = Application.Session.CurrentUser.AddressEntry.GetExchangeUser.PrimarySmtpAddress
'  oAppointment.Companies = Application.Session.CurrentUser.AddressEntry.GetContact.CompanyName


' An opportunity to grab the ThisOutlookSession.Session.CurrentUser.Name
' if so desired or .Address for X.400 address, not thoroughly tested,
' .Name works in Outlook 2010
'  oAppointment.Organizer = ThisOutlookSession.Session.CurrentUser.Name
  
  oAppointment.SaveAs strFilePath & ".ics", olICal 'or olVcal flag (vcal has issues with recurring appts.)
  oAppointment.SaveAs strFilePath & ".msg"

  Set objMsg = Application.CreateItem(olMailItem)
' Next portion is used to generate HTML email
' adjust .To and .CC for your needs
  With objMsg
    .BodyFormat = olFormatHTML
    .To = ToEmail
    .CC = ccEmail
    .Subject = Item.Subject
    .HTMLBody = strHTMLHeader & vbCrLf & strBody
    .Attachments.Add strFilePath & ".ics", olByValue, , "ICS attachment"
    .Attachments.Add strFilePath & ".msg", olByValue, , "Outlook attachment"
' Hope this makes it to the email
    .Categories = "Internet"
'use .Display instead of Send if you want to add a note before sending
    .Send
    
  End With

' Uncomment for plain text email

'  objMsg.To = "felix.reta@alvaria.com"
'  objMsg.CC = "felix.reta@live.com"
'  objMsg.Subject = Item.Subject
  
'  objMsg.Body = strBody

'  objMsg.Attachments.Add strFilePath

'use Display instead of Send if you want to add a note before sending
'   objMsg.Send
' Organizer appears blank when copying an appointment to the local calendar
If oAppointment.Organizer = "" Then
   Organisateur = strUser
Else
   Organisateur = oAppointment.Organizer
End If

   MsgBox ("email sent for " & Organisateur & " meeting request")
' Insert on my Google Calendar
' Previously, RTF/HTML (e.g. <p> statements would not be valid for v2 GCal

        strBody = strBody = "<p>" & strRecurr & "Appointment " & strStatus & strCompany & sHostName & "" & "Organized by: " & Organisateur & " calendar." & _
        vbCrLf & "Subject: " & oAppointment.Subject & "  Location: " & oAppointment.Location & _
        vbCrLf & "Date and time: " & oAppointment.Start & "     Until: " & oAppointment.End & "" & _
        vbCrLf & "User: " & strUser & "" & _
        vbCrLf & "ID: " & strID & ""

strBody = "Original invite: " & oAppointment.Body & vbCrLf & "Built information: " & strBody

' Time Zone Offset
TZOffset = "" ' to use computer's time zone, if it is "Z" is UTC 0

    If sendToGCal Then
        elArchivo = "LasPWs.xml"
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
        Debug.Print strUnCypher
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

Private Sub curCal_ItemChange(ByVal Item As Object)
' Testing to apply for cancelled and updated appointments
' the routine is called when deleting a calendar item, need to check when pressing "remove from calendar"
' Need to initialize ToEmail & ccEmail
' still testing to catch changed/cancelled appointment
' Name                          Value   Description
'olMeeting                      1       The meeting has been scheduled.
'olMeetingCanceled              5       The scheduled meeting has been cancelled.
'olMeetingReceived              3       The meeting request has been received.
'olMeetingReceivedAndCanceled   7       The scheduled meeting has been cancelled but still appears on the user's calendar.
'olNonMeeting                   0       An Appointment item without attendees has been scheduled. This status can be used to set up holidays on a calendar.

  Dim strID As String
  Dim olNS As Outlook.NameSpace
  Dim oMeetingItem As Outlook.MeetingItem
  Dim oResponse As Outlook.MeetingItem
  Dim oAppointment As Outlook.AppointmentItem
  Dim objMsg As MailItem
  Dim enviro As String
  Dim strFilePath As String
  Dim iSensitivity As Integer
  Dim elStatus As Integer
  Dim strBody As String
  Dim lProcess As Boolean
  Dim OID As String
  Dim OPW As String
  Dim oExUser As Outlook.ExchangeUser
  Dim strReason As String
  
' Debug.Print "inProgress switch: " & inProgress

Set olNS = Application.Session 'Sets the stage
Set oExUser = olNS.CurrentUser.AddressEntry.GetExchangeUser
Email = oExUser.PrimarySmtpAddress
  
  If inProgress Then
        Exit Sub
  Else
        inProgress = False
  End If
  
  enviro = CStr(Environ("TEMP"))
  strFilePath = enviro & "\Outlook Appointment"

' Build Company variables OID will contain the To email address and OPW the Company name
   elArchivo = "LasPWs.xml"
   LaCuenta = "ALV"
   OID = ""
   OPW = ""
   strTypeReq = "Company"
   Call GetAcctInfoFromXML(strTypeReq, elArchivo, LaCuenta, OID, OPW)
   ToEmail = OID
' Modify for cc'd email if any
   ccEmail = "felix.reta@live.com"
' Get the Company name if available, or default it
strCompany = oExUser.CompanyName

If strCompany = "" Then
   strCompany = OPW
End If

lProcess = False

' Next statement bombs
'  strID = Item.EntryID


' On Error Resume Next
' Set newCalFolder = GetFolderPath("\\Internet Calendars")
    
    If Item.Class = olMeeting Then
        Set oMeetingItem = Item
    Else
        Set oAppointment = Item
        strID = oAppointment.GlobalAppointmentID

    End If

Debug.Print "Item Entry ID: " & strID & " Appointment Status at ItemChange Sub is: " & oAppointment.MeetingStatus
    
'   set oAppointment = Item.GetAssociatedAppointment(True)
' Catch dismissed calendar entries after reminder is shown, apparently you get a meeting status of 3
        elStatus = oAppointment.MeetingStatus
        Select Case oAppointment.MeetingStatus
            Case olMeetingReceived
                lProcess = False
            Case olMeetingCanceled Or olMeetingReceivedAndCanceled
                strReason = "CANCELLED: "
                strBody = "Appointment cancelled at " & strCompany & " " & "Organized by: " & oAppointment.Organizer & " calendar." & _
                vbCrLf & "Subject: " & oAppointment.Subject & "     Location: " & oAppointment.Location & _
                vbCrLf & "Date and time: " & oAppointment.Start & "     Until: " & oAppointment.End
                lProcess = True
            Case olnomeeting
                lProcess = False
            Case Else
                strReason = "UPDATED: "
                strBody = "An existing appointment was changed/updated at " & strCompany & " " & "Organized by: " & oAppointment.Organizer & " calendar." & _
                vbCrLf & "Subject: " & oAppointment.Subject & "     Location: " & oAppointment.Location & _
                vbCrLf & "Date and time: " & oAppointment.Start & "     Until: " & oAppointment.End
                lProcess = True
        End Select

' OLD code
'        If oAppointment.MeetingStatus = olMeetingReceived Then
'           lProcess = False
'        Else
'          If oAppointment.MeetingStatus = olMeetingCanceled Or oAppointment.MeetingStatus = olMeetingReceivedAndCanceled Then
'            elStatus = oAppointment.MeetingStatus
'            oAppointment.Subject = "CANCELLED: " & oAppointment.Subject
'            strBody = "Appointment cancelled at " & strCompany & " " & "Organized by: " & oAppointment.Organizer & " calendar." & _
'            vbCrLf & "Subject: " & oAppointment.Subject & "     Location: " & oAppointment.Location & _
            vbCrLf & "Date and time: " & oAppointment.Start & "     Until: " & oAppointment.End
'            lProcess = True
'          Else
'            oAppointment.Subject = "UPDATED: " & oAppointment.Subject
'            iSensitivity = oAppointment.Sensitivity
'            oAppointment.Sensitivity = olPersonal
'            strBody = "An existing appointment was changed/updated at " & strCompany & " " & "Organized by: " & oAppointment.Organizer & " calendar." & _
'            vbCrLf & "Subject: " & oAppointment.Subject & "     Location: " & oAppointment.Location & _
'            vbCrLf & "Date and time: " & oAppointment.Start & "     Until: " & oAppointment.End
'            lProcess = True
'          End If
'        End If


' A recurrent appointmant fails here

  If lProcess Then
'      oAppointment.Subject = strReason & oAppointment.Subject & " status: " & elStatus
      oAppointment.SaveAs strFilePath & ".ics", olICal 'or olVcal flag (vcal has issues with recurring appts.)
      oAppointment.SaveAs strFilePath & ".msg"

      Set objMsg = Application.CreateItem(olMailItem)
      With objMsg
        .To = ToEmail
        .CC = ccEmail
        .Subject = strReason & oAppointment.Subject & " status: " & elStatus
        .Body = strBody
        .Attachments.Add strFilePath & ".ics", olByValue, , "ICS attachment"
        .Attachments.Add strFilePath & ".msg", olByValue, , "Outlook attachment"
' Hope this makes it to the email
        .Categories = "Internet"
'use .Display instead of Send if you want to add a note before sending
        .Send
      End With

' watch out
'      objMsg.Attachments.Add strFilePath

'use Display instead of Send if you want to add a note before sending
    
' Return original sensitivity (not sure this is necessary as the original appointment was not
' modified at all)
'   oAppointment.Sensitivity = iSensitivity
End If

Set objMsg = Nothing

'  If you'd like automatic response, activate following lines
'  Set oResponse = oAppointment.Respond(olMeetingAccepted)
'  oResponse.Send

'  oAppointment.Save
'  oMeetingoAppointment.Save


  Set oAppointment = Nothing
  Set oMeetingItem = Nothing

  
End Sub
