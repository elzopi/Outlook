Attribute VB_Name = "Module4"
' Attribute VB_Name = "LJournal"
Sub logToJournal()
Dim filesys
  Dim myOlApp As Object
  Dim myJItem As JournalItem
  Dim logFile As String
  Dim WholeLine As String
' Watch out for this array, if you have a large event log
' containing many log-in log-out times (e.g. 9009 & 9006 record types)
  Dim strlogtimes(400) As String
  Dim sortedLog() As String
  Dim intLowBound As Integer
  Dim intUpBound As Integer
  Dim dteToCheck As String
  Dim sameDate As Boolean
  Set filesys = CreateObject("Scripting.FileSystemObject")
  Set myOlApp = CreateObject("Outlook.Application")
  Set myJItem = myOlApp.CreateItem(olJournalItem)


' ====================================================
' Access the System log file
' Process only 6009 (start log) and 6006 (stop log) records
' to figure logout and login times for journal
' ====================================================
Category = "N/A"
Computer_Name = "N/A"
event_code = "N/A"
Message = "N/A"
Record_Number = "N/A"
Source_name = "N/A"
time_written = "N/A"
Event_Type = "N/A"
User = "N/A"
Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
compname = WshShell.ExpandEnvironmentStrings("%computername%")
strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colLoggedEvents = objWMIService.ExecQuery("Select * from Win32_NTLogEvent Where Logfile='System'", , 48)
ignored = 0

Open "c:\Temp\log" & Left(DT(Now), 8) & ".txt" For Output As 1
i = LBound(strlogtimes())
For Each objEvent In colLoggedEvents
    Category = objEvent.Category
    Computer_Name = objEvent.ComputerName
    event_code = objEvent.EventCode
    Message = objEvent.Message
    Record_Number = objEvent.RecordNumber
    Source_name = objEvent.SourceName
    time_written = objEvent.TimeWritten
    Event_Type = objEvent.Type
    User = objEvent.User
    time_written = Left(time_written, (InStr(time_written, ".") - 1))
'  Message section
    If (IsNull(Message)) Then
        Message = "NA"
    Else
        Message = Replace(Message, Chr(13), " ")
        Message = Replace(Message, Chr(10), "")
        Message = Replace(Message, ",", "")
        Message = Replace(Message, Chr(34), "`")
        Message = Replace(Message, "'", "")
        Message = Mid(Message, 1, (Len(Message) - 2))
        Message = Left(Message, 254)
    End If

'    line_to_insert = "'" & time_written & "','" & Category & "','" & event_code & "','" & Event_Type & "','" & Source_name & "','" & Message & "'"
    line_to_insert = "'" & time_written & "','" & event_code & "'"
    Select Case event_code
      Case "6009"            ' log started = login
        Print #1, line_to_insert
'        MsgBox line_to_insert
        strlogtimes(i) = time_written & "|" & event_code
'        MsgBox strLogTimes(i)
        i = i + 1
      Case "6006"            ' log stopped = logout
        Print #1, line_to_insert
'        MsgBox line_to_insert
        strlogtimes(i) = time_written & "|" & event_code
'        MsgBox strLogTimes(i)
       i = i + 1
      Case Else
        ignored = ignored + 1
    End Select

Next
strlogtimes(i) = "29991231235959|9999"
Close #1
maxindice = i
MsgBox "Done, " & "Records Processed: " & i & vbCrLf & "Records ignored: " & ignored

' Let's take a look at the un-sorted bunch
'For i = LBound(strLogTimes()) To maxindice
'    MsgBox strLogTimes(i)
'Next

sortArray strlogtimes, (maxindice)

' MsgBox sortArray(sortedLog)

' Let's take a look at the sorted bunch
' For i = LBound(strlogtimes()) To maxindice
'    MsgBox strlogtimes(i)
' Next

' Build journal records, the process assumes log records are sorted

intLowBound = LBound(strlogtimes())
intUpBound = maxindice
stopflag = False
startflag = False

' if 1st entry is a logout, ignore
If Right(strlogtimes(intLowBound), 4) = "6006" Then
   intLowBound = intLowBound + 1
End If

i = intLowBound
dteToCheck = Left(strlogtimes(intLowBound), 8)
StartDate = dteToCheck
StartTime = Mid(strlogtimes(i), 9, 6)

For i = intLowBound To intUpBound
    If dteToCheck = Left(strlogtimes(i + 1), 8) Then
       sameDate = True
    Else
       stoptime = Mid(strlogtimes(i), 9, 6)
       sameDate = False
    End If

      If Not sameDate Then
        myJItem.Subject = "J@" + Format(Now, "mm-dd-yyyy HH:MM:SS")
        myJItem.Start = CDate(cnvDate(StartDate & StartTime))
        myJItem.End = CDate(cnvDate(StartDate & stoptime))
        hrs = Format((myJItem.End - myJItem.Start), "HH:MM")
        myJItem.Categories = "Time entry"
        myJItem.Body = myJItem.Body & "Logged from: " & StartDate & " - " & StartTime & " to: " & stoptime & " ===> " & hrs & " Hours" & vbCrLf
        myJItem.Type = "Posted time"
        myJItem.Companies = "Siemens"
        myJItem.Save
        StartDate = ""
        StartTime = ""
        startflag = False
        stopdate = ""
        stoptime = ""
        stopflag = False
        sameDate = False
        dteToCheck = Left(strlogtimes(i + 1), 8)
        StartDate = dteToCheck
        StartTime = Mid(strlogtimes(i + 1), 9, 6)
    End If

' Got to 9999, finish it up!
If Right(strlogtimes(i + 1), 4) = "9999" Then
   i = intUpBound
End If

Next

End Sub

Function DT(dDate As Date) As String
  Dim num As String
  num = Year(dDate)
  If Len(num) < 4 Then num = String(4 - Len(num), "0") & num
  DT = DT & num
  num = Month(dDate)
  If Len(num) < 2 Then num = String(2 - Len(num), "0") & num
  DT = DT & num
  num = Day(dDate)
  If Len(num) < 2 Then num = String(2 - Len(num), "0") & num
  DT = DT & num
  DT = DT & "T"
  num = Hour(dDate)
  If Len(num) < 2 Then num = String(2 - Len(num), "0") & num
  DT = DT & num
  num = Minute(dDate)
  If Len(num) < 2 Then num = String(2 - Len(num), "0") & num
  DT = DT & num
  num = Second(dDate)
  If Len(num) < 2 Then num = String(2 - Len(num), "0") & num
  DT = DT & num
End Function
' Bubble sort
Function sortArray(ByRef strlogtimes() As String, mxi As Integer)
Dim i As Integer
i = LBound(strlogtimes())
j = mxi
For i = i To mxi - 1

    For j = 0 To mxi - 1
    
    If strlogtimes(j) > strlogtimes(j + 1) Then
    
        Temp = strlogtimes(j)
        strlogtimes(j) = strlogtimes(j + 1)
        strlogtimes(j + 1) = Temp
        
    End If
    
    Next

Next
End Function

Function cnvDate(strDate)

mm = Mid(strDate, 5, 2)
dd = Mid(strDate, 7, 2)
yy = Left(strDate, 4)
hh = Mid(strDate, 9, 2)
mi = Mid(strDate, 11, 2)
cnvDate = mm & "/" & dd & "/" & yy & " " & hh & ":" & mi
a = 1
End Function
