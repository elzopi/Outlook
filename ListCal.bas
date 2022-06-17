Attribute VB_Name = "ListCal"
Public Enum eUnicodeConst
    LeftArrow
    RightArrow
    Clock2
    SailBoat
    PeaceLove
    CheckOK
    Ghost
    mapleleaf
    Calaca
    Corazon
    Finger1
    HandBro
    Airplane
    XMark
    Umbrella
    DogFace
End Enum


Sub ExportAppointmentsToExcel()
' Loosely based on: https://techniclee.wordpress.com/2013/06/21/exporting-appointments-from-outlook-to-excel/
' Need references to Outlook and VBscript regular expressions

    Const SCRIPT_NAME = "Export Appointments to Excel (Rev 1)"
    Const xlAscending = 1
    Const xlYes = 1
    Dim olkFld As Object, _
        olkLst As Object, _
        olkRes As Object, _
        olkApt As Object, _
        olkRec As Object, _
        excApp As Object, _
        excWkb As Object, _
        excWks As Object, _
        lngRow As Long, _
        lngCnt As Long, _
        strFil As String, _
        strLst As String, _
        strDat As String, _
        datBeg As Date, _
        datEnd As Date, _
        arrTmp As Variant

Dim OutlookApp As Outlook.Application
Dim OutlookNamespace As Namespace
Dim Folder As MAPIFolder
Dim ProjColl As MatchCollection
Dim ProjClient As Match

Set OutlookApp = New Outlook.Application
Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
Set olkFld = OutlookNamespace.Folders("Felix.Reta@Alvaria.com").Folders("Calendar")

Clear_Table_Contents "Calendar", "Table1"
'    Set olkFld = Application.ActiveExplorer.CurrentFolder
    If olkFld.DefaultItemType = olAppointmentItem Then
        strDat = InputBox("Enter the date range of the appointments to export in the form ""mm/dd/yyyy to mm/dd/yyyy""", SCRIPT_NAME, Date & " to " & Date)
        arrTmp = Split(strDat, "to")
        datBeg = IIf(IsDate(arrTmp(0)), arrTmp(0), Date) & " 12:00am"
        datEnd = IIf(IsDate(arrTmp(1)), arrTmp(1), Date) & " 11:59pm"
 '       strFil = InputBox("Enter a filename (including path) to save the exported appointments to.", SCRIPT_NAME)
 '       If strFil <> "" Then
 '           Set excApp = CreateObject("Excel.Application")
 '           Set excWkb = excApp.Workbooks.Add()
 '           Set excWks = excWkb.Worksheets(1)
            'Write Excel Column Headers
            
        Set excWks = Sheets("Calendar") 'arbitrarily set to existing tab/worksheet
        
            With excWks
                .Cells(1, 1) = "Project"
                .Cells(1, 2) = "Category"
                .Cells(1, 3) = "Subject"
                .Cells(1, 4) = "Starting Date"
                .Cells(1, 5) = "Ending Date"
                .Cells(1, 6) = "Start Time"
                .Cells(1, 7) = "End Time"
                .Cells(1, 8) = "Hours"
                .Cells(1, 9) = "Attendees"
            End With
            lngRow = 2
            Set olkLst = olkFld.Items
            olkLst.Sort "[Start]"
            olkLst.IncludeRecurrences = True
            Set olkRes = olkLst.Restrict("[Start] >= '" & Format(datBeg, "ddddd h:nn AMPM") & "' AND [Start] <= '" & Format(datEnd, "ddddd h:nn AMPM") & "'")
            'Write appointments to spreadsheet
            For Each olkApt In olkRes
                'Only export appointments
                If olkApt.Class = olAppointment Then
                    strLst = ""
                    For Each olkRec In olkApt.Recipients
                        strLst = strLst & olkRec.Name & ", "
                    Next
                    If strLst <> "" Then strLst = Left(strLst, Len(strLst) - 2)
                    'Add a row for each field in the message you want to export
                    excWks.Cells(lngRow, 2) = olkApt.Categories
                    excWks.Cells(lngRow, 3) = olkApt.Subject
        ' Check for project number in the Subject
                    strSubject = olkApt.Subject

                    Set ProjColl = findProjects(olkApt.Subject)

        ' Collection has at least 1 item, process accordingly
                    If ProjColl.Count > 0 Then
'                        Debug.Print strTime & "-" & ProjColl.Count & " Projects or clients found in Subject"
                        For Each ProjClient In ProjColl
                            excWks.Cells(lngRow, 1) = ProjClient
'                            Debug.Print ProjClient
                        Next ProjClient
                    Else
'                            Debug.Print "no PA found"
                            excWks.Cells(lngRow, 1) = "Other"
                    End If
                    
                    excWks.Cells(lngRow, 4) = Format(olkApt.Start, "mm/dd/yyyy")
                    excWks.Cells(lngRow, 5) = Format(olkApt.End, "mm/dd/yyyy")
                    excWks.Cells(lngRow, 6) = Format(olkApt.Start, "hh:nn ampm")
                    excWks.Cells(lngRow, 7) = Format(olkApt.End, "hh:nn ampm")
                    excWks.Cells(lngRow, 8) = DateDiff("n", olkApt.Start, olkApt.End) / 60
                    excWks.Cells(lngRow, 8).NumberFormat = "0.00"
                    excWks.Cells(lngRow, 9) = strLst
                    lngRow = lngRow + 1
                    lngCnt = lngCnt + 1
                End If
            Next
            excWks.Columns("A:H").AutoFit
            excWks.Range("A1:I" & lngRow - 1).Sort Key1:="Category", Order1:=xlAscending, Header:=xlYes
'            excWks.Cells(lngRow, 7) = "=sum(G2:G" & lngRow - 1 & ")"
'            excWkb.SaveAs strFil
'            excWkb.Close
            MsgBox "Process complete.  A total of " & lngCnt & " appointments were exported.", vbInformation + vbOKOnly, SCRIPT_NAME
        
        
    Else
        MsgBox "Operation cancelled.  The selected folder is not a calendar.  You must select a calendar for this macro to work.", vbCritical + vbOKOnly, SCRIPT_NAME
    End If
    
Set olkFld = OutlookNamespace.Folders("Felix.Reta@Alvaria.com").Folders("Tasks")
    
    Set excWks = Nothing
    Set excWkb = Nothing
    Set excApp = Nothing
    Set olkApt = Nothing
    Set olkLst = Nothing
    Set olkFld = Nothing
End Sub

Sub ExportTasksToExcel()
' Loosely based on: https://techniclee.wordpress.com/2013/06/21/exporting-appointments-from-outlook-to-excel/
' Need references to Outlook and VBscript regular expressions

    Const SCRIPT_NAME = "Export Tasks to Excel (Rev 1)"
    Const xlAscending = 1
    Const xlYes = 1
    Dim olkFld As Object, _
        olkLst As Object, _
        olkRes As Object, _
        olkApt As Object, _
        olkRec As Object, _
        olkTask As TaskItem, _
        excApp As Object, _
        excWkb As Object, _
        excWks As Object, _
        lngRow As Long, _
        lngCnt As Long, _
        strFil As String, _
        strLst As String, _
        strDat As String, _
        datBeg As Date, _
        datEnd As Date, _
        arrTmp As Variant

Dim OutlookApp As Outlook.Application
Dim OutlookNamespace As Namespace
Dim Folder As MAPIFolder
Dim ProjColl As MatchCollection
Dim ProjClient As Match
Dim n As Integer
Dim xTaskItems As Outlook.Items

Set OutlookApp = New Outlook.Application
Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
Set olkFld = OutlookNamespace.Folders("Felix.Reta@Alvaria.com").Folders("Tasks")

Clear_Table_Contents "Tasks", "Table3"
'    Set olkFld = Application.ActiveExplorer.CurrentFolder
    
    If olkFld.DefaultItemType = olTaskItem Then
        strDat = InputBox("Enter the date range of the appointments to export in the form ""mm/dd/yyyy to mm/dd/yyyy""", SCRIPT_NAME, Date & " to " & Date)
        arrTmp = Split(strDat, "to")
        datBeg = IIf(IsDate(arrTmp(0)), arrTmp(0), Date) & " 12:00am"
        datEnd = IIf(IsDate(arrTmp(1)), arrTmp(1), Date) & " 11:59pm"
 '       strFil = InputBox("Enter a filename (including path) to save the exported appointments to.", SCRIPT_NAME)
 '       If strFil <> "" Then
 '           Set excApp = CreateObject("Excel.Application")
 '           Set excWkb = excApp.Workbooks.Add()
 '           Set excWks = excWkb.Worksheets(1)
            'Write Excel Column Headers
            
        Set excWks = Sheets("Tasks") 'arbitrarily set to existing tab/worksheet
        
            With excWks
                .Cells(1, 1) = "Project"
                .Cells(1, 2) = "Category"
                .Cells(1, 3) = "Subject"
                .Cells(1, 4) = "Starting Date"
                .Cells(1, 5) = "Ending Date"
                .Cells(1, 6) = "Start Time"
                .Cells(1, 7) = "End Time"
                .Cells(1, 8) = "Hours"
                .Cells(1, 9) = "Attendees"
                
                .Cells(1, 10) = UniConst(CheckOK)
                
                ' & " " & UniConst(SailBoat) & " " & "Mili" & " " & UniConst(Calaca) & UniConst(DogFace)
            End With
            lngRow = 2
            Set olkLst = olkFld.Items
'            olkLst.Sort "[StartDate]"
'            olkLst.IncludeRecurrences = True
'            Set olkRes = olkLst.Restrict("[StartDate] >= '" & Format(datBeg, "ddddd h:nn AMPM") & "' AND [StartDate] <= '" & Format(datEnd, "ddddd h:nn AMPM") & "'")
'            Debug.Print olkLst.Count
            
            Set xTaskItems = olkFld.Items
            
'            For Each olkTask In xTaskItems
'                Debug.Print olkTask.Owner
'            Next olkTask
            
            'Write Tasks to spreadsheet
            For Each olkTask In xTaskItems
                'Only export appointments
                If olkTask.Class = olTask Then
                    If olkTask.Subject <> "" Then
' Capture recipients, originally for emails
'                    strLst = ""
'                    For Each olkRec In olkTask.Recipients
'                        strLst = strLst & olkRec.Name & ", "
'                    Next
'                    If strLst <> "" Then strLst = Left(strLst, Len(strLst) - 2)
                    'Add a row for each field in the message you want to export
                    excWks.Cells(lngRow, 2) = olkTask.Categories
                    excWks.Cells(lngRow, 3) = olkTask.Subject
        ' Check for project number in the Subject
                    strSubject = olkTask.Subject

                    Set ProjColl = findProjects(olkTask.Subject)

        ' Collection has at least 1 item, process accordingly
                    If ProjColl.Count > 0 Then
'                        Debug.Print strTime & "-" & ProjColl.Count & " Projects or clients found in Subject"
                        For Each ProjClient In ProjColl
                            excWks.Cells(lngRow, 1) = ProjClient
'                            Debug.Print ProjClient
                        Next ProjClient
                    Else
'                            Debug.Print "no PA found"
                            excWks.Cells(lngRow, 1) = "Other"
                    End If
                    
                    excWks.Cells(lngRow, 4) = Format(olkTask.StartDate, "mm/dd/yyyy")
                    excWks.Cells(lngRow, 5) = Format(olkTask.DateCompleted, "mm/dd/yyyy")
                    excWks.Cells(lngRow, 6) = Format(olkTask.StartDate, "hh:nn ampm")
                    excWks.Cells(lngRow, 7) = Format(olkTask.DateCompleted, "hh:nn ampm")
                    excWks.Cells(lngRow, 8) = DateDiff("n", olkTask.StartDate, olkTask.DateCompleted) / 60
                    excWks.Cells(lngRow, 8).NumberFormat = "0.00"
                    excWks.Cells(lngRow, 9) = olkTask.Owner
                    lngRow = lngRow + 1
                    lngCnt = lngCnt + 1
              End If
                End If
            Next
            
            excWks.Columns("A:H").AutoFit
            excWks.Range("A1:H" & lngRow - 1).Sort Key1:="Category", Order1:=xlAscending, Header:=xlYes
'            excWks.Cells(lngRow, 7) = "=sum(G2:G" & lngRow - 1 & ")"
'            excWkb.SaveAs strFil
'            excWkb.Close
            MsgBox "Process complete.  A total of " & lngCnt & " Tasks were exported.", vbInformation + vbOKOnly, SCRIPT_NAME
        
        
    Else
        MsgBox "Operation cancelled.  The selected folder is not a calendar.  You must select a calendar for this macro to work.", vbCritical + vbOKOnly, SCRIPT_NAME
    End If
    
Set olkFld = OutlookNamespace.Folders("Felix.Reta@Alvaria.com").Folders("Tasks")
    
    Set excWks = Nothing
    Set excWkb = Nothing
    Set excApp = Nothing
    Set olkTask = Nothing
    Set olkLst = Nothing
    Set olkFld = Nothing
End Sub



