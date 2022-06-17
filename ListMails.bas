Attribute VB_Name = "ListMails"
Sub GetFromOutlook()
' Need to set named ranges on columns where the email info will appear
' Set reference to MS Outlook
' Loosely based on: https://www.howtoexcel.org/how-to-import-your-outlook-emails-into-excel-with-vba/
' Need references to Outlook

Dim OutlookApp As Outlook.Application
Dim OutlookNamespace As Namespace
Dim Folder As MAPIFolder
Dim OutlookMail As MailItem
Dim ProjColl As MatchCollection
Dim ProjClient As Match
Dim i As Integer
Dim co As Integer

Set OutlookApp = New Outlook.Application
Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")

' Set Folder = OutlookNamespace.Folders("Felix.Reta@Alvaria.com").Folders("Projects").Folders("505742 Consumer Cellular - Via WFM Training Lab")
' Set Folder = OutlookNamespace.Folders("Felix.Reta@Alvaria.com").Folders("Projects")
 Set Folder = OutlookNamespace.Folders("Felix.Reta@Alvaria.com").Folders("Inbox")

' Set Folder = OutlookNamespace.Folders("Felix.Reta@Alvaria.com").Folders("Projects")

i = 1

Clear_Table_Contents "eMails", "Table2"



co = Folder.Folders.Count

If Folder.Folders.Count > 0 Then

    DrillDownX Folder, i

Else
' Specific when no subfolders are present
    For Each Item In Folder.Items
' validate to process ONLY emails
        If TypeName(Item) = "MailItem" Then
           Set OutlookMail = Item
           
            If OutlookMail.ReceivedTime >= Range("From_date").Value Then
                
                    strSubject = OutlookMail.Subject

                    Set ProjColl = findProjects(OutlookMail.Subject)

        ' Collection has at least 1 item, process accordingly
                    If ProjColl.Count > 0 Then
'                        Debug.Print strTime & "-" & ProjColl.Count & " Projects or clients found in Subject"
                        For Each ProjClient In ProjColl
                            Range("email_Folder").Offset(i, 0).Value = ProjClient
'                            Debug.Print ProjClient
                        Next ProjClient
'                    Else
'                            Debug.Print "no PA found"
'                        Range("email_Folder").Offset(i, 0).Value = "Other"
                    End If
                
                
                
                
                
                Range("eMail_subject").Offset(i, 0).Value = OutlookMail.Subject
                Range("eMail_date").Offset(i, 0).Value = OutlookMail.ReceivedTime
                Range("eMail_sender").Offset(i, 0).Value = OutlookMail.SenderName
                Range("eMail_text").Offset(i, 0).Value = Left(OutlookMail.Body, 20)
                Range("Source").Offset(i, 0).Value = "Inbox email"
        
                i = i + 1
            End If
        End If
        
      Next Item


End If

' GetAllEmailsInFolder Folder

'For Each outlookMail In Folder.Items
'    If outlookMail.ReceivedTime >= Range("From_date").Value Then
'
'       Range("email_Folder").Offset(i, 0).Value = Folder
'
'        Range("eMail_subject").Offset(i, 0).Value = outlookMail.Subject
'        Range("eMail_date").Offset(i, 0).Value = outlookMail.ReceivedTime
'        Range("eMail_sender").Offset(i, 0).Value = outlookMail.SenderName
'        Range("eMail_text").Offset(i, 0).Value = outlookMail.Body
'
'        i = i + 1
'    End If
'Next outlookMail

' Set Folder = OutlookNamespace.Folders("Felix.Reta@Alvaria.com").Folders("Inbox")

' bIsInbox = True

 Set Folder = OutlookNamespace.Folders("Felix.Reta@Alvaria.com").Folders("Projects")
 i = i + 1
 DrillDownX Folder, i

' Insert formula to extract project number
Range("Project").Offset(2, 0).Formula = "=LEFT([@Folder],6)"

Set Folder = Nothing
Set OutlookNamespace = Nothing
Set OutlookApp = Nothing

End Sub

Private Function DrillDownX(parentFolder As Outlook.MAPIFolder, startRow As Integer)
' Function to drilldown the folder structure and output to named ranges,from date in named range

    Dim f As MAPIFolder
    Dim fToProc As MAPIFolder
    Dim fileNum As Integer
    Dim OutlookMail As MailItem
    Dim i As Integer
    Dim co As Integer
    
    i = 1

'   co = parentFolder.Folders.Count ' 0 is no subfolders

'   Debug.Print co
   For Each f In parentFolder.Folders
        
    For Each Item In f.Items
' validate to process ONLY emails
        If TypeName(Item) = "MailItem" Then
           Set OutlookMail = Item
           
            If OutlookMail.ReceivedTime >= Range("From_date").Value Then
                Range("email_Folder").Offset(i, 0).Value = f
                Range("eMail_subject").Offset(i, 0).Value = OutlookMail.Subject
                Range("eMail_date").Offset(i, 0).Value = OutlookMail.ReceivedTime
                Range("eMail_sender").Offset(i, 0).Value = OutlookMail.SenderName
                Range("eMail_text").Offset(i, 0).Value = Left(OutlookMail.Body, 20)
                Range("Source").Offset(i, 0).Value = "Folder email"
        
                i = i + 1
            End If
        End If
        
      Next Item
  
    DrillDownX f, i

    Next f

End Function
Private Function DrillDown(parentFolder As Outlook.MAPIFolder)
' Function to drilldown the folder structure and prints to file
    
    Dim f As Outlook.MAPIFolder
    Dim fileNum As Integer
    Dim OutlookMail As MailItem
    
    i = 1
    For Each f In parentFolder.Folders
'        DrillDown f
'        Print #1, f.FolderPath, f.Items.Count
        Debug.Print f.FolderPath, "#" & f.Items.Count
    For Each OutlookMail In f.Items
        
        If OutlookMail.ReceivedTime >= Range("From_date").Value Then
            Range("email_Folder").Offset(i, 0).Value = f
            Range("eMail_subject").Offset(i, 0).Value = OutlookMail.Subject
            Range("eMail_date").Offset(i, 0).Value = OutlookMail.ReceivedTime
            Range("eMail_sender").Offset(i, 0).Value = OutlookMail.SenderName
            Range("eMail_text").Offset(i, 0).Value = Left(OutlookMail.Body, 20)
        
            i = i + 1
        End If
    
    Next OutlookMail

    DrillDown f
        
    Next
End Function

Sub LoopFoldersInInbox()

    Dim ns                  As Object
    Dim objFolder           As Object
    Dim objSubfolder        As Object
    Dim lngCounter          As Long

    Set OutlookApp = New Outlook.Application
    Set ns = OutlookApp.GetNamespace("MAPI")

    
'    Set objFolder = ns.GetDefaultFolder(olFolderInbox) ' 6 also
    Set objFolder = ns.Folders("Felix.Reta@alvaria.com").Folders("Projects")
    
    DrillDownX objFolder
    
    For Each objSubfolder In objFolder.Folders
        With ActiveSheet
            lngCounter = lngCounter + 1
            .Cells(lngCounter, 1) = objSubfolder.Name
            .Cells(lngCounter, 2) = objSubfolder.Items.Count
        End With

'        Debug.Print objSubfolder.Name
'        Debug.Print objSubfolder.Items.Count

    Next objSubfolder

End Sub


Sub OLFoldersDrillDown()
' Prints folder structure and item counts to a file
' Needs reference to MS Outlook Object Library
   
    Dim OLApp As Outlook.Application
    Dim olNs As Outlook.Namespace
    Dim olParentFolder As Outlook.MAPIFolder
    Dim olFolderA As Outlook.MAPIFolder
    Dim olFolderB As Outlook.MAPIFolder
    Dim olFolderC As Outlook.MAPIFolder
    Dim fileNum As Integer
   
    Set OLApp = New Outlook.Application
    Set olNs = OLApp.GetNamespace("MAPI")
    fileNum = FreeFile()

   Open "C:\Users\freta\AppData\Local\Temp\Output.txt" For Output As fileNum


    Set olParentFolder = olNs.Folders("Felix.Reta@alvaria.com").Folders("Projects")
    DrillDown olParentFolder
    Print #fileNum, olParentFolder.FolderPath, olParentFolder.Items.Count
    Close #fileNum
    
End Sub

Private Sub GetAllEmailsInFolder(CurrentFolder As Outlook.Folder)
' Loosely based on: http://www.gregthatcher.com/Scripts/VBA/Outlook/GetListOfOutlookEmailsInCurrentFolder.aspx?AspxAutoDetectCookieSupport=1

    Dim currentItem As MailItem
    
    Report = Report & "Folder Name: " & CurrentFolder.Name & " (Store: " & CurrentFolder.Store.DisplayName & ")" & vbCrLf
    

    For Each currentItem In CurrentFolder.Items
        
        Report = Report & currentItem.CreationTime
        Report = Report & currentItem.Subject
        Report = Report & vbCrLf
    '    Report = Report & CurrentItem.Body
    '    Report = Report & vbCrLf
        Report = Report & "----------------------------------------------------------------------------------------"
        Report = Report & vbCrLf
        Debug.Print Report
        
    Next
    
End Sub



