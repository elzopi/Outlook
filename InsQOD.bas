Attribute VB_Name = "InsQOD"
' Attribute VB_Name = "InsQoD"
' Insert Quote of the Day from RSS feed from brainyquote @ feedburner.com, designed and developed by Felix Reta
' felix.reta@live.com
' Copyright 2005 --> 2021
' Tested on Microsoft Outlook 2003, 2007, 2010, 2016
' References to Microsoft Word and Microsoft XML are necessary [Tools --> References]
' Original CommandBar code from Sue Mosher OL MVP circa 2007
' Modified by Felix Reta to add stationery functionality Jan'2012
' Modified for potential HTML change @ brainyquote website/RSS feed May'2013
' Included code to insert quote after string Jan'2018
' Updated Reply to insert quote


' References to Microsoft Word and Microsoft XML are necessary [Tools --> References]
Sub InsertQOD()

    Dim objMsg As Outlook.MailItem
        
    Set objMsg = Application.CreateItem(olMailItem)
    
    a = objMsg.ConversationIndex
    
    objMsg.Display
    Call AddQOD(objMsg)
'    Call DeleteSig(objMsg)
    Set objMsg = Nothing
End Sub
Sub AddQOD(msg As Outlook.MailItem)
    Dim objDoc As Word.Document
    Dim objSel As Word.Selection
    Dim objBkm As Word.Bookmark
    On Error Resume Next
    
    Set objDoc = msg.GetInspector.WordEditor
    Set objSel = objDoc.Application.Selection
    With objSel
'         .MoveEnd
         .EndKey wdStory, wdMove
         .Font.Name = "Bradley Hand ITC"
         .Font.Bold = True
         .Font.Italic = False
         .Font.Color = wdColorPlum
         .Font.Size = 12
         .InsertAfter TheDailyQuote
         .HomeKey wdStory, wdMove

    End With

End Sub

Function TheDailyQuote() As String
        Dim My_URL As String
        Dim My_Obj As Object
        Dim xmlHttp As New XMLHTTP60 'was XMLHTTP30, changed to 60 on Office 2016
        
        Dim My_Var As String
        '        Dim s As String
        Dim My_Quote As String
        Dim StrQuoteAuthor As String
        Dim IntQuoteStarts As Integer
        Dim IntQuoteEnds As Integer
        Dim QuotesCount As Integer
        Dim AAuthorQuote(10, 2) As String
        Dim iNumQuotes As Integer
        Dim nAleatorio As Integer

'        My_URL = "http://feeds.feedburner.com/brainyquote/QUOTEBR"
        ' Code modified to obtain specifically the 1st. quote on an RSS feed at: http://feeds.feedburner.com/brainyquote/QUOTEBR
        ' http://feeds.feedburner.com/brainyquote/QUOTEFU   for funny quotes
        ' http://feeds.feedburner.com/brainyquote/QUOTEAR   for art quotes
        ' http://feeds.feedburner.com/brainyquote/QUOTENA   for nature quotes
        
        Dim QuotesFeeds(4) ' to hold 4 quote feed options May'2021
        QuotesFeeds(1) = "http://feeds.feedburner.com/brainyquote/QUOTEBR"
        QuotesFeeds(2) = "http://feeds.feedburner.com/brainyquote/QUOTEFU"
        QuotesFeeds(3) = "http://feeds.feedburner.com/brainyquote/QUOTEAR"
        QuotesFeeds(4) = "http://feeds.feedburner.com/brainyquote/QUOTENA"
        
        nAleatorio = Int((4 * Rnd) + 1) ' Genera numero aleatorio entr 1 y 4
        My_URL = QuotesFeeds(nAleatorio)
'        Debug.Print My_URL
        
'        My_Obj = CreateObject("MSXML2.XMLHTTP")
        xmlHttp.Open "GET", My_URL, False
'        My_Obj.Send
        xmlHttp.Send
'        My_Var = My_Obj.responsetext
        My_Var = xmlHttp.responseText
                ' Get Author, should the first <item>, right after a <title> tag

        IntQuoteStarts = InStr(1, My_Var, "<item>") ' find the tag of the 1st quote to get the show on the road

        My_Var = Mid(My_Var, IntQuoteStarts, Len(My_Var) - IntQuoteStarts)


        ' This section used to be the main code to obtain the 1st quote, replaced by the iterative function GetQuotes
        'right after this tag is the author name, just before /title tag
                IntQuoteStarts = InStr(IntQuoteStarts, My_Var, "<title>") + 15
                IntQuoteEnds = InStr(IntQuoteStarts, My_Var, "</title>") - IntQuoteStarts - 1
                StrQuoteAuthor = Mid(My_Var, IntQuoteStarts + 7, IntQuoteEnds - 6)

        ' Get the 1st daily quote, should the first string, right after a <description>" tag [Notice triple quotes]
                IntQuoteStarts = InStr(1, My_Var, "<description>""") + 13
                IntQuoteEnds = InStr(IntQuoteStarts, My_Var, "</description>") - IntQuoteStarts
                My_Quote = Mid(My_Var, IntQuoteStarts, IntQuoteEnds)
    TheDailyQuote = My_Quote & " - " & StrQuoteAuthor
    
End Function

Sub InsertQoDHere()

' Dim sText As String
  Dim sFile As String
  Dim objShape As Object
  Dim strDisclaimer As String
  Dim FName As String
'  Dim strFilename As String: strFilename = Environ("UserProfile") & "\Documents\ML-Disclaimer.txt"
  Dim strFileContent As String
  Dim iFile As Integer: iFile = FreeFile

' sText = TheDailyQuote

On Error GoTo ErrHandler

If TypeName(ActiveWindow) = "Inspector" Then
    If ActiveInspector.IsWordMail And ActiveInspector.EditorType = olEditorWord Then
'        ActiveInspector.WordEditor.Application.Selection.TypeText sText
       ActiveInspector.WordEditor.Application.Selection.TypeText vbCrLf

'    Set x = ActiveInspector.WordEditor.Application.Selection
'    x.Font.Name = "calibri"
    
' Read disclaimer text into variable
'Open strFilename For Input As #iFile
'strDisclaimer = Input(LOF(iFile), iFile)
'Close #iFile
    
'    FName = Environ("UserProfile") & "\Pictures\Capture-ML-letterhead.png"
    
' Insert Signature
' Set objShape = objSel.InlineShapes.AddPicture(FName, False, True)

       With ActiveInspector.WordEditor.Application.Selection
         .Font.Name = "Banff-Normal"
         .Font.Bold = True
         .Font.Italic = False
         .Font.Color = wdColorBlue
         .Font.Size = 24
         .TypeText "Felix Reta "
         .Font.Name = "Calibri"
         .Font.Size = 12
         .Font.Subscript = True
         .TypeText "PMP " & Chr(169) & vbCrLf
         .Font.Subscript = False
'         .InlineShapes.AddPicture FName
         .TypeText vbCrLf
         .Font.Name = "Calibri"
         .Font.Size = 12
         .Font.Italic = True
         .Font.Color = wdColorGreen
         .TypeText "(954) 779-6179" & vbCrLf

'         .TypeText "MagicLeap" & vbCrLf
'         .HomeKey wdStory, wdMove

       End With
      
      With ActiveInspector.WordEditor.Application.Selection
         .Font.Name = "Calibri"
         .Font.Bold = False
         .Font.Italic = True
         .Font.Color = wdColorBlack
         .Font.Size = 14
         .TypeText "Technology Vendor Specialist" & vbCrLf
'         .HomeKey wdStory, wdMove

       End With

' Insert QOD captured in sText
       With ActiveInspector.WordEditor.Application.Selection
         .Font.Name = "Bradley Hand ITC"
         .Font.Bold = True
         .Font.Italic = False
         .Font.Color = wdColorPlum
         .Font.Size = 12
         .TypeText TheDailyQuote
         .TypeText vbCrLf

         
'         .HomeKey wdStory, wdMove
         
'         .Font.Bold = False
' Adding disclaimer

         .Font.Name = "Helvetica Neue"
         .Font.Bold = False
         .Font.Color = wdColorBlueGray
         .Font.Size = 8
         .TypeText strDisclaimer
         .HomeKey wdStory, wdMove
         
       End With
    
    End If
End If

Exit Sub

ErrHandler:
Beep

End Sub

Sub ReplyMSG()

' Updated Jun 2021 to insert and then move to top

    Dim olItem As Outlook.MailItem
    Dim olReply As MailItem ' Reply

    
    For Each olItem In Application.ActiveExplorer.Selection
        Set olReply = olItem.ReplyAll
'            olReply.HTMLBody = "Hello, Thank you. " & vbCrLf & olReply.HTMLBody
        olReply.Display
        
    With ActiveInspector.WordEditor.Application.Selection
         .Font.Name = "Bradley Hand ITC"
         .Font.Bold = True
         .Font.Italic = False
         .Font.Color = wdColorPlum
         .Font.Size = 12
         .TypeText vbCrLf
         .TypeText vbCrLf
         .TypeText TheDailyQuote
         .TypeText vbCrLf
         .HomeKey wdStory

' Return font values to previous

         .Font.Name = "Arial"
         .Font.Bold = False
         .Font.Color = wdColorBlue
         
    End With
        
'        With olReply
'         .Font.Name = "Bradley Hand ITC"
'         .Font.Bold = True
'         .Font.Italic = False
'         .Font.Color = wdColorPlum
'         .Font.Size = 12
'         .InsertAfter TheDailyQuote
'        End With
        
'        Call AddQOD(olReply)

        'olReply.Send
    Next olItem
    Set objMsg = Nothing
End Sub

Sub testing()

     Debug.Print Environ("UserProfile") & "\Pictures"
    

End Sub
Private Sub cmdFileDialog_Click()
  
   ' Requires reference to Microsoft Office 11.0 Object Library.
 
   Dim fDialog As Office.FileDialog
   Dim varFile As Variant
 
   ' Clear listbox contents.
'   Me.FileList.RowSource = ""
 
   ' Set up the File Dialog.
   Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
 
   With fDialog
 
      ' Allow user to make multiple selections in dialog box
      .AllowMultiSelect = True
             
      ' Set the title of the dialog box.
      .Title = "Please select one or more files"
 
      ' Clear out the current filters, and add our own.
      .Filters.Clear
      .Filters.Add "Access Databases", "*.MDB"
      .Filters.Add "Access Projects", "*.ADP"
      .Filters.Add "All Files", "*.*"
 
      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      If .Show = True Then
 
         'Loop through each file selected and add it to our list box.
         For Each varFile In .SelectedItems
            Me.FileList.AddItem varFile
         Next
 
      Else
         MsgBox "You clicked Cancel in the file dialog box."
      End If
   End With
End Sub
Sub c()
   ' Requires reference to Microsoft Office 11.0 Object Library.
   Dim fDialog As Office.FileDialog
   Dim varFile As Variant
   ' Clear listbox contents.
   Me.FileList.RowSource = ""
   ' Set up the File Dialog.
   Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
   With fDialog
      ' Allow user to make multiple selections in dialog box
      .AllowMultiSelect = True
      ' Set the title of the dialog box.
      .Title = "Please select one or more files"
      ' Clear out the current filters, and add our own.
      .Filters.Clear
      .Filters.Add "Access Databases", "*.ACCDB"
      .Filters.Add "Access Projects", "*.ADP"
      .Filters.Add "All Files", "*.*"
      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      If .Show = True Then
         'Loop through each file selected and add it to our list box.
         For Each varFile In .SelectedItems
            Me.FileList.AddItem varFile
         Next
      Else
         MsgBox "You clicked Cancel in the file dialog box."
      End If
   End With
End Sub