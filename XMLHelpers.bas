Attribute VB_Name = "XMLHelpers"
Sub PruebaXML()
' Parsing XML, from
' http://vba2vsto.blogspot.com/2008/12/reading-xml-from-vba.html

 Dim oDoc As DOMDocument
 Dim fSuccess As Boolean
 Dim oRoot As IXMLDOMNode
 Dim oAccount As IXMLDOMNode
 Dim oAttributes As IXMLDOMNamedNodeMap
 Dim oAccountName As IXMLDOMNode
 Dim oChildren As IXMLDOMNodeList
 Dim oChild As IXMLDOMNode
 Dim intI As Integer
 Dim enviro As String
 Dim strXMLFileLoc As String
 Dim strXX As String
 Dim sw As Boolean

  On Error GoTo HandleErr
  Set oDoc = New DOMDocument
         ' Load the  XML from disk, without validating it. Wait
        ' for the load to finish before proceeding.
        oDoc.async = False
        oDoc.validateOnParse = False
  enviro = CStr(Environ("UserProfile"))
  strXMLFileLoc = enviro & "\My Documents\In Transit\My Programs\" & "GoogleCal.xml"

        fSuccess = oDoc.Load( _
          strXMLFileLoc)
         ' If anything went wrong, quit now.
        If Not fSuccess Then
          GoTo ExitHere
        End If
       
        oDoc.LoadXML (strXMLFileLoc)
        Set oRoot = oDoc.SelectNodes("Accounts/Account")
        For Each oChild In oRoot
            Debug.Print oChild.Text
        Next
        

ExitHere:
        Exit Sub
HandleErr:
        MsgBox "Error " & Err.Number & ": " & Err.Description
        Resume ExitHere
        Resume
      
End Sub

Public Sub GetAcctInfoFromXML(strElement, XMLFile, strAcctName, strAcctID, strPW As String)

Dim XDoc As DOMDocument
Dim xDetails As IXMLDOMNode
Dim xData As IXMLDOMNode
Dim xChild As IXMLDOMNode
Dim xAcct As IXMLDOMNode
Dim xName As IXMLDOMNode
Dim xUserID As IXMLDOMNode
Dim xPw As IXMLDOMNode
Dim oRoot As IXMLDOMNode
Dim strXMLFileLoc As String
Dim strTempData As String
Dim strXX As String
Dim lesChild As IXMLDOMNodeList
Dim numElements As Integer
Dim i As Integer

Set XDoc = New DOMDocument
XDoc.async = False
XDoc.validateOnParse = False

enviro = CStr(Environ("UserProfile"))
' strXMLFileLoc = enviro & "\My Documents\In Transit\My Programs\" & "LasPWs.xml"
strXMLFileLoc = enviro & "\My Documents\" & XMLFile

' strTempData = strXMLFileLoc
' strXX = ReadXML(strTempData)


XDoc.Load (strXMLFileLoc)
If LoadError(XDoc) Then Exit Sub

' The following code displays all XML file contents
'If XDoc.HasChildNodes Then
'        Debug.Print "Number of child Nodes: " & XDoc.ChildNodes.Length
'        For Each XMLNode In XDoc.ChildNodes
'            Debug.Print "Node name:" & XMLNode.nodeName
'            Debug.Print "Type:" & XMLNode.nodeTypeString & "(" & XMLNode.NodeType & ")"
'            Debug.Print "Text: " & XMLNode.Text
'        Next XMLNode
'End If

' First matching node
'Set xUserID = XDoc.SelectSingleNode("//Password")
'Debug.Print xUserID.Text

' Select element with modified XML
' Found at http://stackoverflow.com/questions/16538329/how-i-can-read-all-attributes-from-a-xml-with-vba
strXX = ""
For Each xAcct In XDoc.SelectNodes("//Accounts/" & strElement)
    If xAcct.Text = strAcctName Then 'get the hashed Pw from XML file
'    If xAcct.Text = "Google" Then 'get the hashed Pw debug
       strPW = xAcct.Attributes.getNamedItem("Password").Text
       strAcctID = xAcct.Attributes.getNamedItem("UserID").Text
       Debug.Print strPW
       Debug.Print strAcctID
    End If
Next

End Sub

Sub XMLmethod()
' Select all nodes with label
Set lesChild = XDoc.SelectNodes("//UserID")

If Not (lesChild Is Nothing) Then
    For Each xUserID In lesChild
        Debug.Print xUserID.Text
    Next xUserID
End If

Set lesChild = XDoc.getElementsByTagName("UserID")
For Each xUserID In lesChild
    For Each xPw In xUserID.ChildNodes
        Debug.Print xUserID.nodeName & "=" & xUserID.Text
    Next xPw
Next xUserID


Set xDetails = XDoc.DocumentElement
Set xData = xDetails.FirstChild

For Each xData In xDetails.ChildNodes
    For Each xChild In xData.ChildNodes
        Debug.Print xChild.BaseName & " " & xChild.Text
    Next xChild
Next xData

' Set lesChild = XDoc.ChildNodes
' Call RecurseChildNodes(XDoc, lesChild)



End Sub
Public Function RecurseChildNodes(xmlDoc As MSXML2.DOMDocument30, childNode As IXMLDOMNodeList)
   
   Dim CurrChildNode  As IXMLDOMNodeList
   Dim intNodeCounter As Integer
   Dim elTexto As String
   
   Set CurrChildNode = childNode
   
   For intNodeCounter = 0 To CurrChildNode.Length - 1
      
      If CurrChildNode.Length > 0 Then
         Set childNode = CurrChildNode.Item(intNodeCounter).ChildNodes
         If childNode.Length > 0 Then
            RecurseChildNodes xmlDoc, childNode
            elTexto = elTexto & CurrChildNode.Item(intNodeCounter).nodeName & "= " & CurrChildNode.Item(intNodeCounter).nodeTypedValue & vbCrLf
         End If
      End If
      
        
   Next intNodeCounter
   
End Function
Function GetNode(parentNode As Object, nodeNumber As Long) As Object
 
 On Error Resume Next
 ' if parentNode is a MSXML2.IXMLDOMNodeList
 Set GetNode = parentNode.Item(nodeNumber - 1)

 ' if parentNode is a MSXML2.IXMLDOMNode
 If GetNode Is Nothing Then
    Set GetNode = parentNode.ChildNodes(nodeNumber - 1)
 End If

End Function

Function LoadError(xmlDoc As Object) As Boolean

' checks if a xml file load error occurred
LoadError = (xmlDoc.parseError.ErrorCode <> 0)


End Function
Function ReadXML(fileName As String) As String()
' see http://www.jpsoftwaretech.com/read-xml-files-using-dom/

Dim xmlDoc As Object  ' MSXML2.DOMDocument60
Dim myvalues As Object  ' MSXML2.IXMLDOMNode
Dim values As Object  ' MSXML2.IXMLDOMNode
Dim value As Object  ' MSXML2.IXMLDOMNode
Dim tempString() As String
Dim numRows As Long, numColumns As Long
Dim i As Long, j As Long

' check if file exists

If Len(Dir(fileName)) = 0 Then Exit Function

' create MSXML 6.0 document and load existing file
Set xmlDoc = GetDomDoc
If xmlDoc Is Nothing Then Exit Function

xmlDoc.Load fileName
If LoadError(xmlDoc) Then Exit Function

' second node starts the node tree
Set myvalues = GetNode(xmlDoc, 2)
' array size? add +1 for header row
numColumns = myvalues.ChildNodes.Length
numRows = GetNode(myvalues, 1).ChildNodes.Length + 1
ReDim tempString(1 To numColumns + 1, 1 To numRows + 1)
For i = 1 To numColumns
Set values = GetNode(myvalues, i)     ' first value in every column is node name
tempString(i, 1) = values.nodeName
For j = 1 To numRows - 1
tempString(i, j + 1) = GetNode(values, j).nodeTypedValue
Next j
Next i
ReadXML = tempString

End Function

Function GetDomDoc() As Object

' MSXML2.DOMDocument
On Error Resume Next
Set GetDomDoc = CreateObject("MSXML2.DOMDocument.6.0")

End Function

