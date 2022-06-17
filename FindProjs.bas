Attribute VB_Name = "FindProjs"
Public Function findProjects(strInput As String) As MatchCollection
' Need references to Outlook and VBscript regular expressions
' Felix Reta felix.reta@gmail.com (c)2022

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

