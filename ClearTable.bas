Attribute VB_Name = "ClearTable"
'VBA Clear Table Content
Sub Clear_Table_Contents(wksName As String, tblName As String)
' Adapted From:     https://vbaf1.com/table/clear-table-content/

    'Definf Sheet and table name
     With Sheets(wksName).ListObjects(tblName) ' Calendar entries
        
        'Check If any data exists in the table
        If Not .DataBodyRange Is Nothing Then
            'Clear Content from the table
            .DataBodyRange.ClearContents
        End If
        
    End With
    
End Sub
