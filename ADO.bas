Attribute VB_Name = "Module1"
Option Explicit

Sub ADO()
    'add ADO references: Microsoft ActiveX Data Objects Objects 6.1 Library

    Dim fName As String, adoConn As New ADODB.Connection, strSQL As String, recSet As Recordset
    
    fName = ThisWorkbook.Path & Application.PathSeparator & "Book1ADO.xslx"
    
    'get connection
    adoConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & fName & ";" & _
        "Extended Properties=""Excel 12.0;HDR=Yes;"";"
    
    strSQL = "Select [First Name], Sum(Amount) From [Sale$] Group by [First Name]"
    
    recSet.Open strSQL, adoConn
    
    shResult.Cells.ClearContents
    shResult.Range("A1").CopyFromRecordset recSet
    
    'close connection
    adoConn.Close
    
    
    
End Sub
