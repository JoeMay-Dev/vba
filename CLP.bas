Attribute VB_Name = "Module1"
Option Explicit
Public tRowsCRQ As Long
Public tRowsAFQ As Long
Public vName As String
Public vNum As String
Public pStart As String
Public pEnd As String


Sub CLPExhibit()
            
    SortQuery
    
    VarSelect
    
    VendSelect
    
    Summary
    
    NewFolder
    
    ExhibitCopy


End Sub

Sub SortQuery()
     
    Sheets("Customer Rebate Query").Activate
    tRowsCRQ = Cells(1, 1).CurrentRegion.Rows.Count
    Columns("A:V").Select
    Sheets("Customer Rebate Query").Sort.SortFields.Clear
    Sheets("Customer Rebate Query").Sort.SortFields.Add2 Key:= _
        Range("C2:C" & tRowsCRQ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    Sheets("Customer Rebate Query").Sort.SortFields.Add2 Key:= _
        Range("A2:A" & tRowsCRQ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Customer Rebate Query").Sort
        .SetRange Range("A1:V" & tRowsCRQ)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Sheets("Admin Fee Query").Activate
    tRowsAFQ = Cells(1, 1).CurrentRegion.Rows.Count
    Columns("A:V").Select
    Sheets("Admin Fee Query").Sort.SortFields.Clear
    Sheets("Admin Fee Query").Sort.SortFields.Add2 Key:= _
        Range("C2:C" & tRowsAFQ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    Sheets("Admin Fee Query").Sort.SortFields.Add2 Key:= _
        Range("A2:A" & tRowsAFQ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Admin Fee Query").Sort
        .SetRange Range("A1:V" & tRowsAFQ)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub
Sub VarSelect()

    Sheets("Customer Rebate Query").Activate
    vName = Cells(2, 3)
    vNum = Cells(2, 2)
    pStart = Mid(Cells(2, 1), 8)
    pEnd = Mid(Cells(tRowsCRQ, 1), 8)

End Sub

Sub VendSelect()

    Dim cCount As Double
    Dim aCount As Double
        
        cCount = 2
       
            Do While Cells(cCount, 3) = vName
                Worksheets("Customer Rebate Query").Range("A2:V" & cCount).Select
                cCount = cCount + 1
            Loop
    
        cCount = cCount - 1
                  
    Selection.Cut
    Worksheets("CUSTOMER REBATE").Activate
    Cells(2, 1).Select
    ActiveSheet.Paste
    Worksheets("Customer Rebate Query").Range("A2:V" & cCount).Delete (xlUp)
    Sheets("Admin Fee Query").Activate
    
        aCount = 2
        
            Do While Cells(aCount, 3) = vName
                Worksheets("Admin Fee Query").Range("A2:V" & aCount).Select
                aCount = aCount + 1
            Loop

        aCount = aCount - 1

  Selection.Cut
  Worksheets("ADMIN FEE").Activate
  Cells(2, 1).Select
  ActiveSheet.Paste
  Worksheets("Admin Fee Query").Range("A2:V" & aCount).Delete (xlUp)

End Sub

Sub Summary()

    Worksheets("SUMMARY").Activate
    Cells(1, 1) = vName
    Cells(2, 1) = "CO VEN #" & vNum
    Cells(4, 2) = "P" & pStart & " -P" & pEnd
    
End Sub

Sub NewFolder()

'    If Len(Dir("G:\National SIS\Shared\Leads\Customer Loyalty Program\Test P" & pStart & " -P" & pEnd, vbDirectory)) = 0 Then
'        MkDir "G:\National SIS\Shared\Leads\Customer Loyalty Program\Test P" & pStart & " -P" & pEnd
'    End If
    
    If Len(Dir("C:\Users\resil\Desktop\Test P" & pStart & " -P" & pEnd, vbDirectory)) = 0 Then
        MkDir "C:\Users\resil\Desktop\Test P" & pStart & " -P" & pEnd
    End If
    
End Sub

Sub ExhibitCopy()

Dim wb As Workbook
Dim sFileName As String
Dim sFilePath As String

    With Sheets(Array("SUMMARY", "CUSTOMER REBATE", "ADMIN FEE"))
            .Select
            .Copy
    End With
    
    sFileName = vName & " " & vNum & " CLP P" & pStart & " - P" & pEnd & " 2018"
'    sFilePath = "G:\National SIS\Shared\Leads\Customer Loyalty Program\Test P" & pStart & " -P" & pEnd & "\"
    sFilePath = "C:\Users\resil\Desktop\Test P" & pStart & " -P" & pEnd & "\"
    Set wb = ActiveWorkbook
    wb.SaveAs Filename:=sFilePath & sFileName, FileFormat:=51
    ActiveWorkbook.Close
    
End Sub
Sub VendorCountComp()
Dim crqRng As Range, afqRange As Range
Dim rCountC As Long, rCountA As Long, crqVC As Long, afqVC As Long

tRowsCRQ = Sheets("Customer Rebate Query").Cells(1, 1).CurrentRegion.Rows.Count
tRowsAFQ = Sheets("Admin Fee Query").Cells(1, 1).CurrentRegion.Rows.Count

Worksheets("Customer Rebate Query").Range("B2:B" & tRowsCRQ).Select

crqRng = [B2:B & tRowsCRQ]






End Sub

Function NumUniqueValues(Rng As Range) As Long
Dim myCell As Range
Dim UniqueVals As New Collection
Application.Volatile
On Error Resume Next
For Each myCell In Rng
    UniqueVals.Add myCell.Value, CStr(myCell.Value)
Next myCell
On Error GoTo 0
NumUniqueValues = UniqueVals.Count
End Function
