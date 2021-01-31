Attribute VB_Name = "OpenChangeClose"
Option Explicit
Sub Open_Write_Close_Save()

Dim path As String
Dim i As Integer

For i = 1 To 6

    path = Environ("UserProfile") & "\github\test" & i & ".xlsx"
    
    Workbooks.Open path
    
    Range("A" & i).Value = "This workbook is named test" & i & ".xlsx !"
    
    'auto-saves workbook and eliminates dialogue box
    
    ActiveWorkbook.Close savechanges:=True
    
Next i


End Sub

Sub Open_Delete_Close_Save()

Dim path As String
Dim i As Integer

For i = 1 To 6

    path = Environ("UserProfile") & "\github\test" & i & ".xlsx"
    
    Workbooks.Open path
    
    Cells.ClearContents
                    
    'auto-saves workbook and eliminates dialogue box
    
    ActiveWorkbook.Close savechanges:=True
    
Next i

End Sub





