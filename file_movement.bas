Attribute VB_Name = "file_movement"
'Enable Microsoft Scripting Runtime
Option Explicit
Private fso As Scripting.FileSystemObject
Private newFolderPath As String
Sub Create_Folder_Copy_Files()
    
    Dim oldFolderPath As String
    
    newFolderPath = Environ("UserProfile") & "\vba\New_Folder"
    oldFolderPath = Environ("UserProfile") & "\vba\VBA Files"
    
    Set fso = New Scripting.FileSystemObject
    
    If fso.FolderExists(oldFolderPath) Then
        
        If Not fso.FolderExists(newFolderPath) Then
            fso.CreateFolder (newFolderPath)
        End If
                
        CopyExcelFiles oldFolderPath
        
    End If
    
    Set fso = Nothing
    
End Sub

Sub CopyExcelFiles(startFolderPath As String)

    Dim fil As Scripting.File
    Dim subfold As Scripting.Folder
    Dim oldFolder As Scripting.Folder
    
    Set oldFolder = fso.GetFolder(startFolderPath)
    
    For Each fil In oldFolder.Files
    
        If Left(fso.GetExtensionName(fil.path), 3) = "xls" Then
             fil.Copy newFolderPath & "\" & fil.Name
        End If
    
    Next fil

    For Each subfold In oldFolder.SubFolders
        Call CopyExcelFiles(subfold.path)
    Next subfold

End Sub




