Attribute VB_Name = "Module1"
'Enable Microsoft Scripting Runtime
Option Explicit
Sub UsingTheScriptingRuntimeLibrary()

    Dim fso As Scripting.FileSystemObject
    Dim fil As Scripting.File
    Dim oldFolder As Scripting.Folder
    Dim newFolderPath As String
    Dim oldFolderPath As String
    
    
    newFolderPath = Environ("UserProfile") & "\vba\New_Folder"
    oldFolderPath = Environ("UserProfile") & "\vba\VBA Files\Test"
    
    Set fso = New Scripting.FileSystemObject
    
    If fso.FolderExists(oldFolderPath) Then
        Set oldFolder = fso.GetFolder(oldFolderPath)
        
        If Not fso.FolderExists(newFolderPath) Then
            fso.CreateFolder (newFolderPath)
        End If
        
        For Each fil In oldFolder.Files
            
            If Left(fso.GetExtensionName(fil.Path), 3) = "xls" Then
                 fil.Copy newFolderPath & "\" & fil.Name
            End If
        
        Next fil
        
    End If
     
    
    
'    If fso.FileExists(oldFolderPath & "\move_this.xlsx") Then
'
'        Set fil = fso.GetFile(oldFolderPath & "\move_this.xlsx")
'
'        If fil.Size > 6000 Then
'            fil.Copy newFolderPath & "\" & fil.Name
'        End If
'
'        fso.CopyFile _
'            Source:=oldFolderPath & "\move_this.xlsx" _
'            , Destination:=newFolderPath & "\move_this_copy.xlsx"
'    End If
    
    Set fso = Nothing
         
    
End Sub




