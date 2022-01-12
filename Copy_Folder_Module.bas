Attribute VB_Name = "Copy_Folder_Module"
Function copy_folder(source_folder, target_folder, Optional OverwriteFiles)
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If IsMissing(OverwriteFiles) Then
    
        fso.copyfolder source_folder, target_folder
        
    Else:
    
        fso.copyfolder source_folder, target_folder, OverwriteFiles
        
    End If
    
End Function
