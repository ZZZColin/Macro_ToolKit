Attribute VB_Name = "Check_Folder_Exist_Module"
Function check_folder_exist(folder_path)

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(folder_path) Then
    
        check_folder_exist = True
        
    Else:
    
        check_folder_exist = False
        
    End If

End Function
