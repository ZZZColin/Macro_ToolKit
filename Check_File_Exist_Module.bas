Attribute VB_Name = "Check_File_Exist_Module"
Function check_file_exist(file_path)

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.fileexists(file_path) Then
    
        check_file_exist = True
        
    Else:
    
        check_file_exist = False
        
    End If
    
End Function
