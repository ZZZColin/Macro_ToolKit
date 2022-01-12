Attribute VB_Name = "Move_Folder_Module"
Function move_folder(source_folder, target_folder)

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fso.movefolder source_folder, target_folder
    
End Function
