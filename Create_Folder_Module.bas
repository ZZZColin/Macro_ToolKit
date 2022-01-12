Attribute VB_Name = "Create_Folder_Module"
Function create_folder(folder_path)

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fso.CreateFolder folder_path

End Function
