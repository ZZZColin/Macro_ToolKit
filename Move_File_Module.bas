Attribute VB_Name = "Move_File_Module"
Function move_file(source_file, target_file)

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fso.movefile source_file, target_file
    
End Function
