Attribute VB_Name = "Copy_File_Module"
Function copy_file(source_file, target_file, Optional OverwriteFiles)
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If IsMissing(OverwriteFiles) Then
    
        fso.copyfile sourcefile, destinationfolder
        
    Else:
    
        fso.copyfile sourcefile, destinationfolder, OverwriteFiles
        
    End If
    
End Function
