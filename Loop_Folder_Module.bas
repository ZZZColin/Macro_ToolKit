Attribute VB_Name = "Loop_Folder_Module"
Function loop_folder(file_extension)

    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then
            folder_path = .SelectedItems(1)
        Else:
            Exit Function
        End If
    End With
    
    MyFile = Dir(folder_path & "\*." & file_extension)

    Do While MyFile <> ""
        
        
        
        MyFile = Dir
        If MyFile = "" Then
            Exit Do
        End If

    Loop
    
End Function
