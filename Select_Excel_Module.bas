Attribute VB_Name = "Select_Excel_Module"
Function select_excel(display_message)

    MsgBox ("Please select " & display_message)
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls;*.xlsx;*.xlsm;*.csv"
        .Filters.Add "All Files", "*.*"
        If .Show = -1 Then
            select_excel = .SelectedItems(1)
        Else:
            select_excel = ""
        End If
    End With
    
    Debug.Print "File Path: " & select_excel
    
End Function
