Attribute VB_Name = "Get_Data_Module"
Function get_data(file_path, Optional sheet_name)
    
    Set eapp = CreateObject("excel.application")
    
    Set efile = eapp.Workbooks.Open(file_path)
    
    If IsMissing(sheet_name) Then
    
        Data = efile.Sheets(1).UsedRange.Value
        
    Else:
    
        Data = efile.Sheets(sheet_name).UsedRange.Value
        
    End If
    
    eapp.Application.DisplayAlerts = False
    efile.Close
    Set efile = Nothing
    eapp.Application.DisplayAlerts = True
    
    eapp.Quit
    Set eapp = Nothing
    
    get_data = Data

End Function
