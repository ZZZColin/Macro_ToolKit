Attribute VB_Name = "Get_Column_Dic_Module"
Function get_column_dic(data, Optional row_index)

    Set column_dic = CreateObject("scripting.dictionary")
    
    If IsMissing(row_index) Then
    
        For i = 1 To UBound(data, 2)
        
            column_dic(data(1, i)) = i
            
        Next i
        
    Else:
    
        For i = 1 To UBound(data, 2)
        
            column_dic(data(row_index, i)) = i
            
        Next i
        
    End If
    
    Set get_column_dic = column_dic

End Function
