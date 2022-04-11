VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColinFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Data, ColumnDic


Sub INI(FilePath, Optional SheetName, Optional TargetRow)

    If IsMissing(SheetName) Then
    
        Data = GetData(FilePath)
        
    Else:
    
        Data = GetData(FilePath, SheetName)
        
    End If
    
    If IsMissing(TargetRow) Then
    
        Set ColumnDic = GetColumnDic(Data)
        
    Else:
    
        Set ColumnDic = GetColumnDic(Data, TargetRow)
        
    End If

End Sub

Private Function GetData(FilePath, Optional SheetName)
    
    CurrentFile = Split(FilePath, "\")(UBound(Split(FilePath, "\")))
    
    Debug.Print "Processing " & CurrentFile & "..."
    
    Set eapp = CreateObject("excel.application")
    
    Set efile = eapp.Workbooks.Open(FilePath)
    
    If IsMissing(SheetName) Then
    
        Data = efile.Sheets(1).UsedRange.Value
    
    Else:
    
        Data = efile.Sheets(SheetName).UsedRange.Value
        
    End If
    
    eapp.Application.DisplayAlerts = False
    efile.Close
    Set efile = Nothing
    eapp.Application.DisplayAlerts = True
    
    eapp.Quit
    Set eapp = Nothing
    
    Debug.Print "Date Row Count: " & UBound(Data, 1) & ", Column Count: " & UBound(Data, 2)
    
    GetData = Data
    
End Function

Private Function GetColumnDic(Data, Optional TargetRow)
    
    Set Dic = CreateObject("scripting.dictionary")
    
    If IsMissing(TargetRow) Then
        
        r = 1
        
    Else:
    
        r = TargetRow
        
    End If
        
    For i = 1 To UBound(Data, 2)
    
        If Dic.exists(Data(r, i)) Then
        
            Dic(Data(r, i)).Add i
            
        Else:
        
            Set Dic(Data(r, i)) = New Collection
            
            Dic(Data(r, i)).Add i
            
        End If
        
    Next i
    
    For Each k In Dic.keys

        For Each c In Dic(k)
        
            Debug.Print "Column Name: " & k & " at: " & c
            
        Next
        
    Next
    
    Set GetColumnDic = Dic

End Function
