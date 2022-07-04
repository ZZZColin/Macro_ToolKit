Attribute VB_Name = "Compare_String_Module"
Function compare_string(str1, str2) As Double

    If str1 = "" Or str2 = "" Then
    
        similarity = 0
        
        GoTo final
        
    End If
    
    'Extract only letter and number --------------------------------------------------------------
    
    Set re = CreateObject("vbscript.regexp")
    
    With re
    
        .Pattern = "[A-Za-z0-9]+"
        .Global = True
        .MultiLine = True
        .ignorecase = True
        
    End With
    
    If re.test(str1) Then
    
        Set str1TempMatch = re.Execute(str1)
        
        For Each Match In str1TempMatch
            str1Temp = str1Temp & Match
        Next
        
    Else:
        str1Temp = str1
    End If
    
    If re.test(str2) Then
    
        Set str2TempMatch = re.Execute(str2)
        
        For Each Match In str2TempMatch
            str2Temp = str2Temp & Match
        Next
            
    Else:
        str2Temp = str2
    End If
    
    Debug.Print "After process: " & str1Temp & ", " & str2Temp
    
    '--------------------------------------------------------------------------------------------
    
    len1 = Len(str1Temp)
    
    len2 = Len(str2Temp)
    
    If len1 >= len2 Then
    
        loopStr = str1Temp
        sampleStr = str2Temp
        
    Else:
    
        loopStr = str2Temp
        sampleStr = str1Temp
        
    End If
    
    Dim result() As Integer
    ReDim result(1 To Len(loopStr) + Len(sampleStr) - 1)
    
    For i = 1 - Len(sampleStr) To Len(loopStr) - 1
        
        m = 0

        If i < 0 Then
            
            compareStr1 = Mid(loopStr, 1, Len(sampleStr) + i)
            compareStr2 = Mid(sampleStr, -i + 1, Len(sampleStr) + i)
            
        ElseIf i >= 0 And i < Len(sampleStr) Then
            
            compareStr1 = Mid(loopStr, i + 1, Len(sampleStr))
            compareStr2 = Mid(sampleStr, 1, Len(loopStr) - i)
            
        ElseIf i >= Len(sampleStr) Then
            
            compareStr1 = Mid(loopStr, i + 1, Len(loopStr) - i)
            compareStr2 = Mid(sampleStr, 1, Len(loopStr) - i)
            
        End If
        
        Debug.Print compareStr1, compareStr2
        
        For l = 1 To Len(compareStr1)
        
            If Mid(compareStr1, l, 1) = Mid(compareStr2, l, 1) Then
            
                m = m + 1
                
            End If
            
        Next l
        
        result(i + Len(sampleStr)) = m
        
    Next i
    
    resultMax = WorksheetFunction.Max(result)
    
    similarity = resultMax / Len(loopStr)
    
final:
    
    compare_string = similarity
    
End Function
