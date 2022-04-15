Attribute VB_Name = "GetContentByPosition_Module"
Function GetContentByPosition(ByVal pstrPdfFilename As String)
    
    Set ContentDic = CreateObject("scripting.dictionary")
    
    Dim PDDoc As Object
    Dim CAcroRect As New Acrobat.AcroRect
    Dim PDPage As Acrobat.AcroPDPage
    Dim PDTxtSelect As Acrobat.AcroPDTextSelect
    Dim CArcoPoint As Acrobat.AcroPoint
    Dim iNumWords As Integer
    Dim iMax As Long
    Dim Content As String
    Dim i As Integer

    Set fso = CreateObject("scripting.FileSystemObject")
    
    If fso.FileExists(pstrPdfFilename) Then
    
        Set PDDoc = CreateObject("AcroExch.PDDoc")
        PDDoc.Open pstrPdfFilename
        
        For PageNum = 0 To PDDoc.GetNumPages() - 1
            
            Content = ""
            
            Set PDPage = PDDoc.AcquirePage(PageNum)
            Set CArcoPoint = PDPage.GetSize()
            
            'Debug.Print CArcoPoint.y, CArcoPoint.x
            
            CAcroRect.Top = 50 'CArcoPoint.y
            CAcroRect.Left = 0
            CAcroRect.Right = CArcoPoint.x 'CArcoPoint.x
            CAcroRect.bottom = 0
            
            Set PDTxtSelect = PDDoc.CreateTextSelect(PageNum, CAcroRect)
            
            If PDTxtSelect Is Nothing Then
                
                Debug.Print "Fail"
                
                iNumWords = 0
                iMax = 0
                Content = ""
    
            Else
                
                Debug.Print "Success"
                
                iNumWords = PDTxtSelect.GetNumText
                iMax = iNumWords - 1

                For i = 0 To iMax
                    Content = Content & PDTxtSelect.GetText(i)
                Next
                
            End If
            
            ContentDic(PageNum) = Content
            
        Next PageNum
        
        PDDoc.Close
        
    End If
    
    Set fso = Nothing
    Set PDDoc = Nothing
    Set CAcroRect = Nothing
    Set PDPage = Nothing
    Set PDTxtSelect = Nothing
    Set CArcoPoint = Nothing
    
    Set GetContentByPosition = ContentDic
    
End Function

