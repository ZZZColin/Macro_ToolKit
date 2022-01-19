Attribute VB_Name = "PDF2Excel_Module"
Sub main()
     
    FolderPath = ""
    
    Set fso = CreateObject("scripting.filesystemobject")
    
    'Initialize Acrobat by creating App object.
    Set objAcroApp = CreateObject("AcroExch.App")

    'Set AVDoc object.
    Set objAcroAVDoc = CreateObject("AcroExch.AVDoc")
    
    'fso.FileExists(Filepath)

    PDFPath = Dir(FolderPath & "\*.pdf")
    
    Do While PDFPath <> ""
        
        PDFName = Left(PDFPath, Len(PDFPath) - 4)
        
        ExcelPath = FolderPath & "\" & PDFName & ".xlsx"
        
        FileIndex = 1
        
        While fso.FileExists(ExcelPath)

            ExcelPath = FolderPath & "\" & PDFName & " - " & FileIndex & ".xlsx"
            
            FileIndex = FileIndex + 1
            
        Wend

        PDFPath = FolderPath & "\" & PDFPath

        'Open the PDF file.
        boResult = objAcroAVDoc.Open(PDFPath, "")
        
        'Set the PDDoc object.
        Set objAcroPDDoc = objAcroAVDoc.GetPDDoc
        
        'Set the JS Object - Java Script Object.
        Set objJSO = objAcroPDDoc.GetJSObject
        
        'Save PDF file to the new format.
        boResult = objJSO.SaveAs(ExcelPath, "com.adobe.acrobat.xlsx")
        
        'Close the PDF file without saving the changes.
        boResult = objAcroAVDoc.Close(True)
    
        'Release the objects.
        Set objAcroPDDoc = Nothing
        
        PDFPath = Dir
        If PDFPath = "" Then
            Exit Do
        End If

    Loop

    'Close the Acrobat application.
    boResult = objAcroApp.Exit
    
    Set objAcroAVDoc = Nothing
    Set objAcroApp = Nothing
    
    MsgBox ("Done")

End Sub
