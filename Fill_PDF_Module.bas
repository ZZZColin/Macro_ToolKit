Attribute VB_Name = "Fill_PDF_Module"
Function fill_pdf(FileNm, FieldName)

    Dim FileNm, gApp, avDoc, pdDoc, jso

    Set gApp = CreateObject("AcroExch.app")
    
    Set avDoc = CreateObject("AcroExch.AVDoc")
    If avDoc.Open(FileNm, "") Then
        Set pdDoc = avDoc.GetPDDoc()
        Set jso = pdDoc.GetJSObject
    
        jso.getField(FieldName).Value = "myValue"
        pdDoc.Save PDSaveIncremental, FileNm 'Save changes to the PDF document
        pdDoc.Close
    End If
    
    'Close the PDF; the True parameter prevents the Save As dialog from showing
    avDoc.Close (True)
    
    'Some cleaning
    Set gApp = Nothing
    Set avDoc = Nothing
    Set pdDoc = Nothing
    Set jso = Nothing

End Function
