Attribute VB_Name = "RemoveInvalidCharacter_Module"
Function RemoveInvalidCharacter(wording)

    arr = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    
    For Each a In arr
        
        wording = Replace(wording, a, "")
        
    Next
    
    RemoveInvalidCharacter = wording
    
End Function
