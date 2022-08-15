Attribute VB_Name = "Menu_Module"
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)

    Dim objPopup As CommandBarButton
    
    If Target.Column = 3 Then
        With Application.CommandBars("cell")
            .Reset
            Set objPopup = .Controls.Add(msoControlButton)
            With objPopup
                .Caption = "²é¿´"
                .OnAction = "a"
            End With
         End With
    Else:
        Application.CommandBars("cell").Reset
    End If
    
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If Target.Column <> 3 Then Application.CommandBars("cell").Reset
    
End Sub
