Option Compare Database

Private Sub Bezeichnungsfeld21_Click()
    Me!PROP_VALUE = Null
End Sub

Private Sub PROP_VALUE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Then
        Select Case Me!PROP_KEY
            Case "Netzwerkname"
                Forms!Einstellungen!Untergeordnet88.SetFocus
                
            Case "Netzwerkname2"            'Untergeordnet88
                Forms!Einstellungen!Untergeordnet75.SetFocus
                
            Case "WLanKW"                   'Untergeordnet75
                Forms!Einstellungen!Einstellungen_Runden.SetFocus
        End Select
        DoCmd.CancelEvent
    End If
End Sub


