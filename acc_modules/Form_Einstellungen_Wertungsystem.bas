Option Compare Database

Private Sub Bezeichnungsfeld21_Click()
    Me!PROP_VALUE = Null
End Sub

Private Sub PROP_VALUE_AfterUpdate()
    If Forms!Einstellungen!Untergeordnet96.Form!PROP_VALUE = "EWS2" Then
        Forms!Einstellungen!Einstellungen_Properties.Visible = True
        Forms!Einstellungen!Text18.Visible = True
        Forms!Einstellungen!Untergeordnet66.Visible = True
        Forms!Einstellungen!Text19.Visible = True
    Else
        Forms!Einstellungen!Einstellungen_Properties.Visible = False
        Forms!Einstellungen!Text18.Visible = False
        Forms!Einstellungen!Untergeordnet66.Visible = False
        Forms!Einstellungen!Text19.Visible = False
    End If
End Sub

Private Sub PROP_VALUE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Then
        Forms!Einstellungen!Untergeordnet78.SetFocus
    End If
End Sub

