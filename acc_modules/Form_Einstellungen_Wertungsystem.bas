Option Compare Database

Private Sub Bezeichnungsfeld21_Click()
    Me!PROP_VALUE = Null
End Sub

Private Sub PROP_VALUE_AfterUpdate()
    If Forms!einstellungen!Untergeordnet96.Form!PROP_VALUE = "EWS2" Then
        Forms!einstellungen!Einstellungen_Properties.Visible = True
        Forms!einstellungen!Text18.Visible = True
        Forms!einstellungen!Untergeordnet66.Visible = True
        Forms!einstellungen!Text19.Visible = True
    Else
        Forms!einstellungen!Einstellungen_Properties.Visible = False
        Forms!einstellungen!Text18.Visible = False
        Forms!einstellungen!Untergeordnet66.Visible = False
        Forms!einstellungen!Text19.Visible = False
    End If

End Sub
