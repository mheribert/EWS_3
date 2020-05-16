Option Compare Database

Public Function btn(btnNr)
    Dim ctl As Control
    Set ctl = Me(btnNr)
    
    Debug.Print ctl.BackColor
    Forms!einstellungen!Einstellungen_PPT.Form!PPT_Color = ctl.BackColor
    DoCmd.Close acForm, "Einstellungen_Color"

End Function

Private Sub Abbrechen_Click()
    DoCmd.Close acForm, "Einstellungen_Color"
End Sub
