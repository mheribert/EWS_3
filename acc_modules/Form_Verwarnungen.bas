Option Compare Database
Option Explicit

Private Sub btn_close_Click()
    DoCmd.Close acForm, "Verwarnungen"
End Sub

Private Sub Form_Current()
    Me!B_Anzahl.Caption = "noch " & (250 - Len(Nz(Me!Verwarnung)) & " Zeichen")

End Sub

Private Sub Verwarnung_KeyUp(KeyCode As Integer, Shift As Integer)
    Me!B_Anzahl.Caption = "noch " & (250 - Len(Nz(Me!Verwarnung.text)) & " Zeichen")
    
End Sub
