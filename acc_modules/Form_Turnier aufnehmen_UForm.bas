Option Compare Database

Private Sub Kombinationsfeld8_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub SelectWR_AfterUpdate()
    Me!AnzahlWR = Me!SelectWR.Column(0)
End Sub

Private Sub SelectWR_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub
