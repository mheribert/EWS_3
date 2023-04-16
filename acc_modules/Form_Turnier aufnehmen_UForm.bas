Option Compare Database

Private Sub Kombinationsfeld8_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub SelectWR_AfterUpdate()
    If IsNull(Me!SelectWR.Column(0)) Then
        MsgBox "falsche Eingabe!"
        Me!SelectWR = Null
    Else
        Me!AnzahlWR = Me!SelectWR.Column(0)
    End If
End Sub

Private Sub SelectWR_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub
