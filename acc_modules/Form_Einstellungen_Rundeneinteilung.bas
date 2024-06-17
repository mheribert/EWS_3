Option Compare Database

Private Sub Text15_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 And Me.NewRecord = True Then
        Forms!Einstellungen!Untergeordnet96.SetFocus
    End If
End Sub
