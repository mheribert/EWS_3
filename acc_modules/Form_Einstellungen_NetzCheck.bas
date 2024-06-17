Option Compare Database
Option Explicit

Private Sub Bezeichnungsfeld21_Click()
    Me!PROP_VALUE = Null
End Sub

Private Sub PROP_VALUE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Then
        Forms!Einstellungen!IPAddr.SetFocus
    End If
    DoCmd.CancelEvent
End Sub

