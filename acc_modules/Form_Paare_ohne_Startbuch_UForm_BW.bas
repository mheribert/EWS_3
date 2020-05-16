
Private Sub btnDruckEntschuldigung_Click()
    If (SBS_ID_BW_D = 0 And SBS_ID_BW_H = 0) Then
        MsgBox "Dieses Paar hat seine Startkarten nicht vergessen!"
        Exit Sub
    End If
    
    [Form_A-Programmübersicht]![Report_TP_ID] = TP_ID
    Dim stDocName As String

    stDocName = "Bestaetigung_ohne_Buch"
    DoCmd.OpenReport stDocName, acPreview

End Sub

Private Sub btnHaftungsausschluss_Click()
    stDocName = "Haftungsausschluss"
    DoCmd.OpenReport stDocName, acPreview, , "TP_ID = " & [TP_ID]

End Sub
