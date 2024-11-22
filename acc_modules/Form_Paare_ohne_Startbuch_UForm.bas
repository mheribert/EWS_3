

Private Sub btnDruckEntschuldigung_Click()
    If (SBS_ID = 0) Then
        MsgBox "Dieses Paar / Formation hat sein Startbuch nicht vergessen!"
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
