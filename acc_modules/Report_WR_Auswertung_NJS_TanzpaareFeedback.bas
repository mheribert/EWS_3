Option Compare Database


Private Sub Report_Activate()
    If Nz(Forms![Ausdrucke]![Runde_einstellen]) = "" Then
        MsgBox ("Bitte Runde auswählen")
        DoCmd.Close acReport, Me.Name
    End If
End Sub

Private Sub Report_Load()
    Dim fil As String
    If InStr(1, Forms![Ausdrucke]![Runde_einstellen], "schnell") > 0 Then
        'Left([runde],4)
        fil = "Runde LIKE '" & left(Forms![Ausdrucke]![Runde_einstellen], 4) & "*' AND Startklasse = '" & Forms![Ausdrucke]!Startklasse_einstellen & "'"
    Else
        fil = "Runde = '" & Forms![Ausdrucke]![Runde_einstellen] & "' AND Startklasse = '" & Forms![Ausdrucke]!Startklasse_einstellen & "'"
    End If
    Me.Filter = fil
    Me.FilterOn = True

End Sub
