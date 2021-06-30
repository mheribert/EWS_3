Option Compare Database

Private Sub Named_Click()
    Forms!Aktive_uebernehmen!BVName_Dame = Vorname
    Forms!Aktive_uebernehmen!BNName_Dame = Nachname
    Forms!Aktive_uebernehmen!BSTkarteD = Buchnr
    Forms!Aktive_uebernehmen!BAlter_Dame = [Geb-Dat-geprüft]
    Forms!Aktive_uebernehmen!BVName_Herr = HRV
    Forms!Aktive_uebernehmen!BNName_Herr = HRN
    Forms!Aktive_uebernehmen!BSTkarteH = Buchnr

End Sub

Private Sub Named_DblClick(Cancel As Integer)
    Named_Click
    Form_Aktive_uebernehmen.Befehl114_Click
End Sub
