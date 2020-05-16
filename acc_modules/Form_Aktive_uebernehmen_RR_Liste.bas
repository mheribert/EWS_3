Option Compare Database


Private Sub Named_Click()
    Forms!Aktive_uebernehmen!VName_Dame = Da_Vorname
    Forms!Aktive_uebernehmen!NName_Dame = Da_NAchname
    Forms!Aktive_uebernehmen!VName_Herr = He_Vorname
    Forms!Aktive_uebernehmen!NName_Herr = He_Nachname
    Forms!Aktive_uebernehmen!STBuchnum = Buchnr
    Forms!Aktive_uebernehmen!Alter_Dame = Da_Alterskontrolle
    Forms!Aktive_uebernehmen!Alter_Herr = He_Alterskontrolle
End Sub

Private Sub Named_DblClick(Cancel As Integer)
    Form_Aktive_uebernehmen.btnAddPaar_Click
End Sub
