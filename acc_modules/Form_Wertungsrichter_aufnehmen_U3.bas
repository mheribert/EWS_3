Option Compare Database

Private Sub Text15_Click()
    Forms!Wertungsrichter_aufnehmen!VName = WVorname
    Forms!Wertungsrichter_aufnehmen!NName = WName
    Forms!Wertungsrichter_aufnehmen!Lizenznr = Lizenzn
    Forms!Wertungsrichter_aufnehmen!Club = Club
End Sub

Private Sub Text15_DblClick(Cancel As Integer)
    Form_Wertungsrichter_aufnehmen.btnAddOffiziellen_Click
End Sub

