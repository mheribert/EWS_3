Option Compare Database

Private Sub Text15_Click()
Forms!TL_BS_aufnehmen!VName = WVorname
Forms!TL_BS_aufnehmen!NName = WName
Forms!TL_BS_aufnehmen!Lizenznr = Lizenzn
Forms!TL_BS_aufnehmen!Club = Club
End Sub

Private Sub Text15_DblClick(Cancel As Integer)
    Form_TL_BS_aufnehmen.btnAddOffiziellen_Click
End Sub

