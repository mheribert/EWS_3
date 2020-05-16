Option Compare Database

Private Sub AuswahlStartklasse2_AfterUpdate()
    DoCmd.Requery "UForm_Ohne_Buch_BW"
End Sub

Private Sub Befehl0_Click()
 DoCmd.Close
End Sub

Private Sub skl_AfterUpdate()
 Forms!paare_ohne_startbuch!Unter_Form_Paare_ohne_buch.Form!klasse = Forms!paare_ohne_startbuch!skl.Column(1)
 Me.Refresh

End Sub
