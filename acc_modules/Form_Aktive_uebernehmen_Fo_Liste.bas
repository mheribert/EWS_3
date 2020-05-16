Option Compare Database

Private Sub Named_Click()
    Forms!Aktive_uebernehmen!formationsname = formationsname
    Forms!Aktive_uebernehmen!Clubname_kurz = Clubname_kurz
    Forms!Aktive_uebernehmen!FBuch = Buchnume
    
'    Dim skStr As String
'
'        If ([Boogie-Woogie] = True) Then
'            skStr = "F_BW"
'        ElseIf ([Rock_n_Roll] = True) Then
'            skStr = "F_RR"
'        End If
'
'        If ([Feld1] = True) Then
'            skStr = skStr & "_LF"
'        ElseIf ([Feld2] = True) Then
'            skStr = skStr & "_BS"
'        ElseIf ([Feld3] = True) Then
'            skStr = skStr & "_M"
'        ElseIf ([Feld4] = True) Then
'            skStr = skStr & "_GF"
'        ElseIf ([Feld5] = True) Then
'            skStr = skStr & "_J"
'        ElseIf ([Feld6] = True) Then
'            skStr = skStr & "_ST"
'        Else
'            'MsgBox "Formation wird als DUO-Formation übernommen."
'            skStr = skStr & "_DUO"
'        End If
    Forms!Aktive_uebernehmen!FStartklasse = [Startklasse]
End Sub

Private Sub Named_DblClick(Cancel As Integer)
    Form_Aktive_uebernehmen.Befehl34_Click
End Sub
