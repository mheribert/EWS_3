Option Compare Database

Private Sub btnOK_Click()
On Error GoTo Err_btnOK_Click

    Dim dbs As Database
    Set dbs = CurrentDb
    
    ' Als erstes die Dateneingabe überprüfen
    ' wenn falsche Daten, dann nicht weitermachen
    If Not checkData() Then
        Exit Sub
    End If
    
    ' 1. Alle Paare mit Platz < AnzahlPaareDirektWeiter in der ausgewählten
    Dim Turniernr As Integer
    Turniernr = get_aktTNr
    
    ' Überprüfen bei KO-Runde, dass nur gerade Anzahl weitergenommen wird
    If cbNaechsteRunde.Column(7) Like "*KO*" Then
        'And (AnzahlPaareDirektWeiter Mod 2 <> 0)
        ' Anzahl der Paare bis Platz X ermitteln
        Dim PaareBisPlatz As DAO.Recordset
        Set PaareBisPlatz = dbs.OpenRecordset("SELECT Count(Majoritaet.TP_ID) AS AnzahlvonTP_ID FROM Majoritaet WHERE (((Majoritaet.RT_ID)= " & cbAktuelleRunde & ") AND ((Majoritaet.Platz)<= " & AnzahlPaareDirektWeiter & "));")
        PaareBisPlatz.MoveFirst
        If (PaareBisPlatz!AnzahlvonTP_ID Mod 2 <> 0) Then
            MsgBox "Bitte beachten, dass in der KO-Runde eine gerade Anzahl an Paare sein muss.", vbOKOnly
        End If
    End If
    
    Call PaareInDieNaechsteRunde(Turniernr, cbAktuelleRunde, cbNaechsteRunde, AnzahlPaareDirektWeiter, cbNaechsteRunde.Column(1))
    
    Dim isStichrunde As Boolean
    Dim vonPlatz As Integer
    Dim bisPlatz As Integer
    
    isStichrunde = (grpHoffnungsrunde = 2)
    
    If (cbHoffnungsrundeDurchfuehren = True) Then
          
        vonPlatz = AnzahlPaareDirektWeiter + 1
        If (isStichrunde) Then
            bisPlatz = SRWeiterBisPlatz
        Else
            bisPlatz = 10000
        End If
        
        Call PaareInDieNaechsteRunde2(Turniernr, cbAktuelleRunde, cbHoffnungsrunde, vonPlatz, bisPlatz, cbNaechsteRunde.Column(1))
        Call PaarePlatzieren(cbAktuelleRunde, GetPaareBisPlatz(cbAktuelleRunde, bisPlatz) + 1)
    Else
        ' Ermitteln, wieviele Paare die nächste Runde erreicht haben
        Dim paareInRunde As Integer
        paareInRunde = GetPaareInRunde(cbNaechsteRunde)
        vonPlatz = GetPaareBisPlatz(cbAktuelleRunde, AnzahlPaareDirektWeiter) + 1
        Dim offset As Integer
        offset = paareInRunde - GetPaareBisPlatz(cbAktuelleRunde, AnzahlPaareDirektWeiter)
        
        Call PaarePlatzierenMitHoffnungsrunde(cbAktuelleRunde, vonPlatz, offset)
    End If
    
    Form_Paare_schon_qualifiziert.Requery
    
    DoCmd.Close

Exit_btnOK_Click:
    Exit Sub

Err_btnOK_Click:
    MsgBox err.Description
    Resume Exit_btnOK_Click
    
End Sub

Private Function checkData() As Boolean
    
    If (IsNull(AnzahlPaareDirektWeiter) Or Not IsNumeric(AnzahlPaareDirektWeiter) Or AnzahlPaareDirektWeiter < 1) Then
        MsgBox "Bitte geben Sie die Anzahl der Paare, welche sich direkt für die nächste Runde qualifiziert!"
        checkData = False
        Exit Function
    End If
    
    If (IsNull(cbNaechsteRunde)) Then
        MsgBox "Bitte wählen Sie die nächste Runde aus!"
        checkData = False
        Exit Function
    End If
    
    Dim isStichrunde As Boolean
    isStichrunde = (grpHoffnungsrunde = 2)
    
    If (cbHoffnungsrundeDurchfuehren = True) Then
        If (IsNull(cbHoffnungsrunde)) Then
            MsgBox "Bitte wählen Sie die Hoffnungsrunde aus!"
            checkData = False
            Exit Function
        End If
        If (isStichrunde And (IsNull(SRWeiterBisPlatz) Or Not IsNumeric(SRWeiterBisPlatz)) Or SRWeiterBisPlatz < 1) Then
            MsgBox "Bitte geben Sie an, bis zu welchem Platz die Stichrunde durchgeführt werden soll!"
            checkData = False
            Exit Function
        End If
        If (isStichrunde And AnzahlPaareDirektWeiter >= SRWeiterBisPlatz) Then
            MsgBox "Der Platz für die Stichrunde muss größer dem Platz sein, bis zu dem die Paare direkt weiterkommen!"
            checkData = False
            Exit Function
        End If
    End If
    
    ' Warnung, falls 90% überschritten oder 40% unterschritten
    Dim paareInRunde As Integer
    Dim RT_ID As Integer
    RT_ID = cbAktuelleRunde
    paareInRunde = GetPaareInRunde(RT_ID)
    Dim prozentSatz As Double
    prozentSatz = (AnzahlPaareDirektWeiter / paareInRunde) * 100
    
    If ((prozentSatz < 40) And (cbHoffnungsrundeDurchfuehren = False)) Then
        result = MsgBox("Gemäß TSO haben Sie zu wenige Paare für die nächste Runde ausgewählt. Wollen Sie trotzdem weitermachen?", vbYesNo)
        
        If (result = vbNo) Then
            checkData = False
            Exit Function
        End If
    End If
    
    If (prozentSatz > 90 And (cbHoffnungsrundeDurchfuehren = False)) Then
        result = MsgBox("Gemäß TSO haben Sie zu viele Paare für die nächste Runde ausgewählt. Wollen Sie trotzdem weitermachen?", vbYesNo)
        
        If (result = vbNo) Then
            checkData = False
            Exit Function
        End If
    End If
    
    checkData = True
    
End Function

Private Sub btnAbbrechen_Click()
On Error GoTo Err_btnAbbrechen_Click


    DoCmd.Close

Exit_btnAbbrechen_Click:
    Exit Sub

Err_btnAbbrechen_Click:
    MsgBox err.Description
    Resume Exit_btnAbbrechen_Click
    
End Sub

Private Sub cbHoffnungsrundeDurchfuehren_AfterUpdate()
    Call ActivateHoffnungsrunde
End Sub

Private Sub ActivateHoffnungsrunde()
    cbHoffnungsrunde.Enabled = cbHoffnungsrundeDurchfuehren
    SRWeiterBisPlatz.Enabled = cbHoffnungsrundeDurchfuehren And grpHoffnungsrunde = 2
    optHoffnungsrunde.Enabled = cbHoffnungsrundeDurchfuehren
    optStichrunde.Enabled = cbHoffnungsrundeDurchfuehren
    grpHoffnungsrunde.Enabled = cbHoffnungsrundeDurchfuehren
End Sub

Private Sub cbNaechsteRunde_AfterUpdate()
    Form_Majoritaet_ausrechnen.nächste_Runde = cbNaechsteRunde
End Sub

Private Sub Form_Open(Cancel As Integer)
    Call ActivateHoffnungsrunde
    
    Dim dbs As Database
    Set dbs = CurrentDb
    
    Dim rst As Recordset
    Dim Anzahl As Double
    ProzentGrenze = " "
    ' Anzahl Endrundenpaare ermitteln
    Set rst = dbs.OpenRecordset("select * from Majoritaet where rt_id=" & cbAktuelleRunde)
    
    If Not rst.EOF() Then
        rst.MoveLast
        ProzentGrenze = "40%=" & (rst.RecordCount / 10 * 4) & " / 90%=" & (rst.RecordCount / 10 * 9)
    End If
End Sub

Private Sub grpHoffnungsrunde_AfterUpdate()
    Call ActivateHoffnungsrunde
End Sub
