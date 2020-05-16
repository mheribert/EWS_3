Option Compare Database
Option Explicit

Private Sub Befehl2_Click()
    DoCmd.Close
End Sub

Private Sub Befehl43_Click()
    Dim dbs As Database
    Dim re As Recordset
    Set dbs = CurrentDb
    Set re = dbs.OpenRecordset("select * from Rundentab where RT_ID=" & Runde_suchen & ";")
    If Not DLookup("Getrennte_Auslosung", "Turnier", "Turniernum = " & Turniernr) Then
        Runde_übertragen re!Runde, re!Startklasse
    End If
    Requery
End Sub

Private Sub Befehl47_Click()
'****AB**** V13_04 - neue Funktion/button zum Ausdrucken der Observer Wertungsbögen
    Dim stDocName As String
    If Not Me.Runde_suchen = " " Then
        If Me.Runde_suchen.Column(4) = "End_r" Or Me.Runde_suchen.Column(4) = "End_r_akro" Then
            stDocName = "ObserverWertungsbogenEndrunde"
        Else
            stDocName = "ObserverWertungsbogen"
        End If
        DoCmd.OpenReport stDocName, acPreview, , "RT_ID = " & Me.Runde_suchen.Column(0) & ""
    Else
        MsgBox ("Bitte Runde auswählen")
    End If

End Sub

Private Sub Befehl58_Click()
    Dim st          ' Nocheinmal starten
    Dim back
    If Me!nochmal = True Then
        back = MsgBox("Das Paar startet schon nocheinmal!" & vbCrLf & vbCrLf & "Wirklich nochmal starten?", vbYesNo)
    Else
        back = MsgBox("Nocheinmal starten?", vbYesNo)
    End If
    If back = vbNo Then
        Exit Sub
    Else
        If get_properties("EWS") = "EWS3" Then
            st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=nochmal_starten&text=" & TP_ID)
            If st = "eingetragen" Then
                Me!nochmal = True
                DoCmd.Requery
            Else
                MsgBox "Die Wiederholung der Runde wurde nicht eingetragen!"
            End If
        End If
    End If
End Sub

Private Sub btnDruckRundeneinteilung_Click()
    Dim db As Database
    Dim re As Recordset
    Set db = CurrentDb()
    ' ***** HM14.03 *****
    ' man kann jetzt die Kopie für Turnierunterlagen weglassen
    ' Hier check ob mindestens Eine eingegeben ist.
    If IsNull(Me!Runde_suchen) Then
        MsgBox "Bitte Runde wählen"
    Else
    
    
        Set re = db.OpenRecordset("SELECT COUNT(*) AS anz FROM Kopien WHERE Kopie_an <> 'HTML-Seiten' AND Kopie_an <> 'PPT-Folien' AND Kopie_an <> 'HTML-Moderator';")
        If re!anz = 0 Then
            MsgBox "Es wurden keine Kopien in Einstellungen angelegt!"
            Exit Sub
        End If
        ' *****
        Set re = db.OpenRecordset("Select * from Kopien where T_ID =" & get_aktTNr & " AND Kopie_an= ""PPT-Folien"";")
        If re.RecordCount > 0 Then
            Call FolieRunden_Click
        End If
        Set re = db.OpenRecordset("Select * from Kopien where T_ID =" & get_aktTNr & " AND Kopie_an= ""HTML-Seiten"";")
        If re.RecordCount > 0 Then
            HTML_Seiten_Click
        End If
        
        
        [Form_A-Programmübersicht].Report_RT_ID = Me.Runde_suchen
    
    
        DoCmd.OpenReport "Startliste_Runden", acPreview
    End If
    
End Sub

Private Sub btnPaareInDieserRunde_Click()
    [Form_A-Programmübersicht].Report_RT_ID = Me.Runde_suchen
    
On Error GoTo Err_btnDruckRundeneinteilung_Click

    Dim stDocName As String
    If IsNull(Me!Runde_suchen) Then
        MsgBox "Bitte Runde wählen"
    Else
        stDocName = "Startliste_startende_Paare"
        DoCmd.OpenReport stDocName, acPreview
    End If

Exit_btnDruckRundeneinteilung_Click:
    Exit Sub

Err_btnDruckRundeneinteilung_Click:
    MsgBox err.Description
    Resume Exit_btnDruckRundeneinteilung_Click
End Sub

Private Sub btnRundeneinteilungZeit_Click()
    [Form_A-Programmübersicht].Report_RT_ID = Me.Runde_suchen
    
On Error GoTo Err_btnDruckRundeneinteilung_Click

    Dim stDocName As String
    If IsNull(Me!Runde_suchen) Then
        MsgBox "Bitte Runde wählen"
    Else
        stDocName = "Startliste_Runden_Zeit"
        DoCmd.OpenReport stDocName, acPreview
    End If

Exit_btnDruckRundeneinteilung_Click:
    Exit Sub

Err_btnDruckRundeneinteilung_Click:
    MsgBox err.Description
    Resume Exit_btnDruckRundeneinteilung_Click
End Sub

Private Sub FolieRunden_Click()
    Dim dbs As Database
    Dim re As Recordset
    Set dbs = CurrentDb
    Dim t As Long
    
    If IsNull(Me!Runde_suchen) Then
        MsgBox "Bitte Runde wählen"
    Else
        If Me.RecordsetClone.RecordCount = 0 Then
            MsgBox "Es gibt keine Paare in dieser Runde"
        Else
            If InStr(1, Me!Runde_suchen.Column(4), "_Akro") > 0 And Me!Runde_suchen.Column(2) = False Then
                Set re = dbs.OpenRecordset("SELECT * from RundenTab WHERE Startklasse = '" & Me!Runde_suchen.Column(6) & "' AND Runde = '" & Mid(Me!Runde_suchen.Column(4), 1, 3) & "_r_Fuß';", DB_OPEN_DYNASET)
                If re.EOF Then
                    MsgBox "Es fehlt die Fußtechnikrunde!"
                Else
                    Call gen_Folien(Me.RecordsetClone, Me!Runde_suchen.Column(7), Mid(Me!Runde_suchen.Column(4), 1, 3) & "runde Fußtechnik", re!Rundenreihenfolge)
                    For t = 1 To 10000000: Next  ' warten dass PPT zu ist
                End If
            End If
            Call gen_Folien(Me.RecordsetClone, Me!Runde_suchen.Column(7), Me!Runde_suchen.Column(8), Trim(str(Me!Runde_suchen.Column(12))))
        End If
    End If
End Sub

Private Sub Form_Load()
    Select Case Forms![A-Programmübersicht]!Turnierausw.Column(8)
        Case "SL"
            Me!FolieRunden.Visible = True
            Me!Rechteck47.Visible = True
        Case "BY"
            Me!FolieRunden.Visible = True
            Me!Rechteck47.Visible = True
            
        Case Else
    End Select

End Sub

Private Sub HTML_Seiten_Click()
    Dim dbs As Database
    Dim re As Recordset
    Dim rde, rd As String
    Dim f_rt As Integer
    Requery
    Set dbs = CurrentDb
    
    If IsNull(Me!Runde_suchen) Then
        MsgBox "Bitte Runde wählen"
    Else
        If Me.RecordsetClone.RecordCount = 0 Then
            MsgBox "Es gibt keine Paare in dieser Runde"
        Else
            Me!Runde_suchen.Locked = True
            rde = Mid(Me!Runde_suchen.Column(4), 1, 6)
            If InStr(1, Me!Runde_suchen.Column(4), "_Akro") > 0 And Me!Runde_suchen.Column(2) = False Then 'Hier wird bei A/B Fuß und Akro erstellt
                Set re = get_rde(Me!Runde_suchen.Column(6), rde & "Fuß")
                'dbs.OpenRecordset("SELECT * from RundenTab WHERE Startklasse = '" & Me!Runde_suchen.Column(6) & "' AND Runde = '" & rde & "Fuß';", DB_OPEN_DYNASET)
                If re.EOF Then 'keine Runde vorhanden
                    MsgBox "Es fehlt die Fußtechnikrunde!"
                Else
                    rde = re!Runde
                    f_rt = re!RT_ID
                    rd = re!Rundentext
                    Set re = Me.RecordsetClone
                    Call build_html(re, f_rt, rde)
                    make_a_round Me.RecordsetClone, Me!Runde_suchen.Column(7), rd, f_rt
                End If
            End If
            'normale Runden erstellen
            Set re = Me.RecordsetClone
            Call build_html(re, Me!Runde_suchen.Column(0), Me!Runde_suchen.Column(4))
            make_a_round Me.RecordsetClone, Me!Runde_suchen.Column(7), Me!Runde_suchen.Column(8), Me!Runde_suchen.Column(0)
            Me!Runde_suchen.Locked = False
        End If
    End If
    DoCmd.Requery

End Sub

Private Sub Kombinationsfeld32_AfterUpdate()
    If (Not hasWertungen(TP_ID)) Then
        Dim dbs As Database
        Set dbs = CurrentDb
        Dim rst As Recordset
        Dim stmt As String
        stmt = "Select * from Paare p where tp_id=" & TP_ID
        Set rst = dbs.OpenRecordset(stmt)
        Do While (Not rst.EOF)
            rst.Edit
            rst!Anwesent_Status = Anwesend_Status
            rst.Update
            rst.MoveNext
        Loop
        rst.Close
    End If
    If Me!Kombinationsfeld32 = 2 Then
        Me!Runde = Null
    End If
End Sub

Private Sub Kombinationsfeld32_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub Kontrollkästchen44_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub runde_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub Runde_suchen_AfterUpdate()
    
    ' Falls Rundeneinteilung in einer geteilten Endrunde, dann die Paare aus dem
    ' anderen Durchgang mit diesem abgleichen
    If (IsNull(Runde_suchen.Column(1))) Then
        Exit Sub
    End If
    
    Dim dbs As Database
    Dim rs As Recordset
    
    ' Bezug auf aktuelle Datenbank zurückgeben.
    Set dbs = CurrentDb
    
    ' Ermittlung der Tanzrunde
    Dim sqlstr As String
    Dim Tanzrund, Startklasse, Turniernr As String
    Dim InRundeneinteilung As Integer
    sqlstr = "select rt.Runde, rt.Startklasse, rt.Turniernr, tr.InRundeneinteilung from Rundentab rt, Tanz_Runden tr where tr.Runde=rt.Runde and rt.RT_ID=" & Runde_suchen
    Set rs = dbs.OpenRecordset(sqlstr)
    Tanzrund = rs!Runde
    Startklasse = rs!Startklasse
    Turniernr = rs!Turniernr
    InRundeneinteilung = rs!InRundeneinteilung
    rs.Close
    
    If (InRundeneinteilung = 2) Then
        Dim MasterRunde As Integer
        Dim MasterRunde_Text As String
        
        MasterRunde_Text = "NEIN"
        
        If (Tanzrund = "End_r_Fuß") Then
            MasterRunde_Text = "End_r_Akro"
        ElseIf (Tanzrund = "End_r_lang") Then
            MasterRunde_Text = "End_r"
        ElseIf (Tanzrund = "End_r_schnell") Then
            MasterRunde_Text = "End_r"
        End If
        
        If (MasterRunde_Text <> "NEIN") Then
            sqlstr = "select * from Rundentab where Turniernr=" & Turniernr & " and Runde='" & MasterRunde_Text & "' and Startklasse='" & Startklasse & "'"
            Set rs = dbs.OpenRecordset(sqlstr)
            If (rs.NoMatch) Then
                MsgBox ("Ed wurde die dazugehörige Akrobatikrunde nicht gefunden!")
                rs.Close
                Exit Sub
            End If
            MasterRunde = rs!RT_ID
            rs.Close
            Call UpdateRundenqualifikation(MasterRunde, Runde_suchen, False)
        End If
    End If
    
    Me.Requery
    Me!Feld138.SetFocus
End Sub

Private Sub Auslosung_Click()
    If Me!Feld52 = 1 Then
        zufallszahl
    Else
        umgekehrte_Reihenfolge
    End If
End Sub

Private Sub umgekehrte_Reihenfolge()
    Dim dbs As Database
    Dim rstauslosung As Recordset
    Dim rs As Recordset
    Dim rstpaare, RundenPaare As Recordset
    Dim Fußtechnik_checken As Boolean
    Dim reihenf As Integer
    Dim Anzahl, zufall, trunde, was As Integer
    Dim fil As String
    
    If (IsNull(Runde_suchen.Column(1))) Then
        Exit Sub
    End If
    
    Set dbs = CurrentDb
    
    Set rs = dbs.OpenRecordset("SELECT * FROM Rundentab INNER JOIN Tanz_Runden_fix ON Rundentab.Runde = Tanz_Runden_fix.Runde WHERE (Rundentab.Rundenreihenfolge < " & Me!Runde_suchen.Column(12) & ") And (Rundentab.Startklasse = '" & Me!Runde_suchen.Column(6) & "') ORDER BY Rundentab.Rundenreihenfolge;")

    If rs.RecordCount > 0 Then
        rs.MoveLast
        fil = rs!RT_ID
        If rs![Rundentab.Runde] = "Hoff_r" Or rs![Rundentab.Runde] = "Stich_r" Then
            rs.MovePrevious
            fil = fil & " OR Paare_Rundenqualifikation.RT_ID=" & rs!RT_ID
        End If
    Else
        MsgBox "Es gibt keine Tanzunde vor    " & Me!Feld138
        Exit Sub
    End If
    If Me!Feld52 = 2 Then
        ' Startreihenfolge
        Set rstpaare = dbs.OpenRecordset("SELECT Paare_Rundenqualifikation.TP_ID, Count(Paare_Rundenqualifikation.RT_ID) AS AnzahlvonRT_ID, Last(Paare_Rundenqualifikation.Rundennummer) AS LetzterWertvonRundennummer FROM Paare_Rundenqualifikation WHERE (Paare_Rundenqualifikation.RT_ID=" & fil & ") GROUP BY Paare_Rundenqualifikation.TP_ID ORDER BY Count(Paare_Rundenqualifikation.RT_ID) DESC, Last(Paare_Rundenqualifikation.Rundennummer) DESC;")
    ElseIf Me!Feld52 = 3 Then
        ' umgekehrte Platzierung
        Set rstpaare = dbs.OpenRecordset("SELECT Majoritaet.TP_ID, Count(Majoritaet.RT_ID) AS AnzahlvonRT_ID, Min(Majoritaet.Platz) AS Platzierung, Last(Paare_Rundenqualifikation.Rundennummer) AS LetzterWertvonRundennummer FROM Majoritaet INNER JOIN Paare_Rundenqualifikation ON (Majoritaet.RT_ID = Paare_Rundenqualifikation.RT_ID) AND (Majoritaet.TP_ID = Paare_Rundenqualifikation.TP_ID) WHERE (Paare_Rundenqualifikation.RT_ID=" & fil & ") GROUP BY Majoritaet.TP_ID ORDER BY Count(Majoritaet.RT_ID) DESC, Min(Majoritaet.Platz) DESC;")
    ElseIf Me!Feld52 = 4 Then
        ' gleiche Platzierung
        Set rstpaare = dbs.OpenRecordset("SELECT Majoritaet.TP_ID, Count(Majoritaet.RT_ID) AS AnzahlvonRT_ID, Min(Majoritaet.Platz) AS Platzierung, Last(Paare_Rundenqualifikation.Rundennummer) AS LetzterWertvonRundennummer FROM Majoritaet INNER JOIN Paare_Rundenqualifikation ON (Majoritaet.RT_ID = Paare_Rundenqualifikation.RT_ID) AND (Majoritaet.TP_ID = Paare_Rundenqualifikation.TP_ID) WHERE (Paare_Rundenqualifikation.RT_ID=" & fil & ") GROUP BY Majoritaet.TP_ID ORDER BY Count(Majoritaet.RT_ID), Min(Majoritaet.Platz) DESC;")
    Else
        MsgBox "Fehler bei der Sortierreihenfolge!"
    End If
    
    Set RundenPaare = dbs.OpenRecordset("SELECT * FROM Paare_Rundenqualifikation WHERE Paare_Rundenqualifikation.RT_ID=" & Runde_suchen & ";")
    
    reihenf = 0
    
    If rstpaare.RecordCount > 0 Then
        rstpaare.MoveFirst
        Do Until rstpaare.EOF
            RundenPaare.FindFirst "TP_ID = " & rstpaare.TP_ID
            If Not RundenPaare.NoMatch Then
                RundenPaare.Edit
                RundenPaare!Rundennummer = Int(reihenf / Me!Paaranzahl) + 1
                RundenPaare.Update
                reihenf = reihenf + 1
            End If
            rstpaare.MoveNext
        Loop
    Else
        MsgBox "Es gibt keine Platzierungen aus der vorhergehenden Runde!"
    End If
    DoCmd.Requery
    Exit Sub
    
    
    
    
    
    
    
    
    
    
    
    
    Fußtechnik_checken = False
    
    
    'Wenn es sich um eine Endrunde mit Fuß- und Akrobatikrunde handelt muss bei der Auslosung die FT-Runde gecheckt werden
    If Runde_suchen.Column(1) = "A-Klasse Endrunde" Or Runde_suchen.Column(1) = "B-Klasse Endrunde" Then
        Fußtechnik_checken = True
    End If
    
    ' Ermittlung der Tanzrunde
    Dim sqlstr As String
    Dim Tanzrunde, Startklasse, Turniernr As String
    sqlstr = "select * from Rundentab rt where rt.RT_ID=" & Runde_suchen
    Set rs = dbs.OpenRecordset(sqlstr)
    Tanzrunde = rs!Runde
    Startklasse = rs!Startklasse
    Turniernr = rs!Turniernr
    rs.Close
    
    'Wenn es sich nicht um eine Endrunde handelt, dann kann keine Auslosung in umgekehrter Reihenfolge gemacht werden
    'hier in Zukunft eventuell Abzweig möglich wenn Auslosung erster gegen letzte stattfinden soll
    If Not Tanzrunde Like "*End*" Then
        MsgBox "Auslosung in umgekehrter Reihenfolge nur in der Endrunde möglich!", vbOKOnly, "Auslosung umgekehrte Reihenfolge"
        Exit Sub
    End If
    
    sqlstr = "select * from Paare_Rundenqualifikation where RT_ID= " & Runde_suchen
    Set rstauslosung = dbs.OpenRecordset(sqlstr)
    
    
    ' Abbruch, wenn keine Daten vorhanden sind
    If (rstauslosung.EOF) Then
        rstauslosung.Close
        Exit Sub
    End If
    
    'Abbruch, wenn keine Rock'n'Roll Turnierklasse
    If Not Startklasse Like "RR*" And Not Startklasse Like "BW*" Then
        rstauslosung.Close
        Exit Sub
    End If
    
    
    ' vorherige Tanzrunde herausfinden
    Dim vorherigeTanzrundeID, FußtechnikrundeID, AkrobatikrundeID As Long
    sqlstr = "select * from Rundentab rt where rt.Startklasse='" & Startklasse & "' ORDER BY Rundenreihenfolge"
    Set rs = dbs.OpenRecordset(sqlstr)
    rs.FindFirst "Runde = '" & Tanzrunde & "'"
    If Not rs.NoMatch Then
        rs.MovePrevious
        vorherigeTanzrundeID = rs!RT_ID
        If Fußtechnik_checken Then
            If rs!Runde = "End_r_Fuß" Then
                FußtechnikrundeID = rs!RT_ID
                rs.MovePrevious
                vorherigeTanzrundeID = rs!RT_ID
                'AkrobatikrundeID = Tanzrunde
            End If
        End If
    End If
    rs.Close
    
    
    If Fußtechnik_checken Then
        'Prüfen ob in der Fußtechnikrunde schon Daten drin stehen, dann anhand dieser die Rundeneinteilung vornehmen
        sqlstr = "SELECT Paare_Rundenqualifikation.RT_ID, Majoritaet.RT_ID, Majoritaet.Platz, Paare_Rundenqualifikation.Rundennummer, Paare_Rundenqualifikation.TP_ID FROM Paare_Rundenqualifikation INNER JOIN Majoritaet ON Paare_Rundenqualifikation.TP_ID = Majoritaet.TP_ID WHERE (((Paare_Rundenqualifikation.RT_ID)=" & Runde_suchen & ") AND (Majoritaet.RT_ID)= " & FußtechnikrundeID & " ) ORDER BY Majoritaet.WR7;"
        Set rstauslosung = dbs.OpenRecordset(sqlstr)
        If (rstauslosung.EOF) Then
            'wenn in der Fußtechnikrunde noch keine Ergebnisse, dann die vorherige Runde wählen
            sqlstr = "SELECT Paare_Rundenqualifikation.RT_ID, Majoritaet.RT_ID, Majoritaet.Platz, Paare_Rundenqualifikation.Rundennummer, Paare_Rundenqualifikation.TP_ID FROM Paare_Rundenqualifikation INNER JOIN Majoritaet ON Paare_Rundenqualifikation.TP_ID = Majoritaet.TP_ID WHERE (((Paare_Rundenqualifikation.RT_ID)=" & Runde_suchen & ") AND (Majoritaet.RT_ID)= " & vorherigeTanzrundeID & " ) ORDER BY Majoritaet.WR7;"
            Set rstauslosung = dbs.OpenRecordset(sqlstr)
        End If
    Else
        ' Ergebnis der vorherigen Runde zur Startreiehnfolge nutzen
        sqlstr = "SELECT Paare_Rundenqualifikation.RT_ID, Majoritaet.RT_ID, Majoritaet.Platz, Paare_Rundenqualifikation.Rundennummer, Paare_Rundenqualifikation.TP_ID FROM Paare_Rundenqualifikation INNER JOIN Majoritaet ON Paare_Rundenqualifikation.TP_ID = Majoritaet.TP_ID WHERE (((Paare_Rundenqualifikation.RT_ID)=" & Runde_suchen & ") AND (Majoritaet.RT_ID)= " & vorherigeTanzrundeID & " ) ORDER BY Majoritaet.WR7;"
        Set rstauslosung = dbs.OpenRecordset(sqlstr)
    End If

    Set RundenPaare = dbs.OpenRecordset("SELECT Paare_Rundenqualifikation.RT_ID, Paare_Rundenqualifikation.TP_ID, Paare_Rundenqualifikation.Rundennummer FROM Paare_Rundenqualifikation WHERE (((Paare_Rundenqualifikation.RT_ID)= " & Runde_suchen & " ));")
    
    If (rstauslosung.EOF) Then
        MsgBox "Noch keine Ergebnisse in der vorher getanzten Runde vorhanden!", vbOKOnly
        rstauslosung.Close
        Exit Sub
    End If
    
    rstauslosung.MoveLast
    trunde = 1
    Anzahl = rstauslosung.RecordCount
    rstauslosung.MoveFirst
    Do While Not rstauslosung.EOF()
        RundenPaare.FindFirst "TP_ID = " & rstauslosung!TP_ID
        If Not RundenPaare.NoMatch Then
            RundenPaare.Edit
            RundenPaare!Rundennummer = trunde
            RundenPaare.Update
        End If
        rstauslosung.MoveNext
        trunde = trunde + 1
    Loop
    rstauslosung.Close
    


    ' Wenn Vorrunde oder Endrunde der RR-A oder RR-B
    ' dann die Rundeneinteilung in die Fußtechnik und Akrobatik
    ' Endrunde übernehmen
    ' DLookup("Getrennte_Auslosung", "Turnier", "Turniernum = " & Turniernr) = true
    If InStr(1, Tanzrunde, "_Akro") > 0 And Me!Runde_suchen.Column(2) = False Then
        Dim stmtr As String
        Dim rstr As Recordset
        Dim rt_id_er_fuss As Integer
        
        stmtr = "Select * from Rundentab where Turniernr=" & Turniernr & " and Startklasse='" & Startklasse & "' and Runde = '" & left(Tanzrunde, 3) & "_r_Fuß'"
        
        Set rstr = dbs.OpenRecordset(stmtr)
        If (rstr.NoMatch) Then
            MsgBox "Fußtechnik Enrunde für RR wurde nicht gefunden!"
            GoTo BW_RR_Error
        End If
        rt_id_er_fuss = rstr!RT_ID
        rstr.Close
        
        Call UpdateRundenqualifikation(Runde_suchen, rt_id_er_fuss, True)
    End If
    
BW_RR_Error:
    dbs.Close
    
    Me.Requery
    DoCmd.RepaintObject , ""
    DoCmd.GoToRecord , "", acFirst
    DoCmd.SetWarnings True

End Sub

Private Sub zufallszahl()
    Dim dbs As Database
    Dim rstauslosung As Recordset
    Dim rs As Recordset
    Dim rstpaare As Recordset
    Dim stmt As String
    Dim rst As Recordset
    Dim Anzahl, zufall, trunde, was As Integer
    
    Set dbs = CurrentDb
    
    If (IsNull(Runde_suchen.Column(1))) Then
        Exit Sub
    End If
    
    ' Ermittlung der Tanzrunde
    Dim sqlstr As String
    Dim Tanzrund, Startklasse, Turniernr As String
    sqlstr = "select * from Rundentab rt where rt.RT_ID=" & Runde_suchen
    Set rs = dbs.OpenRecordset(sqlstr)
    Tanzrund = rs!Runde
    Startklasse = rs!Startklasse
    Turniernr = rs!Turniernr
    rs.Close
    sqlstr = "select * from Paare_Rundenqualifikation where RT_ID= " & Runde_suchen
    Set rstauslosung = dbs.OpenRecordset(sqlstr)
    
    ' Abbruch, wenn keine Daten vorhanden sind
    If (rstauslosung.EOF) Then
        rstauslosung.Close
        Exit Sub
    End If
    
    rstauslosung.MoveLast
    was = 1
    trunde = 1
    Anzahl = rstauslosung.RecordCount
    rstauslosung.MoveFirst
    Do While Not rstauslosung.EOF()
        zufall = Int(Anzahl * Rnd + (rstauslosung!Anwesend_Status - 1) * (1000)) ' Zufallszahlen generieren.
        rstauslosung.Edit
        rstauslosung!Auslosung = zufall
        rstauslosung.Update
        rstauslosung.MoveNext
    Loop
    rstauslosung.Close
    
    sqlstr = "select * from Paare_Rundenqualifikation where RT_ID= " & Runde_suchen & " order by auslosung"
    Set rstauslosung = dbs.OpenRecordset(sqlstr)
    
    was = 1
    trunde = 1
    rstauslosung.MoveFirst
    
    Do While Not rstauslosung.EOF()
        rstauslosung.Edit
        If was > Anz_Paare Then
            trunde = trunde + 1
            was = 1
        End If
        was = was + 1
'        If (rstauslosung!Auslosung >= 1000) Then
'            rstauslosung!Rundennummer = Null
'        Else
            rstauslosung!Rundennummer = trunde
'        End If
        
        rstauslosung.Update
        rstauslosung.MoveNext
    Loop
    rstauslosung.Close
    '  Anfang
    '  verhindern, dass mehrere Paare aus dem gleichen Verein in der gleichen Runde tanzen
    '
    Call Rundenauslosung(Runde_suchen, Anz_Paare)
    ' getrennte Auslosung ?
    If Not DLookup("Getrennte_Auslosung", "Turnier", "Turniernum = " & Turniernr) Then
    
        Runde_übertragen Tanzrund, Startklasse
    End If
BW_RR_Error:
    dbs.Close
    
    Me.Requery
    DoCmd.RepaintObject , ""
    DoCmd.GoToRecord , "", acFirst
    DoCmd.SetWarnings True
  
End Sub

Private Sub Runde_übertragen(Tanzrund, Startklasse)
    Dim dbs As Database
    Dim rst As Recordset
    Dim stmt As String
    Set dbs = CurrentDb
    ' Wenn Vor/Endrunde der BW-Hauptklasse oder BW-Oldieklasse dann die Rundeneinteilung in die schnelle und langsame übernehmen
    If (InStr(1, Tanzrund, "_r_schnell") And (Startklasse = "BW_MA" Or Startklasse = "BW_SA")) Then
        stmt = "Select * from Rundentab where Turniernr=" & Turniernr & " and Startklasse='" & Startklasse & "' and Runde='" & left(Tanzrund, 3) & "_r_lang'"
        Set rst = dbs.OpenRecordset(stmt)
        
        If rst.NoMatch Then
            MsgBox "Langsame Runde für Boogie-Woogie wurde nicht gefunden!"
        Else
            Call UpdateRundenqualifikation(Runde_suchen, rst!RT_ID, True)
        End If
    End If
    
    ' Wenn Vorrunde oder Endrunde der RR-A oder RR-B dann die Rundeneinteilung in die Fußtechnik und Akrobatik übernehmen
    If InStr(1, Tanzrund, "_Akro") > 0 And Me!Runde_suchen.Column(2) = False Then
        stmt = "Select * from Rundentab where Turniernr=" & Turniernr & " and Startklasse='" & Startklasse & "' and Runde = 'End_r_Fuß'"
        Set rst = dbs.OpenRecordset(stmt)
        
        If (rst.NoMatch) Then
            MsgBox "Fußtechnik Enrunde für RR wurde nicht gefunden!"
        Else
            Call UpdateRundenqualifikation(Runde_suchen, rst!RT_ID, True)
        End If
    End If
End Sub

