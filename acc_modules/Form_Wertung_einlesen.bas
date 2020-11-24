Option Compare Database
Option Explicit
    Dim db As Database
    Dim ausw As Recordset
    Dim aktuelleTanzRunde As Long

Private Sub AutomatischWertungenEinlesen_Click()
    If Me.AutomatischWertungenEinlesen = True Then
        Me.AutomatischWertungenEinlesen.Caption = "STOP"
    ElseIf Me.AutomatischWertungenEinlesen = False Then
        Me.AutomatischWertungenEinlesen.Caption = "START"
    End If
    
End Sub

Private Sub Befehl27_Click()
    DoCmd.Close

End Sub

Private Sub bereich_msg_AfterUpdate()
    Me!sende_text = Me!bereich_msg.Column(1)
End Sub

Private Sub Form_Load()
    Form_Resize
    If get_properties("EWS") = "EWS3" Then
        Me!Runde_starten.Visible = True
'        Me!nochmal_starten.Visible = True
   Else
        Me!Runde_starten.Visible = False
'        Me!nochmal_starten.Visible = False
    End If
    Select Case Forms![A-Programmübersicht]!Turnierausw.Column(8)
        Case "SL"
            'Me!Wertung_drucken.Visible = False
        Case "D"
            Me!Wertung_drucken.Visible = False
            
        Case Else
    End Select

End Sub

Private Sub Form_Resize()
    Me!Linie137.Width = Me.InsideWidth - 2
End Sub

Private Sub Form_Timer()

'****AB**** V13_04 - automatisches Einlesen der abgegebenen Wertungen, diese Funktion wird alle 5 Sekunden aufgerufen
'****AB**** V13_05 - erweitert um die Abfrage ob der Button AutomatischWertungenEinlesen gedrückt ist

If Not IsNull(Me.Tanzrunde) And Me.AutomatischWertungenEinlesen = True Then
    'MsgBox ("Aktualisierung")
    Wertungen_einlesen_Click
End If


End Sub

Private Sub nochmal_starten_Click()
    Dim st
    Dim back
    If Me!nochmal = True Then
        back = MsgBox("Das Paar startet schon nocheinmal!" & vbCrLf & vbCrLf & "Wirklich nochmal starten?", vbYesNo)
    Else
        back = MsgBox("Nocheinmal starten?", vbYesNo)
    End If
    If back = vbNo Then
        Exit Sub
    Else
        st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=nochmal_starten&text=" & TP_ID)
        If st = "eingetragen" Then
            Me!nochmal = True
            DoCmd.Requery
        Else
            MsgBox "Die Wiederholung wurde nicht eingetragen"
        End If
    End If

End Sub

Private Sub Platzierung_freigeben_Click()
    Dim db As Database
    Dim re As Recordset
    Dim t As Integer
    Dim fName, fPfad As String
    Set db = CurrentDb
       
    fPfad = getBaseDir & "Apache2\htdocs\"
    fName = Dir(fPfad & "T" & Forms![A-Programmübersicht]!Turnier_Nummer & "R*" & "_K" & Me!Tanzrunde & "_2000.html")
    
    Do Until fName = ""
        FileCopy fPfad & fName, fPfad & Replace(fName, "_2000", "_1000")
        Kill fPfad & fName
        fName = Dir
    Loop
End Sub

Private Sub Runde_auswerten_Click()
    DoCmd.OpenForm "Majoritaet_ausrechnen"
    Forms!Majoritaet_ausrechnen!Startklasse = Me!Tanzrunde
    DoCmd.Close acForm, "Wertung_einlesen"
End Sub

Private Sub Form_AfterUpdate()
    Form_Paare_ohne_Punkte_UF.Requery
End Sub

Private Sub Umschaltfläche147_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim st As String    'Beitensport Taktung
    If Me!Umschaltfläche147.Caption = "Runde starten" Then
        st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=Runde_starten&text=")
        Me!Umschaltfläche147.Caption = "Runde auswerten"
    Else
        st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=Runde_auswerten&text=")
        Me!Umschaltfläche147.Caption = "Runde starten"
    End If
    Debug.Print st
End Sub

Public Sub Runde_starten_Click()
    Dim re, target As Recordset
    Dim st As String
    Dim retl As Integer
    Dim rmax, PaareProRunde As Integer
    Dim rundeninfo As String
    Dim SngSec As Long
    If no_runde_selected Then Exit Sub
    
    Set db = CurrentDb
' nachschauen ob ausgelost
' select count(Rundennummer) As anz FROM Paare_Rundenqualifikation where RT_ID = 28 and Rundennummer > 0;
    Me!Umschaltfläche147.Caption = "Runde starten"
    Set re = db.OpenRecordset("SELECT s.Startklasse_text, t.Rundentext, r.* FROM (rundentab r INNER JOIN Startklasse s ON r.Startklasse = s.Startklasse) INNER JOIN Tanz_Runden_fix t ON r.Runde =t.Runde WHERE (r.gestartet=True AND r.getanzt=False);")
    
    If re.RecordCount > 0 Then
        If re!RT_ID = Me!RT_ID Then
            retl = MsgBox(re!Rundentext & " in der " & re!Startklasse_text & " läuft bereits!" & vbCrLf & "Wirklich nochmal starten?", vbYesNo + vbCritical + vbDefaultButton2)
        Else
            retl = MsgBox("Es läuft gerade die " & re!Rundentext & " in der " & re!Startklasse_text & " Klasse!" & _
                    vbCrLf & vbCrLf & "Soll die " & Me!Feld138 & " wirklich gestartet werden!", vbYesNo + vbCritical)
            re.Edit
            re!getanzt = True
            re.Update
        End If
    Else
        Set re = db.OpenRecordset("SELECT r.RT_ID, [gestartet] And [getanzt] AS Ausdr1 FROM rundentab AS r WHERE r.RT_ID=" & Me!RT_ID & ";")
        If re!Ausdr1 Then
            retl = MsgBox("Runde wurde bereits gewertet!" & vbCrLf & "Wirklich nochmal starten?", vbYesNo + vbCritical + vbDefaultButton2)
        Else
            retl = MsgBox("Runde starten?", vbYesNo)
        End If
    End If
    If retl = vbNo Then Exit Sub
    
    db.Execute "INSERT INTO Analyse (CGI_Input,zeit) VALUES ('" & Me!Tanzrunde.Column(1) & " gestartet', '" & Time & "')"
    db.Execute "UPDATE wert_richter Set WR_func='', WR_status='';"
    db.Execute "UPDATE Wert_Richter INNER JOIN Startklasse_Wertungsrichter ON Wert_Richter.WR_ID = Startklasse_Wertungsrichter.WR_ID SET WR_func = [WR_function], WR_status = 'start' WHERE Startklasse='" & Me!Tanzrunde.Column(3) & "';"
    db.Execute "UPDATE wert_richter Set WR_status='runde' WHERE WR_func='Ob';"
    db.Execute "UPDATE rundentab SET gestartet = true WHERE RT_ID=" & Me!RT_ID & ";"
    rundeninfo = RT_ID
                
    st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=observer_starten&text=" & rundeninfo & "&mdb=" & get_TerNr)
        SngSec = Timer + 1
        Do While Timer < SngSec
            DoEvents
        Loop
    st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=aufWRwartenweiter&text=")

End Sub

Private Sub Runde_beenden_Click()
    If no_runde_selected Then Exit Sub
    
    Dim re As Recordset
    Dim st As String
        
    Set db = CurrentDb
    Wertungen_einlesen_Click
    AuswertenundPlatzieren Me.Tanzrunde, Me.Tanzrunde.Column(3), Me.Tanzrunde.Column(17), Me.Tanzrunde.Column(6), Me.Tanzrunde.Column(7)
    If get_properties("EWS") = "EWS3" Then
        Set re = db.OpenRecordset("Select* from Rundentab Where RT_ID =" & Me!Tanzrunde & ";")
        If re!gestartet = True Then     'And re!getanzt = False Then
            db.Execute ("UPDATE rundentab SET [getanzt] = -1 WHERE RT_ID =" & Me!Tanzrunde & ";")
            st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=beamer_zeitplan&text=" & Me!Tanzrunde & "")
        Else
            'MsgBox "runde wurde noch nicht gestartet"
        End If
    End If
    db.Execute ("UPDATE rundentab SET [HTML] = 0 WHERE RT_ID =" & Me!Tanzrunde & ";")
    db.Execute "INSERT INTO Analyse (CGI_Input,zeit) VALUES ('" & Me!Tanzrunde.Column(1) & " beendet', '" & Time & "')"
    Start_Seite "T" & Forms![A-Programmübersicht]!Turnier_Nummer
    make_a_schedule
End Sub

Private Sub sende_msg_Click()
    Dim st As String
    st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=" & Me!bereich_msg & "&text=" & Me!sende_text)
'    st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=beamer&kopf=Vorrunde&inhalt=<table style=""width: 100%; float: left; padding-left:100px"" id=""table_timetable""><thead><tr role=""row""><th style=""width: 250px;"" colspan=""1"" rowspan=""1"" class=""sorting_disabled"">Beginn</th><th style=""width: auto;"" colspan=""1"" rowspan=""1"" class=""sorting_disabled"">Runde</th></tr></thead><tbody style=""font-size: 1.8vw;""> <tr class=""odd""> <td>19:00</td><td>Vorrunde  Juniorenklasse</td> </tr> <tr class=""odd""><td>19:10</td><td>Endrunde  Schülerklasse</td></tr>")
    
End Sub

Sub Tanzrunde_AfterUpdate()
    Dim dbs As Database
    Dim Turniernr As Integer
    Dim Startklasse_einstellen As String
    Dim sqlstr As String
    Dim where_part As String
    Dim re As Recordset
    Dim AnzahlWRVorgabe, t As Integer
    If Not IsNull(Tanzrunde) Then
        Me!Wertungen_einlesen.ControlTipText = Tanzrunde
        sqlstr = "SELECT Paare_Rundenqualifikation.RT_ID, Paare.Startkl, Paare_Rundenqualifikation.Rundennummer, Paare.Startnr, Paare_Rundenqualifikation.PR_ID, Paare_Rundenqualifikation.nochmal FROM (Paare INNER JOIN Paare_Rundenqualifikation ON Paare.TP_ID = Paare_Rundenqualifikation.TP_ID) WHERE (Paare_Rundenqualifikation.RT_ID= " & Me!Tanzrunde & " AND Paare_Rundenqualifikation.Anwesend_Status=1) ORDER BY Paare_Rundenqualifikation.Rundennummer, Paare.Startnr;"
        Set dbs = CurrentDb
        Set re = dbs.OpenRecordset(sqlstr)
        If re!Rundennummer > 0 Then
            Me.RecordSource = sqlstr
            ' bei Fuß nur FT-Wr
            '*****AB***** V13.02 Fehler es wurde noch auf das alte Feld WR_func im Recordset zugegriffen - hier geänder in: WR_function
            If Right(Me!Tanzrunde.Column(6), 4) = "_Fuß" Then
                where_part = "(Rundentab.RT_ID=" & Me!Tanzrunde & " AND Wert_Richter.Turniernr=" & get_aktTNr & " AND WR_function<>'Ak')"
            Else
                If left(Me!Tanzrunde.Column(6), 3) = "MK_" Then
                    where_part = "(Rundentab.RT_ID=" & Me!Tanzrunde & " AND Wert_Richter.Turniernr=" & get_aktTNr & " AND (Left([WR_function],1)='M' OR WR_function='Ob'))"
                Else
                    where_part = "(Rundentab.RT_ID=" & Me!Tanzrunde & " AND Wert_Richter.Turniernr=" & get_aktTNr & " AND (Left([WR_function],1)<>'M' OR WR_function='Ob'))"
                End If
            End If
            Set re = dbs.OpenRecordset("SELECT [WR_Nachname] & "" "" & [WR_Vorname] AS Ausdr1, Wert_Richter.WR_ID, Wert_Richter.WR_Kuerzel, Startklasse_Wertungsrichter.WR_function, Startklasse_Wertungsrichter.Startklasse, Rundentab.RT_ID FROM Wert_Richter INNER JOIN (Rundentab INNER JOIN Startklasse_Wertungsrichter ON Rundentab.Startklasse = Startklasse_Wertungsrichter.Startklasse) ON Wert_Richter.WR_ID = Startklasse_Wertungsrichter.WR_ID WHERE " & where_part & " ORDER BY Wert_Richter.WR_Kuerzel;")
            Set ausw = dbs.OpenRecordset("Auswertung", DB_OPEN_DYNASET)
            If re.RecordCount = 0 Then
                MsgBox "Es gibt keine eingteilten WR!"
                Exit Sub
            End If
            
            re.MoveFirst
            Me!Feld138.SetFocus
            Me("WR_9") = ""
            For t = 1 To 9
                Me("Pu" & t).Visible = False
                Me("Pl" & t).Visible = False
                Me("PuFe" & t).Visible = False
                Me("PlFe" & t).Visible = False
                Me("Tr" & t).Visible = False
                Me("WR_" & t).Visible = False
                Me("Feld" & t).Visible = False
                Me("Feld" & t).BackStyle = 0
                Me!Feld13.Visible = False
            Next
            t = 1
            Me!Startnr.Visible = True
            Me("Tr0").Visible = True
            Do Until re.EOF
                If re!WR_function = "Ob" Then
                    abzug_anzeige re!WR_ID, re!Ausdr1
                Else
                    Me("Pu" & t).Visible = True
                    Me("Pl" & t).Visible = True
                    Me("PuFe" & t).Visible = True
                    Me("PlFe" & t).Visible = True
                    Me("Tr" & t).Visible = True
                    Me("WR_" & t) = re!WR_ID
                    Me("Feld" & t).Caption = re!Ausdr1
                    Me("Feld" & t).Visible = True
                    t = t + 1
                End If
                Me!Feld14.Visible = True
                Me!Feld13.Visible = True
                re.MoveNext
            Loop
        '    If Left(Me!Tanzrunde.Column(3), 3) = "RR_" Then abzug_anzeige 99, "Beobachter"
            If (Not [Form_A-Programmübersicht]!Getrennte_Auslosung) Then
             '*****AB***** V13.02 if-Clause um neue Boogie Startklassen ergänzt
             '*****AB***** V13.04 BW_SB und BW_MB in Case wieder entfernt, da nur eine Endrunde getanzt wird
                If (Startklasse_einstellen = "BW_H" Or Startklasse_einstellen = "BW_O" Or Startklasse_einstellen = "BW_MA" Or Startklasse_einstellen = "BW_SA") And ([Forms]![Wertung_einlesen]!Tanzrunde.Column(7) = "End_r_lang" Or [Forms]![Wertung_einlesen]!Tanzrunde.Column(7) = "End_r_schnell") Then
                    ' Update der Rundeneinteilung
                    Dim rt_id_endr As Integer
                    rt_id_endr = getRT_ID(Turniernr, Startklasse_einstellen, "End_r")
                    Call UpdateRundenqualifikation(rt_id_endr, Tanzrunde, True)
                End If
            End If
            If (left(Me!Tanzrunde.Column(3), 3) = "BS_" Or left(Me!Tanzrunde.Column(5), 3) = "MK_") And get_properties("EWS") = "EWS3" Then
                Me!Umschaltfläche147.Visible = True
                Me!Umschaltfläche147.Caption = "Runde starten"
            Else
                Me!Umschaltfläche147.Visible = False
            End If

        Else
            Me.RecordSource = "SELECT * FROM Paare_Rundenqualifikation WHERE RT_ID = 0;"
            MsgBox "Es gibt noch keine Rundeneinteilung!"
        End If
    End If
End Sub

Private Sub abzug_anzeige(WR_ID, Ausdr1)
    Me("Pu" & 9).Visible = True
    Me("PuFe" & 9).Visible = True
    Me("Tr" & 8).Visible = True
    Me("Tr" & 9).Visible = True
    If Me("WR_" & 9) <> "" Then
        Me("WR_" & 9) = Me("WR_" & 9) & " / " & WR_ID
    Else
        Me("WR_" & 9) = WR_ID
    End If
    Me("Feld" & 9).Caption = "Observer"
    Me("Feld" & 9).Visible = True
End Sub

Private Sub Wertung_drucken_Click()
    Dim fil As String
    Dim t As Integer
    If IsNull(Me!Tanzrunde) Then
        MsgBox "Bitte Tanzrunde einstellen!"
    Else
        fil = "wr_id=" & Me("WR_1")
        For t = 2 To 9
            If (Me("Feld" & t).Visible = True) Then
                fil = fil & " OR wr_id=" & Replace(Me("WR_" & t), " / ", " OR wr_id=")
            End If
        Next
        DoCmd.OpenReport "Wertungsbogen", acViewPreview, , "rt_ID =" & Me!Tanzrunde & " AND (" & fil & ")"
    End If
End Sub

Private Sub Wertungen_einlesen_Click()
    Dim t As Integer
    Dim db As Database
    Dim wr As Recordset
    Dim gPlatz As String
    Dim fWertu As String
    Dim wrNam As String
    Dim retl As String
    
    If IsNull(Me!Tanzrunde) Then
        MsgBox "Bitte Tanzrunde einstellen!"
    Else
        Set db = CurrentDb
        If get_wertungen(Me!Tanzrunde, Me!Tanzrunde.Column(3), Me!Tanzrunde.Column(6)) = True Then
            'MsgBox "Für diese Runde existiert (noch) kein Datenfile!"
            Me.Status_Wertungen_Einlesen.Visible = True
        Else
            Me.Status_Wertungen_Einlesen.Visible = False
            For t = 1 To 8
                If Me("Feld" & t).Visible Then
                    retl = Wertung_check(Me("WR_" & t), t)  ' rückgabe ob nix, wertung fehlt, oder doppelte Plätze
                    If retl = "p" Then gPlatz = gPlatz & vbCrLf & fetch_wr_name(Forms!Wertung_einlesen("WR_" & t))
                    If retl = "w" Then fWertu = fWertu & vbCrLf & fetch_wr_name(Forms!Wertung_einlesen("WR_" & t))
                End If
            Next
            Set wr = db.OpenRecordset("SELECT * FROM wert_richter WHERE WR_AzuBi=True;")
            If Not wr.EOF Then wr.MoveFirst
            Do Until wr.EOF
                retl = Wertung_check(wr!WR_ID, 0)       ' rückgabe ob nix, wertung fehlt, oder doppelte Plätze
                If retl = "p" Then gPlatz = gPlatz & vbCrLf & fetch_wr_name(wr!WR_ID)
                If retl = "w" Then fWertu = fWertu & vbCrLf & fetch_wr_name(wr!WR_ID)
                wr.MoveNext
            Loop
'            If gPlatz <> "" Then MsgBox "Bei " & gPlatz & vbCrLf & "wurden Plätze mehrfach vergeben. Gleiche Platzvergabe in der Endrunde ist unzulässig!"
            
            '*****AB***** V13.05 - zusätzlich Abfrage ob automatisch Einlesen angeklickt ist, dann keine MsgBox für fehlende Wertungen!
            '*****AB***** V13.05 - automatisches Einlesen beenden sobald alle Wertungen da sind
            If fWertu <> "" And Me.AutomatischWertungenEinlesen = False Then MsgBox "Bei " & fWertu & vbCrLf & "fehlen noch Wertungen!"
            If fWertu = "" And Me.AutomatischWertungenEinlesen = True Then
                Me.AutomatischWertungenEinlesen = False
                Me.AutomatischWertungenEinlesen.Caption = "START"
            End If
            
            '*****AB***** V13.02 - zusätzlich die Wertungen für das Observer Plugin bereitstellen
            '*****AB***** KRITISCH - wenn fehlerhaft, einfach nächste Zeile auskommentieren!!!
            Import_RT_txt Me.Tanzrunde
        End If
        Set ausw = db.OpenRecordset("Auswertung", DB_OPEN_DYNASET)
        '****AB**** V13_04 HTML Seite für den Observer bereitstellen
        ObserverHTML (Me!Tanzrunde.Column(6))
    End If
    Requery
End Sub

Function fetch_wr_name(WR_ID)
    Dim db As Database
    Dim wr As Recordset
    Set db = CurrentDb
    Set wr = db.OpenRecordset("SELECT * FROM wert_richter WHERE WR_ID = " & WR_ID)
    fetch_wr_name = wr!WR_Vorname & " " & wr!WR_Nachname
    wr.Close
    db.Close
End Function

Private Sub Plazierung_einlesen_Click()
    Dim db As Database
    Dim re As Recordset
    Dim t As Integer
    Dim such As Integer

    If IsNull(Me!Tanzrunde) Then
        MsgBox "Bitte Tanzrunde einstellen!"
    Else
        t = get_platzierung(Me!Tanzrunde)
        If t = 2 Then
            MsgBox "Es existiert noch keine Platzierung!"
        Else
            For t = 1 To 8
                If Me("Feld" & t).Visible Then
                    Set db = CurrentDb
                    Set re = db.OpenRecordset("SELECT * from Auswertung a where a.wr_id=" & Forms!Wertung_einlesen("WR_" & t) & " and exists (select 1 from Paare_Rundenqualifikation pr where pr.pr_id=a.pr_id and pr.rt_id=" & Tanzrunde & ") order by a.platz asc")
                    're.Sort = "Platz"
                    If Not re.EOF Then
                        re.MoveFirst
                        Do Until re.EOF
                            such = re!Platz
                            re.MoveNext
                            If Not re.EOF Then
                                If such = re!Platz Then
                                    MsgBox "Die Platzierung wurde nicht richtig erfasst!"
                                    Exit Sub
                                End If
                            End If
                        Loop
                        Me("Feld" & t).BackStyle = 1
                        Me("Feld" & t).BackColor = 65280
                        Me("Feld" & t).ForeColor = 0
                    End If
                End If
            Next
        End If
    End If
    Me.Requery
End Sub

Private Sub Rundenmonitor_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Monitor_Runden"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

End Sub

Public Function Wertung_check(WR_ID, spalte)
    Dim dbs As Database
    Dim rstauswertung As Recordset          ', rstweiter, rstanzahl
    Dim stmt As String
    Dim IsEndrunde As Boolean
    Dim anzahl_p As Integer
    Dim werund, tr As String
    Dim mehrfach As Variant
    Dim Turniernr As Integer
    
    Set dbs = CurrentDb
    ' Anzahl Paare für diese Runden in die Tabelle schreiben
    tr = Tanzrunde.Column(7)
    Turniernr = get_aktTNr
    IsEndrunde = (Tanzrunde.Column(13) = 1)
    
    ' Wertung überprüfen und Plätze vergeben
    Dim zpl As Double, zpu As Double, zpldup As Double
    zpl = 0
    zpu = 0
    Set rstauswertung = dbs.OpenRecordset("SELECT Count(*) AS anz FROM Paare_Rundenqualifikation WHERE RT_ID=" & Tanzrunde & "and anwesend_Status=1;")
    anzahl_p = rstauswertung!anz
    ReDim mehrfach(anzahl_p)
    
    ' Recordset-Objekt vom Typ Dynaset erstellen. Tabelle Auswertung öffnen
    stmt = "SELECT count(*) as anz from Auswertung a, Paare_Rundenqualifikation pr"
    stmt = stmt & " where a.wr_id=" & WR_ID & " and pr.pr_id=a.pr_id and pr.rt_id=" & Tanzrunde
    stmt = stmt & " and Punkte is null"
    Set rstauswertung = dbs.OpenRecordset(stmt)
    Dim Count As Integer
    Count = rstauswertung!anz
    rstauswertung.Close
    If (Count > 0) Then
        Me("Feld" & spalte).BackColor = 255
        Me("Feld" & spalte).BackStyle = 1
        Exit Function
    End If
    
    stmt = "SELECT * from Auswertung a"
    stmt = stmt & " where a.wr_id=" & WR_ID & " and exists (select 1 from Paare_Rundenqualifikation pr where pr.pr_id=a.pr_id and pr.rt_id=" & Tanzrunde & ")"
    stmt = stmt & " order by a.punkte desc, a.platz asc"
    
    Set rstauswertung = dbs.OpenRecordset(stmt)
    If rstauswertung.EOF() Then
        Exit Function
    End If
    With rstauswertung
        .MoveFirst
        If (IsEndrunde) Then
            If !Platz = 0 Then   ' keine Platzvergabe für die Endrunde, wenn schon ein Platz vergeben wurde
                .Edit
                !Platz = 1
                .Update
            Else
                zpl = !Platz
            End If
         Else
            .Edit
            !Platz = 1
            .Update
        End If
        zpl = !Platz
        zpu = !Punkte
        '
        zpldup = 1  ' erster Platz wurde fest einmal vergeben
        .MoveNext
        Do While Not .EOF()
          
          If (IsEndrunde) And !Platz <> 0 Then
            zpl = !Platz
            zpu = !Punkte
          Else
            .Edit
            If !Punkte < zpu Then
                zpl = zpl + zpldup ' nächster zu vergebender Platz
                !Platz = zpl       ' diesen Platz vergeben
                zpldup = 1         ' Platz ist einmal vergeben
                zpu = !Punkte      ' bei diesem Punktestand
            Else
                If !Punkte = zpu Then  ' Platz mehrfach
                    !Platz = zpl         ' nach wie vor diesen Platz
                    zpldup = zpldup + 1  ' aber jetzt einmal mehr
                    mehrfach(0) = 1
                    mehrfach(zpl) = zpldup
                Else
                    If !Punkte > zpu Then
                        MsgBox ("Hier stimmt was nicht mit der Platzvergabe")
                        End
                    End If
                End If
            End If
            .Update
          End If
         .MoveNext
        Loop
    End With
    If (IsEndrunde) And left(Me!Tanzrunde.Column(3), 3) <> "RR_" And left(Me!Tanzrunde.Column(3), 3) <> "F_R" Then
        rstauswertung.MoveFirst
'        If mehrfach(0) = 1 And (Me!Tanzrunde.Column(6) <> "End_r_Fuß") Then
'            Call pg_platzieren(Tanzrunde, rstauswertung!WR_ID, mehrfach, rstauswertung.RecordCount, Me!Tanzrunde.Column(3))
'            'End
'        Else
            Call no_plazieren(Tanzrunde, rstauswertung!WR_ID, mehrfach, rstauswertung.RecordCount, Me!Tanzrunde.Column(3))
'        End If
    End If
    
    stmt = "SELECT Count(*) AS anz from Auswertung a"
    stmt = stmt & " where a.wr_id=" & WR_ID & "  AND ((IsNull([Cgi_Input]))=False) AND exists (select 1 from Paare_Rundenqualifikation pr where pr.pr_id=a.pr_id and pr.rt_id=" & Tanzrunde & ")"
    Set rstauswertung = dbs.OpenRecordset(stmt)
    
    If rstauswertung!anz <> anzahl_p Then
        Me("Feld" & spalte).BackColor = 255
        Me("Feld" & spalte).BackStyle = 1
        Me("Feld" & spalte).ForeColor = 16777215
        Wertung_check = "w"
        'MsgBox "Bei " & Forms!Wertung_einlesen("Feld" & spalte).Caption & " fehlen Wertungen"
'    ElseIf mehrfach(0) = 1 And (IsEndrunde) And Left(Me!Tanzrunde.Column(3), 3) <> "RR_" And Left(Me!Tanzrunde.Column(3), 3) <> "F_R" Then
'        Me("Feld" & spalte).BackColor = 255
'        Me("Feld" & spalte).BackStyle = 1
'        Me("Feld" & spalte).ForeColor = 16777215
'        Wertung_check = "p"
'        'MsgBox ("Bei " & DLookup("WR_nachNAME", "wert_richter", "WR_ID = " & WR_ID) & " gleiche Platzvergabe in der Endrunde ist unzulässig. Es wurden Plätze mehrfach vergeben!")
    Else
        Me("Feld" & spalte).BackStyle = 1
        Me("Feld" & spalte).BackColor = 65280
        Me("Feld" & spalte).ForeColor = 0
        Requery
    End If
End Function

Function Get_Pu(WR_ID, PR_ID)
    Dim vars
    Dim i As Integer
    'Get_Pu = "0,00"
    vars = Split(WR_ID, " / ")
    For i = 0 To UBound(vars)
        ausw.FindFirst "WR_ID=" & vars(i) & " AND PR_ID = " & PR_ID
        'Get_Pu = DLookup("Punkte", "Auswertung", "WR_ID=" & WR_ID & " AND PR_ID = " & PR_ID)
        If Not ausw.NoMatch Then Get_Pu = Format(ausw!Punkte, "###0.00")
    Next
End Function

Function Get_Pl(WR_ID, PR_ID)
    ausw.FindFirst "WR_ID=" & WR_ID & " AND PR_ID = " & PR_ID
    If Not ausw.NoMatch Then Get_Pl = ausw!Platz
End Function

Public Function show_wertung(PR_ID, Startnr, WR_ID)
    Dim db As Database
    Dim re, shw As Recordset
    Dim cgivar, zl
    Dim i As Integer
    
    Set db = CurrentDb
    
    Set re = db.OpenRecordset("SELECT * FROM Auswertung WHERE pr_id =" & PR_ID & " AND wr_id =" & WR_ID & ";")
    db.Execute ("DELETE * from Show")
    Set shw = db.OpenRecordset("Show", DB_OPEN_DYNASET)
    If re.RecordCount > 0 Then
        If Not IsNull(re!Cgi_Input) Then
            If CurrentProject.AllForms("Wertung_zeigen").IsLoaded Then
                DoCmd.Close acForm, "Wertung_zeigen"
            End If
            cgivar = Split(re!Cgi_Input, "&")
            
            For i = 0 To UBound(cgivar)
                zl = Split(cgivar(i), "=")
                shw.AddNew
                shw!SH_Name = zl(0)
                shw!SH_Wert = zl(1)
                shw!SH_sort = Right(zl(0), 2)
                shw.Update
            Next
            DoCmd.OpenForm "Wertung_zeigen"
            Forms!Wertung_zeigen!Text2 = "PR_ID: " & PR_ID & "  StNr: " & Startnr
        Else
            MsgBox "Wertung wurde manuell eingegeben!"
        End If
    End If
End Function

Private Sub wertungen_löschen_Click()
'    Überflüssiges_löschen
End Sub

Private Sub Zeitplan_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If no_runde_selected Then Exit Sub
    
    Forms!Wertung_einlesen!HTML_Select = 1
    get_url_to_string_check ("http://" & GetIpAddrTable() & "/hand?msg=beamer_zeitplan&text=" & Tanzrunde)
    Beamer_generieren
End Sub

Private Sub Runde_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If no_runde_selected Then Exit Sub
    
    Forms!Wertung_einlesen!HTML_Select = 2
    get_url_to_string_check ("http://" & GetIpAddrTable() & "/hand?msg=beamer_runde&text=")
    Beamer_generieren
End Sub

Private Sub Platzierungsliste_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If no_runde_selected Then Exit Sub
    
    Forms!Wertung_einlesen!HTML_Select = 3
    AuswertenundPlatzieren Me.Tanzrunde, Me.Tanzrunde.Column(3), Me.Tanzrunde.Column(17), Me.Tanzrunde.Column(6), Me.Tanzrunde.Column(7)
    get_url_to_string_check ("http://" & GetIpAddrTable() & "/hand?msg=beamer_ranking&text=")
    Beamer_generieren
End Sub

Private Sub Zeitplan_ganz_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If no_runde_selected Then Exit Sub
    
    Forms!Wertung_einlesen!HTML_Select = 4
    Beamer_generieren
End Sub

Private Sub Rundenergebnis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If no_runde_selected Then Exit Sub
    
    Forms!Wertung_einlesen!HTML_Select = 5
    AuswertenundPlatzieren Me.Tanzrunde, Me.Tanzrunde.Column(3), Me.Tanzrunde.Column(17), Me.Tanzrunde.Column(6), Me.Tanzrunde.Column(7)
    Beamer_generieren
End Sub

Private Sub Siegerehrung_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim st As String
    Dim Runde As String
    If no_runde_selected Then Exit Sub
    Runde = Me!Tanzrunde.Column(6)
    If Runde = "End_r_Akro" Or Runde = "End_r_schnell" Or Runde = "End_r" Or Runde = "End_r_2" Then
        If get_properties("EWS") = "EWS3" Then
            st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=beamer_siegerehrung&text=" & Tanzrunde & "&mdb=" & get_TerNr)
        Else
            Forms!Wertung_einlesen!HTML_Select = 6
            AuswertenundPlatzieren Me.Tanzrunde, Me.Tanzrunde.Column(3), Me.Tanzrunde.Column(17), Me.Tanzrunde.Column(6), Me.Tanzrunde.Column(7)
            Beamer_generieren
        End If
    Else
        MsgBox "Es gibt keine Siegerehrung für diese Runde!"
    End If
End Sub

Private Function no_runde_selected()
    If (IsNull(Forms!Wertung_einlesen!Tanzrunde) Or (Forms!Wertung_einlesen!Tanzrunde = 0)) Then
       MsgBox ("Bitte Tanzrunde einstellen!")
       no_runde_selected = True
    End If
End Function
