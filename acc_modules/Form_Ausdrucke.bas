Option Compare Database
Option Explicit
    Dim stDocName As String

Private Sub Befehl0_Click()
    DoCmd.Close
End Sub

Private Sub Befehl2_Click()
    stDocName = "Startliste_Runden"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub Befehl3_Click()
    stDocName = "Rundenpaarung-Erste-Runde"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub Befehl108_Click()
    Me!Runde_einstellen = Null
    Me!Startklasse_einstellen = Null
End Sub

Private Sub Befehl20_Click()
If Not Forms![Ausdrucke]![Runde_einstellen] = " " Then

    stDocName = "Ergebnisliste_Runden_tl"
    DoCmd.OpenReport stDocName, acPreview
Else
    MsgBox ("Bitte Runde auswählen")
End If

End Sub

Private Sub Befehl28_Click()
    stDocName = "Startliste_aller_Runden"
    DoCmd.OpenReport stDocName, acPreview
End Sub


Private Sub Befehl31_Click()
    If Not Forms![Ausdrucke]![Startklasse_einstellen] = " " Then
        stDocName = "Ergebnisliste_Klasse_komplett"
        DoCmd.OpenReport stDocName, acPreview
    Else
        MsgBox ("Bitte Startklasse auswählen")
    End If
End Sub

Private Sub Befehl39_Click()
    [Form_A-Programmübersicht]!Report_Turniernum = [Form_A-Programmübersicht]!Akt_Turnier
    
    stDocName = "Turnierbericht"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub Befehl4_Click()
    If Not Forms![Ausdrucke]![Runde_einstellen] = " " Then
        stDocName = "Startliste_Runden"
        DoCmd.OpenReport stDocName, acPreview
    Else
        MsgBox ("Bitte Runde auswählen")
    End If
End Sub


Private Sub Befehl41_Click()
    If Not Forms![Ausdrucke]![Runde_einstellen] = " " Then
        stDocName = "Startliste_startende_Paare"
        DoCmd.OpenReport stDocName, acPreview
    Else
        MsgBox ("Bitte Runde auswählen!")
    End If
End Sub

Private Sub Befehl42_Click()
    stDocName = "Ergebnisliste_fuer_Presse"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub Befehl43_Click()
    stDocName = "Ergebnisliste_komplett"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub Befehl44_Click()
    stDocName = "unentschuldigt_gefehlte_Paare"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub Befehl46_Click()
    stDocName = "Teilnahme_ohne_Buch"
    DoCmd.OpenReport stDocName, acPreview

End Sub

Private Sub Befehl5_Click()
    If (Me.Startklasse_einstellen = "" Or IsNull(Me.Startklasse_einstellen)) Then
        MsgBox "Bitte wählen Sie zuerst eine Startklasse aus!"
        Exit Sub
    End If

    stDocName = "Startliste"
    DoCmd.OpenReport stDocName, acPreview
End Sub
Private Sub Befehl13_Click()

    stDocName = "Ablaufplanung"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Sub Kombinationsfeld14_AfterUpdate()
    ' Den mit dem Steuerelement übereinstimmenden Datensatz suchen.
    Me.RecordsetClone.FindFirst "[ident] = " & Me![Kombinationsfeld14].Column(1)
    Me.Bookmark = Me.RecordsetClone.Bookmark
End Sub

Private Sub Befehl23_Click()
    stDocName = "Ergebnisliste_Runden_f"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub Befehl26_Click()
    If Not Forms![Ausdrucke]![Runde_einstellen] = " " Then
        stDocName = "Platzierungsliste"
        DoCmd.OpenReport stDocName, acPreview
    Else
        MsgBox ("Bitte Runde auswählen")
    End If
End Sub
Private Sub Befehl27_Click()
    ' Jetzt den Report öffnen
    If Not Forms![Ausdrucke]![Runde_einstellen] = " " Then
        [Form_A-Programmübersicht].Report_RT_ID = Runde_auswaehlen.Column(2)
        Call showReport_Platzierte_Paare
        
        stDocName = "Platzierungsliste_WR"
        DoCmd.OpenReport stDocName, acPreview
    Else
        MsgBox ("Bitte Runde auswählen")
    End If
End Sub

Private Sub Befehl52_Click()
    stDocName = "Startliste_aller_Runden_nach_Vereinen"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub Befehl53_Click()
    If (Me.Startklasse_einstellen = "" Or IsNull(Me.Startklasse_einstellen)) Then
        MsgBox "Bitte wählen Sie zuerst eine Startklasse aus!"
        Exit Sub
    End If

    stDocName = "Startliste_nach_Vereinen"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub Befehl54_Click()
    If left(Runde_einstellen, 3) = "End" Then
        DoCmd.OpenReport "Ergebnisliste_RR_LS", acPreview
    Else
        MsgBox ("Bitte Endrunde der BW-Hauptklasse, BW-Oldieklasse, RR_A oder RR_B auswählen.")
    End If
End Sub

Private Sub Befehl57_Click()
    If Not Forms![Ausdrucke]![Runde_einstellen] = " " Then
        stDocName = "Teamwertung"
        DoCmd.OpenReport stDocName, acPreview
    Else
        MsgBox ("Bitte Runde auswählen")
    End If
End Sub

Private Sub Befehl58_Click()
    stDocName = "Startliste_nach_Vereinen_alle_Klassen"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub Befehl72_Click()
    If Not Forms![Ausdrucke]![Runde_einstellen] = " " Then
        stDocName = "Ergebnisliste_Runden_OWR"
        DoCmd.OpenReport stDocName, acPreview
    Else
        MsgBox ("Bitte Runde auswählen")
    End If
End Sub

Public Function btn_Ausdr(nr)
    Dim doc As String
    Dim lae As String
    lae = Forms![A-Programmübersicht]!Turnierausw.Column(8)
    
    If Me("btn_Ausdrucke_" & nr).Caption = ". . ." Then Exit Function
    doc = DLookup(lae & "_Dokumentation", "Dokumente", "btn = 'btn_Ausdrucke_" & nr & "'")
    If InStr(doc, ".pdf") > 0 Then
        Call showDocument(doc)
    Else
        DoCmd.OpenReport doc, acViewPreview
    End If
End Function

Private Sub btn_Ausdrucke_213_Click()
    If Not Forms![Ausdrucke]![Runde_einstellen] = " " Then
        stDocName = "ObserverWertungsbogen"
        DoCmd.OpenReport stDocName, acPreview, , "Startkl = '" & Me.Startklasse_einstellen & "' AND RT_ID = " & Me.Runde_auswaehlen.Column(2) & ""
    Else
        MsgBox ("Bitte Runde auswählen")
    End If

End Sub

Private Sub Form_Open(Cancel As Integer)
    setzte_buttons Me.Name, Me.Name, Forms![A-Programmübersicht]!Turnierausw.Column(8)
End Sub

Private Sub btnAlterskontrolle_Click()
    stDocName = "Alterskontrolle"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub btnAnwesenheitsliste_Click()
    stDocName = "Anwesenheitsliste"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub btnBetreuerliste_Click()
    Dim dbs As Database
    Set dbs = CurrentDb
    Dim rst As Recordset
    Dim rstBL As Recordset
    Dim Verein As String
    Dim Anzahl As Integer
    Dim isTeam As Boolean
    Call dbs.Execute("DELETE FROM Betreuerliste")
    Dim stmt As String
    Dim i As Integer
    'stmt = "SELECT Verein_Name, Count(*) AS Anzahl_Paare FROM Paare WHERE Anwesent_Status>0 and Turniernr=" & [Form_A-Programmübersicht]![Akt_Turnier] & " GROUP BY Verein_Name"
    ' Neues Statement mit Formationen (isTeam-Spalte)
    stmt = "SELECT Paare.Verein_Name, Count(*) AS Anzahl, Startklasse.isTeam, Paare.Name_Team FROM Startklasse INNER JOIN Paare ON Startklasse.Startklasse = Paare.Startkl WHERE (((Paare.Anwesent_Status)>0) AND ((Paare.Turniernr)=" & [Form_A-Programmübersicht]![Akt_Turnier] & ")) GROUP BY Paare.Verein_Name, Startklasse.isTeam, Paare.Name_Team"
    Set rst = dbs.OpenRecordset(stmt)
    Set rstBL = dbs.OpenRecordset("Betreuerliste")
    Do While (Not rst.EOF)
        Anzahl = Int((rst!Anzahl + 4) / 5)
        isTeam = rst!isTeam
        If (isTeam) Then
            Anzahl = rst!Anzahl * 2
        End If
        
        For i = 1 To Anzahl
            rstBL.AddNew
            rstBL!BL_VEREIN = rst!Verein_Name
            If (isTeam = False) Then
                rstBL!BL_BETREUER = "Paare (" & i & ". Betreuer)"
                rstBL!BL_GRUPPE = "1_Paare"
            Else
                rstBL!BL_BETREUER = "" & rst!Name_Team & " (" & i & ". Betreuer)"
                rstBL!BL_GRUPPE = "2_" & rst!Name_Team
            End If
            rstBL!BL_Turniernr = [Form_A-Programmübersicht]![Akt_Turnier]
            rstBL.Update
        Next
        rst.MoveNext
    Loop
    rst.Close
    dbs.Close
    
    stDocName = "Betreuerliste_Einzelpaarturnier"
    DoCmd.OpenReport stDocName, acPreview

End Sub

Private Sub btnReisekostenabrechnung_Click()
    stDocName = "Reisekosten2"
    DoCmd.OpenReport stDocName, acPreview
    'Call showDocument("Formulare\Reisekostenabrechnung.pdf")
End Sub

Private Sub btnReisekostenabrechnung1_Click()
    stDocName = "Reisekosten1"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub btnReisekostenabrechnung2_Click()
    stDocName = "Reisekostenabrechnung"
    DoCmd.OpenForm stDocName, acNormal
End Sub

Private Sub btnUrkundendaten_Click()
    If Not Forms![Ausdrucke]![Runde_einstellen] = " " Then
        Dim sFilters As String
        sFilters = "Microsoft Excel-Dateien (*.xls)" & vbNullChar & "*.xls" & vbNullChar & vbNullChar
        
        Dim sFilepath As String
        sFilepath = FileSaveAs("Urkundendaten.xls", ".xls", sFilters)
        
        If Len(sFilepath) Then
            DoCmd.OutputTo acQuery, "ausgeschiedene_Paare_Urkunden", "MicrosoftExcel(*.xls)", sFilepath, False, ""
        End If
    Else
        MsgBox ("Bitte Runde auswählen!")
    End If
End Sub

Private Sub btnWertungsbogenBWEinzel_Click()
    stDocName = "WertungsbogenEinzelBW"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub btnWertungsbogenBWFormation_Click()
    stDocName = "WertungsbogenFormBW"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub btnWertungsbogenDUO_Click()
    stDocName = "WertungsbogenDUO"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub btn_Ausdrucke_42_Click()
'*****AB***** V13.02 neuer Funktionsaufruf zur Auswahl der Wertungsbögen
    Dim Auswahl_Wertungsbogenart As String
    
    Auswahl_Wertungsbogenart = InputBox("Bitte die Art des Wertungsbogens auswählen (AK, FT_V, FT_E)", "Auswahl Wertungsbogenart", "FT_V")
    If Auswahl_Wertungsbogenart <> "AK" And Auswahl_Wertungsbogenart <> "FT_E" And Auswahl_Wertungsbogenart <> "FT_V" Then
        MsgBox ("Sie haben keine gültige Art von Wertungsbogen eingegeben, bitte wiederholen Sie den Aufruf.")
    Else
        stDocName = "WertungsbogenEinzelRR_" & Auswahl_Wertungsbogenart
        DoCmd.OpenReport stDocName, acPreview
    End If
End Sub

Private Sub CD_Einleger_Click()
    stDocName = "CD-Einleger"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub Giveaway_Click()
    
    If Nz(Me![Startklasse_einstellen]) = "" Or Nz(Me!Runde_auswaehlen) = "" Then
        MsgBox ("Bitte Startklasse und Runde auswählen!")
    Else
        Print_Givaway Me!Runde_auswaehlen.Column(2), Me.Runde_auswaehlen
    End If

End Sub

Sub Print_Givaway(RundenTab_ID, Runde)
    Dim re As Recordset
    Dim fil As String
    Set re = DBEngine(0)(0).OpenRecordset("SELECT TP_ID FROM Majoritaet WHERE  RT_ID=" & RundenTab_ID & " And RT_ID Is Not Null AND Runde_Report=1;")
'*****AB***** V13.05 - falls es sich um eine Endrunde handelt andere Abfrage ohne Runde_Report
'*****HM 14.07 ***** - auf geteilte Endrunden erweitert
    If Runde = "Endrunde" Or Runde = "Endrunde Akrobatik" Or Runde = "Schnelle Endrunde" Then
        Set re = DBEngine(0)(0).OpenRecordset("SELECT TP_ID FROM Majoritaet WHERE  RT_ID=" & RundenTab_ID & " And RT_ID Is Not Null;")
    End If
    If re.RecordCount = 0 Then
        MsgBox "Es gibt für diese Runde keine platzierten Paare"
    Else
        re.MoveFirst
        Do Until re.EOF
            fil = fil & IIf(Len(fil) = 0, "TP_ID=", " OR TP_ID=") & re!TP_ID
            re.MoveNext
        Loop
        stDocName = "Giveaway"
        DoCmd.OpenReport stDocName, acPreview, , fil
    End If
End Sub

Private Sub Ranglistenexport_Click()
    Dim sFilepath As String
    If (IsNull(Forms![A-Programmübersicht]![Akt_Turnier]) Or (Forms![A-Programmübersicht]![Akt_Turnier] = 0)) Then
       MsgBox ("Bitte Turnier auswählen!")
       Exit Sub
    End If
    
    sFilepath = getBaseDir & "Rangliste " & Forms![A-Programmübersicht]![Turnierbez] & ".xls"
    
    If Len(sFilepath) Then
        DoCmd.OutputTo acQuery, "Ergebnisliste_Text", "MicrosoftExcel(*.xls)", sFilepath, False, ""
    End If

End Sub

Private Sub Runde_auswaehlen_AfterUpdate()
    Startklasse_einstellen = Runde_auswaehlen.Column(4)
    Runde_einstellen = Runde_auswaehlen.Column(3)
    RundenId = Runde_auswaehlen.Column(2)
    [Form_A-Programmübersicht].Report_RT_ID = Runde_auswaehlen.Column(2)
End Sub

Private Sub Befehl49_Click()
    If Not Forms![Ausdrucke]![Runde_einstellen] = " " Then
        stDocName = "Startliste_Runden_Zeit"
        DoCmd.OpenReport stDocName, acPreview
    Else
        MsgBox ("Bitte Runde auswählen")
    End If
End Sub

Private Sub Befehl51_Click()

    stDocName = "Wertungsrichter_Einteilung"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub Startklasse_einstellen_AfterUpdate()
    Dim source As String
    
    source = "SELECT Tanz_Runden.R_NAME_ABLAUF, Runden4Drucken.Startklasse_text, Runden4Drucken.RT_ID, Runden4Drucken.Runde, Runden4Drucken.Startklasse, Runden4Drucken.Turniernum, Runden4Drucken.Turnier_Name, Runden4Drucken.InRundeneinteilung, Startklasse.Reihenfolge, Tanz_Runden.Rundenreihenfolge"
    source = source & " FROM Tanz_Runden INNER JOIN (Startklasse INNER JOIN Runden4Drucken ON Startklasse.Startklasse = Runden4Drucken.Startklasse) ON Tanz_Runden.Runde = Runden4Drucken.Runde"
    source = source & " WHERE (((Runden4Drucken.InRundeneinteilung) > 0) and ((Runden.Startklasse)=[Startklasse_einstellen]))"
    source = source & " ORDER BY Startklasse.Reihenfolge, Tanz_Runden.Rundenreihenfolge;"

    Runde_auswaehlen.RowSource = source
    Runde_auswaehlen.Requery
    Runde_auswaehlen = Null
    Runde_einstellen = Null
End Sub

Private Sub Wertungsbögen_Startklasse_Click()
    Dim dbs As Database
    Dim re As Recordset
    Dim fil As String
    Dim sk As String
    Dim rde As String
    
    Set dbs = CurrentDb
    sk = Nz(Me![Startklasse_einstellen])
    rde = Nz(Me!Runde_auswaehlen.Column(3))
    
    If sk = "" Or Nz(Me!Runde_auswaehlen) = "" Then
        MsgBox ("Bitte Startklasse und Runde auswählen!")
    Else
        Set re = dbs.OpenRecordset("SELECT Startklasse_Wertungsrichter.WR_ID FROM Wert_Richter INNER JOIN Startklasse_Wertungsrichter ON Wert_Richter.WR_ID = Startklasse_Wertungsrichter.WR_ID WHERE Startklasse='" & Me![Startklasse_einstellen] & "' AND Turniernr=" & get_aktTNr & ";")
        re.MoveFirst
        Do Until re.EOF
            fil = fil & IIf(Len(fil) = 0, "wr_id=", " OR wr_id=") & re!WR_ID
            re.MoveNext
        Loop
        
        If (sk = "RR_A" Or sk = "RR_B") And InStr(1, rde, "_Akro") Then
            'print_wait_close "Wertungsbogen", acViewPreview, "rt_ID =" & Me!Runde_auswaehlen.Column(2) & " AND (" & fil & ")"
            'Set re = dbs.OpenRecordset("select * from Rundentab where startklasse = '" & sk & "' and turniernr = " & get_aktTNr & " and runde = '" & Left(Me!nächste_Runde, 3) & "_r_Fuß'")
            'print_wait_close "Wertungsbogen", acViewPreview, "rt_ID =" & Me!Runde_auswaehlen.Column(2) & " AND (" & fil & ")"
        End If
     '*****AB***** V13.02 If-Clause um neue Startklassen ergänzt
     '*****AB***** V13.04 MB und SB wieder entfernt
        If (sk = "BW_H" Or sk = "BW_O" Or sk = "BW_MA" Or sk = "BW_SA") And rde = "End_r" Then
            Set re = dbs.OpenRecordset("select * from Rundentab where startklasse = '" & sk & "' and turniernr = " & get_aktTNr & " and runde='End_r_lang'")
            If re.RecordCount > 0 Then print_wait_close "Wertungsbogen", acViewPreview, "rt_ID =" & re!RT_ID & " AND (" & fil & ")"
                            
            Set re = dbs.OpenRecordset("select * from Rundentab where startklasse = '" & sk & "' and turniernr = " & get_aktTNr & " and runde='End_r_schnell'")
            If re.RecordCount > 0 Then print_wait_close "Wertungsbogen", acViewPreview, "rt_ID =" & re!RT_ID & " AND (" & fil & ")"

            re.Close
        Else
            DoCmd.OpenReport "Wertungsbogen", acViewPreview, , "rt_ID =" & Me!Runde_auswaehlen.Column(2) & " AND (" & fil & ")"
        End If

    
    End If

End Sub
