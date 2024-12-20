Option Compare Database
Option Explicit
    Dim stDocName As String

Private Sub Befehl0_Click()
    DoCmd.Close
End Sub
Private Sub Befehl18_Click()
    stDocName = "Paare_vorrunde_Anfügeabfrage"
    DoCmd.OpenQuery stDocName, acNormal, acEdit
End Sub

Private Sub Befehl19_Click()

    [Form_A-Programmübersicht]![Report_RT_ID] = Startklasse
    
    stDocName = "Ergebnisliste_Runden"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub Befehl20_Click()
    stDocName = "Monitor_Runden"
    DoCmd.OpenForm stDocName
    
End Sub

Private Sub btn_ausw_11_Click()
    [Form_A-Programmübersicht]![Report_RT_ID] = Startklasse
    
    stDocName = "Ergebnisliste_Runden_TL"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub btn_ausw_12_Click()
    [Form_A-Programmübersicht]![Report_RT_ID] = Startklasse
    
    stDocName = "Platzierungsliste"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub btn_ausw_13_Click()
    [Form_A-Programmübersicht]![Report_RT_ID] = nächste_Runde
    
    stDocName = "Startliste_startende_Paare"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub btn_ausw_14_Click()
    If (IsNull(Startklasse) Or Startklasse = "") Then
        MsgBox "Bitte wählen Sie erst eine Runde aus."
        Exit Sub
    End If
    '*****HM***** V13.05D Sperre für RR raus, RR-WR-Sperre bei showReport_Platzierte_Paare
    [Form_A-Programmübersicht]![Report_RT_ID] = Startklasse
    Call showReport_Platzierte_Paare

    stDocName = "Platzierungsliste_WR"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub btn_ausw_15_Click()
    Print_Givaway Me.Startklasse.Column(0), Me.Startklasse.Column(5)
End Sub

Private Sub btn_ausw_16_Click()
    Dim db As Database
    Dim re As Recordset
    Dim fil As String

    [Form_A-Programmübersicht]![Report_RT_ID] = Startklasse
    Set db = CurrentDb
    Set re = db.OpenRecordset("SELECT DISTINCT Paare.TP_ID FROM Paare WHERE (((Paare.RT_ID_Ausgeschieden)=" & Forms![A-Programmübersicht]!Report_RT_ID & " And (Paare.RT_ID_Ausgeschieden) Is Not Null));")
    If re.RecordCount > 0 Then
        re.MoveFirst
        Do Until re.EOF
            fil = fil & IIf(Len(fil) > 3, " OR TP_ID=", "TP_ID=") & re!TP_ID
        
            re.MoveNext
        Loop
        stDocName = "WR_Auswertung_NJS_TanzpaareFeedback"
        DoCmd.OpenReport stDocName, acPreview, , fil
    Else
        MsgBox "Zu dieser Runde gibt es keine platzierten Paare!"
    End If
End Sub

Private Sub btn_ausw_17_Click()
    If left(Me!Startklasse.Column(7), 5) = "End_r" Then
        gen_Ergebnisliste Me.RecordsetClone, Me!Startklasse.Column(4), Me!Startklasse.Column(4)
    Else
        MsgBox "Dies ist keine Endrunde"
    End If
End Sub

Private Sub btn_ausw_18_Click()
    If IsNull(Me!nächste_Runde.Column(1)) Then
        MsgBox "Keine weitere Runde gewählt"
    Else
        gen_NächsteRunde Me!Paare_Rundenqualifikation_Unterformular.Form.RecordsetClone, Me!nächste_Runde.Column(3), Me!nächste_Runde.Column(2), Me!nächste_Runde.Column(11)
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
    setzte_buttons "Majoritaet_ausrechnen", "ausw", Forms![A-Programmübersicht]!Turnierausw.Column(8)
    If get_properties("EWS") = "EWS3" Then
'        Me!btn_ausw_2.Visible = True
'        Me!btn_ausw_1.Visible = True
    End If
    If get_properties("sort_ablauf") Then
        Me!Startklasse.RowSource = "SELECT  * FROM Runden4AuswertenWeiternehmen ORDER BY Rundentab.Rundenreihenfolge;"
    End If
End Sub

Private Sub btnPaareWeiternehmen_Click()
    If (IsNull(Startklasse) Or Startklasse = "") Then
        MsgBox "Bitte wählen Sie erst eine Runde aus."
        Exit Sub
    End If

    stDocName = "Paare_weiternehmen"
    DoCmd.OpenForm stDocName, , , , , acDialog

    Me.Requery
End Sub

Private Sub btnMajoritaetLoeschen_Click()
    
    If (IsNull(Startklasse) Or Startklasse = "") Then
        MsgBox "Sie müssen erst eine Runde auswählen."
        Exit Sub
    End If
    
    Dim stmt As String
    stmt = "Delete from Majoritaet where RT_ID=" & Startklasse
    Dim dbs As Database
    Set dbs = CurrentDb
    dbs.Execute (stmt)
    
    Me.Requery
    
End Sub

Private Sub Anmerkung_Disqualifikation_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub DQ_ID_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub DQ_ID_AfterUpdate()
' HK 27.11.2011  Disqualifikation bei Eingabe in das Feld berechnen und nicht mehr
'                 über einen separaten Button
    SendKeys "+{ENTER}"
    majori_Click
Exit Sub
    
    
    If (IsNull(Startklasse) Or Startklasse = "") Then
        MsgBox "Sie müssen erst eine Runde auswählen."
        Exit Sub
    End If
    
    Dim dbs As Database
    Set dbs = CurrentDb
    Dim stmt As String
    Dim AnzahlFehler As Integer
    Dim strRecordSource
    Dim rst As Recordset
    Dim Platz As Integer
    Dim zrtid As Integer
    zrtid = RT_ID
    Dim ztpid As Integer
    ztpid = TP_ID
    Dim zdqid As Integer
    zdqid = DQ_ID
    AnzahlFehler = Kombinationsfeld104
    strRecordSource = Me.RecordSource
    Me.RecordSource = ""
    
    stmt = "Select * from majoritaet where rt_id=" & zrtid & " and tp_id=" & ztpid
    Set rst = dbs.OpenRecordset(stmt)
    If Not rst.EOF Then
        rst.Edit
        ' Die Disqualifikation einarbeiten
        rst!DQ_ID = zdqid
        rst.Update
        rst.Close
        Me.RecordSource = strRecordSource
        If left(Me!Startklasse.Column(3), 3) = "RR_" Then
            If Me!Startklasse.Column(7) = "KO_r" Then
                Call RR_KO_Sieger_ermitteln(zrtid)
                Call RR_platz_vergeben(zrtid, 0)
            Else
                Call RR_platz_vergeben(zrtid, 0)
            End If
        Else
            Call Kombinationsfeld104_AfterUpdate
        End If
    Else
        MsgBox ("Paar " & ztpid & " wurde in der Majoritätstabelle nicht gefunden")
    End If
    Me.RecordSource = strRecordSource
    
    Me.Requery
End Sub

Private Sub Kombinationsfeld104_AfterUpdate()
' HK 27.11.2011  Verstoß bei Eingabe in das Feld berechnen und nicht mehr
'                 über einen separaten Button
    
    
majori_Click
Exit Sub
    
    
    Dim strRecordSource
    Dim Runde As String
    Dim Turniernr As Integer
    Dim Startkl As String
    Dim AnzahlWR As Integer
    Dim ztpid As Integer
    Dim AnzahlFehler As Integer
    Dim RT_ID As Integer
    If (IsNull(Startklasse) Or Startklasse = "") Then
        MsgBox "Sie müssen erst eine Runde auswählen."
        Exit Sub
    End If
    
    RT_ID = Startklasse
    Turniernr = [Form_A-Programmübersicht]![Akt_Turnier]
    Runde = Startklasse.Column(7)
    Startkl = Startklasse.Column(3)
    AnzahlWR = Startklasse.Column(9)
    ztpid = TP_ID
    AnzahlFehler = Kombinationsfeld104
    strRecordSource = Me.RecordSource
    Me.RecordSource = ""
    
    Call RR_Punkteabzug(RT_ID, Startkl, ztpid, AnzahlFehler, Me.Startklasse.Column(7))
    
    If Startklasse.Column(8) = 1 Then 'falls Endrunde
        Call PaarePlatzieren(Startklasse, 1)
    End If
    Me.RecordSource = strRecordSource
    If Me!Startklasse.Column(7) = "End_r" Then
        make_a_siegerehrung Me!Startklasse          'HTML-Moderation
    End If

    Me.Requery
End Sub

Private Sub Kombinationsfeld104_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub majori_Click()
    If (IsNull(Startklasse) Or Startklasse = "") Then
        MsgBox "Sie müssen erst eine Runde auswählen."
        Exit Sub
    End If
    
    Me.Refresh
    
    '*****AB***** V14.02 Auswerten ausgelagert in externe Funktion, Parameter StartkalsseID, Startklassekurztext, WR-Anzahl, Rundenart, IsEndrunde
    AuswertenundPlatzieren Me.Startklasse, Me.Startklasse.Column(3), Me.Startklasse.Column(9), Me.Startklasse.Column(7), Me.Startklasse.Column(8)
    
    Me.Requery
End Sub

Private Function getBWRunde(Turniernr As Integer, Startklasse As String, Runde As String) As Integer
    Dim dbs As Database
    Set dbs = CurrentDb
    
    Dim rst As Recordset
    Set rst = dbs.OpenRecordset("Select * from Rundentab where Turniernr=" & Turniernr & " and Startklasse='" & Startklasse & "' and Runde='" & Runde & "'")
    If (rst.EOF) Then
        getBWRunde = -1
    Else
        getBWRunde = rst!RT_ID
    End If
    rst.Close
End Function

Private Sub Runde_AfterUpdate()
    DoCmd.RepaintObject acForm, "Majoritaet_ausrechnen"
    DoCmd.RunCommand acCmdRefresh
End Sub

Private Sub btn_ausw_1_Click()
    Dim st As String
    Dim Runde As String
    If no_runde_selected Then Exit Sub
    Runde = Me!Startklasse.Column(7)
    If Runde = "End_r_Akro" Or Runde = "End_r_schnell" Or Runde = "End_r" Or Runde = "End_r_2" Or Runde = "MK_5_TNZ" Then
        If Me!btn_ausw_2 = "Siegerehrung starten" Then
            If get_properties("EWS") = "EWS1" Then
                Beamer_generieren
                Me!Feld138.SetFocus
                Me!btn_ausw_1.Visible = False
            Else
                Me!btn_ausw_2 = Me.RecordsetClone.RecordCount
                Me!btn_ausw_1.Caption = "Platz " & Me!btn_ausw_2 & " anzeigen"
                st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=beamer_siegerehrung&text=" & Startklasse & "&mdb=" & get_TerNr & "&Platz=" & Me!btn_ausw_2 + 1)
            End If
        Else
            st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=beamer_siegerehrung&text=" & Startklasse & "&mdb=" & get_TerNr & "&Platz=" & Me!btn_ausw_2)
            If st = "beamer_siegerehrung" & Startklasse And Me!btn_ausw_2 > 0 Then
                Me!btn_ausw_2 = Me!btn_ausw_2 - 1
                Me!btn_ausw_1.Caption = "Platz " & Me!btn_ausw_2 & " anzeigen"
                If Me!btn_ausw_2 = 0 Then
                    Me!Feld138.SetFocus
                    Me!btn_ausw_1.Visible = False
                End If
            End If
        End If
    Else
        MsgBox "Es gibt keine Siegerehrung für diese Runde!"
    End If

End Sub

Private Function no_runde_selected()
    If (IsNull(Forms!Majoritaet_ausrechnen!Startklasse) Or (Forms!Majoritaet_ausrechnen!Startklasse = 0)) Then
       MsgBox ("Bitte Tanzrunde einstellen!")
       no_runde_selected = True
    End If
End Function

Function ueberschrift(cap)
    Dim ar, i
    If InStr(cap, ",") = 0 Then
        Me!ausw_anz.Caption = cap
    Else
        ar = Split(cap, ",")
        For i = 1 To 7
            Me("rde" & i).Caption = ar(i - 1)
        Next
    End If

End Function

Public Sub Startklasse_Change()
    
    Dim dbs As Database
    Set dbs = CurrentDb
    
    ' Test, ob in der aktuellen Runde, schon Majoritätseinträge vorhanden sind oder nicht
    ' Wenn nein, dann automatisch eine Wertung durchführen
    Dim rs As Recordset
    Dim ausw_anz
    Dim anz As Integer
    Dim anz_Wertungen As Integer
    Dim Startkl As String
    Dim ANZAHL_WR As Integer
    
    ausw_anz = Split(Nz(DLookup("ausw_anz", "Tanz_Runden_fix", "Runde='" & Startklasse.Column(7) & "'")), ";")
    ' RR   BW   F   BS
    If UBound(ausw_anz) > 0 Then
        If Startklasse.Column(7) = "MK_5_TNZ" And DLookup("Mehrkampfstationen", "Turnier", "Turniernum = 1") = "Kondition und Koordination" Then
            ueberschrift ausw_anz(1)
        Else
            Select Case left(Startklasse.Column(3), 3)
                Case "RR_"
                    ueberschrift ausw_anz(0)
                Case "BW_"
                    ueberschrift ausw_anz(1)
                Case "F_R"
                    ueberschrift ausw_anz(2)
                Case "F_B"
                    ueberschrift ausw_anz(1)
                Case "BS_"
                    ueberschrift ausw_anz(3)
                Case Else
                    ueberschrift " "
            End Select
        End If
    Else
        ueberschrift Nz(DLookup("ausw_anz", "Tanz_Runden_fix", "Runde='" & Startklasse.Column(7) & "'"))
    End If
    '***** 14_11 ***** Abfrage ob schon Wertungen vorhanden sind falls nein keine automatische Auswertung
    Set rs = dbs.OpenRecordset("SELECT count(*) as anzahl FROM Auswertung a INNER JOIN Paare_Rundenqualifikation p ON A.PR_ID = P.PR_ID WHERE p.RT_ID=" & Me!Startklasse & ";")
    anz_Wertungen = rs!Anzahl
    If Startklasse.Column(7) = "KO_r" Then
        Me!Ko_Sieger.Visible = True
        Me!Feld112.Visible = True
    Else
        Me!Ko_Sieger.Visible = False
        Me!Feld112.Visible = False
    End If
    If (Startklasse.Column(7) = "End_r" Or Startklasse.Column(7) = "End_r_Akro" Or Startklasse.Column(7) = "End_r_schnell" _
        Or Startklasse.Column(7) = "End_r_2" Or Startklasse.Column(7) = "MK_5_TNZ") Then '  And get_properties("EWS") = "EWS3" Then
        Me!btn_ausw_1.Visible = True
    Else
        Me!btn_ausw_1.Visible = False
    End If
    If anz_Wertungen = 0 Then
        Me!btn_ausw_1.Visible = False
        MsgBox "Zu dieser Runde gibt es noch keine Wertungen!"
    Else
        Me!btnPaareWeiternehmen.Visible = Me!Startklasse.Column(13)
        Startkl = Startklasse.Column(3)
        
        '                     Startklasse_Wertungsrichter
        Set rs = dbs.OpenRecordset("Select count(*) as AnzahlWR from Startklasse_wertungsrichter where Startklasse='" & Startkl & "';")
        ANZAHL_WR = rs!AnzahlWR
        Set rs = dbs.OpenRecordset("Select count(*) as anzahl from Majoritaet where rt_id=" & Startklasse & ";")
        anz = rs!Anzahl
        rs.Close
        gRT_ID = Startklasse
        Dim Runde As String
        Runde = Startklasse.Column(7)
        If anz_Wertungen <> anz * ANZAHL_WR Then
            Call majori_Click
        End If
        nächste_Runde = -1
        
        DoCmd.RepaintObject acForm, "Majoritaet_ausrechnen"
        
        nächste_Runde.Requery
    End If
    Requery
    If Runde = "End_r_Akro" Or Runde = "End_r_schnell" Or Runde = "End_r" Or Runde = "End_r_2" Or Runde = "MK_5_TNZ" Then
        Me!btn_ausw_2 = "Siegerehrung starten"
        Me!btn_ausw_1.Caption = "Siegerehrung starten"
    End If
    Me!Feld138.SetFocus
End Sub
