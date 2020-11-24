Option Compare Database
Option Explicit

Private Sub Befehl27_Click()
On Error GoTo Err_Befehl27_Click


    DoCmd.Close

Exit_Befehl27_Click:
    Exit Sub

Err_Befehl27_Click:
    MsgBox err.Description
    Resume Exit_Befehl27_Click
    
End Sub

Sub Kombinationsfeld30_AfterUpdate()
    ' Den mit dem Steuerelement übereinstimmenden Datensatz suchen.
    Me.RecordsetClone.FindFirst "[ident] = " & Me![Kombinationsfeld30]
    Me.Bookmark = Me.RecordsetClone.Bookmark
End Sub

Private Sub Form_Open(Cancel As Integer)
    Select Case Forms![A-Programmübersicht]!Turnierausw.Column(8)
        Case "D"
            If DLookup("Mehrkampfstationen", "Turnier", "Turniernum = 1") <> "" Then
                Me!mehrkampf_einlesen.Visible = True
            End If
        Case "SL"
            Me!mehrkampf_einlesen.Visible = False
        Case "BY"
            Me!mehrkampf_einlesen.Visible = False
        Case Else
            Me!mehrkampf_einlesen.Visible = False
    End Select

    Call Turnier_aktuell_check_VB
End Sub

Private Sub Form_Resize()
    Me.[Wertung aufnehmen1 Unterformular].Height = Me.InsideHeight - 2400
    Me.[Paare_ohne_Punkte_UF].Height = Me.InsideHeight - 2400
    
End Sub

Sub Tanzrunde_AfterUpdate()
    Wertungsrichter_einstellen.Requery
    Wertungsrichter_einstellen = Null
    [Form_Wertung aufnehmen1 Unterformular].Requery
    Form_Paare_ohne_Punkte_UF.Requery
    [Form_A-Programmübersicht]![Tanzrunde] = Me!Tanzrunde
    Me!Wertungsrichter_einstellen.Requery
    AnzahlWR = Wertungsrichter_einstellen.ListCount
    Dim dbs As Database
    Dim Turniernr As Integer
    Dim Startklasse_einstellen As String
    Dim AnzahlWRVorgabe As Integer
    Turniernr = [Form_A-Programmübersicht]![Akt_Turnier]
    Startklasse_einstellen = [Forms]![wertung_aufnehmen]!Tanzrunde.Column(3)
    Set dbs = CurrentDb
    Dim rs As Recordset
    Set rs = dbs.OpenRecordset("Select * from startklasse sk, startklasse_Turnier skt where sk.Startklasse='" & Startklasse_einstellen & "' and skt.startklasse=sk.startklasse and skt.Turniernr=" & Turniernr)
    If Right([Forms]![wertung_aufnehmen]!Tanzrunde.Column(1), 19) = "Endrunde Fußtechnik" Then
        Maxwertung = 100
        Else
        If Right([Forms]![wertung_aufnehmen]!Tanzrunde.Column(1), 18) = "Endrunde Akrobatik" Then
        Maxwertung = 100
          Else
            If Right([Forms]![wertung_aufnehmen]!Tanzrunde.Column(1), 13) = "Zwischenrunde" And (Startklasse_einstellen = "RR_A" Or Startklasse_einstellen = "RR_B") Then
            Maxwertung = 100
            Else
            Maxwertung = rs!Maxwertung
            End If
          End If
    End If
    AnzahlWRVorgabe = rs!AnzahlWR
    rs.Close
    
    If (Not [Form_A-Programmübersicht]!Getrennte_Auslosung) Then
     '*****AB***** V13.02 if-Clause um neue Boogie Startklassen erweitert
     '*****AB***** V13.04 BW_SB und BW_MB in Case wieder entfernt, da nur eine Endrunde getanzt wird
        If (Startklasse_einstellen = "BW_H" Or Startklasse_einstellen = "BW_O" Or Startklasse_einstellen = "BW_MA" Or Startklasse_einstellen = "BW_SA") And ([Forms]![wertung_aufnehmen]!Tanzrunde.Column(7) = "End_r_lang" Or [Forms]![wertung_aufnehmen]!Tanzrunde.Column(7) = "End_r_schnell") Then
            ' Update der Rundeneinteilung
            Dim rt_id_endr As Integer
            rt_id_endr = getRT_ID(Turniernr, Startklasse_einstellen, "End_r")
            Call UpdateRundenqualifikation(rt_id_endr, Tanzrunde, False)
        End If
    End If
    Me!Feld72.SetFocus
    ' WR-Auswahl funktioniert nur, wenn die Anzahl der zugewiesenen
    ' WR mit der Anzahl aus den Turnierdaten übereinstimmt
'    Wertungsrichter_einstellen.Enabled = (AnzahlWRVorgabe = AnzahlWR)
    
'    If (AnzahlWRVorgabe <> AnzahlWR) Then
'        Call MsgBox("Die Anzahl der zugewiesenen Wertungsrichter stimmt nicht mit der Vorgabe aus den Turnierdaten überein:" & Chr(13) & Chr(13) & "Anzahl der Wertungsrichter gem. Turnierdaten: " & AnzahlWRVorgabe & Chr(13) & "Anzahl der tatsächlich eingeteilten Wertungsrichter: " & AnzahlWR & Chr(13) & Chr(13) & "Aus diesem Grund können Sie keine Wertungen eingeben.", vbInformation Or vbOKOnly)
'    End If
    
End Sub

Private Sub Befehl56_Click()
On Error GoTo Err_Befehl56_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Check_Wertungen"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Befehl56_Click:
    Exit Sub

Err_Befehl56_Click:
    MsgBox err.Description
    Resume Exit_Befehl56_Click
    
End Sub

Private Sub Befehl57_Click()
On Error GoTo Err_Befehl57_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Monitor_Runden"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Befehl57_Click:
    Exit Sub

Err_Befehl57_Click:
    MsgBox err.Description
    Resume Exit_Befehl57_Click
    
End Sub

Private Sub mehrkampf_einlesen_Click()
    If tanzrunde_selected Then Exit Sub
    
    Dim tr As String
    tr = Tanzrunde.Column(4)
    If InStr(7, tr, "klasse") Then
        tr = Replace(Tanzrunde.Column(4), "klasse", "")
    End If
    lese_Auswerteunterlagen tr, Tanzrunde.Column(3)
        
    Debug.Print tr
End Sub

Private Sub Wertung_aufnehmen1_Unterformular_Enter()
'    Call ActivateTextfields
End Sub

Public Sub ActivateTextfields()
    'Dim Runde As String
    If [Forms]![wertung_aufnehmen]!Tanzrunde.Column(8) = 1 Then
       [Wertung aufnehmen1 Unterformular]!Platz.TabStop = True
       [Wertung aufnehmen1 Unterformular]!Platz.Enabled = True
    Else
       [Wertung aufnehmen1 Unterformular]!Platz.TabStop = False
       [Wertung aufnehmen1 Unterformular]!Platz.Enabled = False
    End If
End Sub

Public Sub Wertung_aufnehmen1_Unterformular_Exit(Cancel As Integer)
    
    Dim dbs As Database
    Dim rstauswertung, rstweiter, rstanzahl As Recordset
    ' Bezug auf aktuelle Datenbank zurückgeben.
    Set dbs = CurrentDb
    ' Anzahl Paare für diese Runden in die Tabelle schreiben
    Dim anzahl_p As Double
    Dim werund, tr As String
    Dim Turniernr As Integer
    If IsNull(Tanzrunde.Column(7)) Then Exit Sub
    Turniernr = [Form_A-Programmübersicht]![Akt_Turnier]
    Dim stmt As String
    Dim IsEndrunde As Boolean
    IsEndrunde = (Tanzrunde.Column(14) = 1)
    
    ' Wertung überprüfen und Plätze vergeben
    Dim zpl As Double, zpu As Double, zpldup As Double
    zpl = 0
    zpu = 0
    ' Recordset-Objekt vom Typ Dynaset erstellen. Tabelle Auswertung öffnen
    stmt = "SELECT count(*) as anz from Auswertung a, Paare_Rundenqualifikation pr"
    stmt = stmt & " where a.wr_id=" & Wertungsrichter_einstellen & " and pr.pr_id=a.pr_id and pr.rt_id=" & Tanzrunde
    stmt = stmt & " and Punkte is null"
    Set rstauswertung = dbs.OpenRecordset(stmt)
    Dim Count As Integer
    Count = rstauswertung!anz
    rstauswertung.Close
    If (Count > 0) Then
        Exit Sub
    End If
    
    stmt = "SELECT * from Auswertung a"
    stmt = stmt & " where a.wr_id=" & Wertungsrichter_einstellen & " and exists (select 1 from Paare_Rundenqualifikation pr where pr.pr_id=a.pr_id and pr.rt_id=" & Tanzrunde & ")"
    stmt = stmt & " order by a.punkte desc, a.platz asc"
    
    Set rstauswertung = dbs.OpenRecordset(stmt)
    If rstauswertung.EOF() Then
        Exit Sub
    End If
    rstauswertung.MoveFirst
    With rstauswertung
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
    Wertung_aufnehmen1_Unterformular.Requery
    rstauswertung.Sort = "Platz"
    rstauswertung.MoveFirst
    If Not rstauswertung.EOF() Then
        zpl = !Platz
        rstauswertung.MoveNext
        If (IsEndrunde) Then ' Falls Endrunde
            Do While Not rstauswertung.EOF()
              If !Platz > zpl Then
                 zpl = !Platz
              Else
'                 MsgBox ("Gleiche Platzvergabe in der Endrunde ist unzulässig. Platz " & !Platz & " wurde mehrfach vergeben!")
'                 End
              End If
            rstauswertung.MoveNext
            Loop
        End If
    End If
    
    End With
    Me.Refresh
End Sub

Private Sub Wertungsrichter_einstellen_AfterUpdate()
    
    Dim dbs As Database
    Dim rstauswertung, Qualifikation As Recordset
    Dim rtid, wrid, Turniernr As Integer
    Dim updCmd As String
    rtid = Me![Tanzrunde].Column(0)
    wrid = Wertungsrichter_einstellen.Column(0)
    Turniernr = [Form_A-Programmübersicht]![Akt_Turnier]
    
    Set dbs = CurrentDb
    ' Recordset-Objekt vom Typ Dynaset erstellen. Tabelle Auswertung öffnen
    Dim sqlcmd As String
    
    ' Fehlende Wertungen hinzufügen
    sqlcmd = "select * from Paare_Rundenqualifikation pr where rt_id=" & rtid & " and anwesend_Status=1 and rundennummer is not null"
    sqlcmd = sqlcmd & " and not exists (select 1 from Auswertung a where a.pr_id=pr.pr_id and a.WR_ID=" & wrid & ")"
     
    Dim rsAddWertung As Recordset
    
    Set rsAddWertung = dbs.OpenRecordset(sqlcmd)
    
    Do While (Not rsAddWertung.EOF())
        Dim insCmd As String
        insCmd = "insert into Auswertung(PR_ID, WR_ID, Punkte, Platz, Reihenfolge)"
        insCmd = insCmd & " values(" & rsAddWertung!PR_ID & ", " & wrid & ", null, 0, " & rsAddWertung!Rundennummer & ")"
        
        dbs.Execute (insCmd)
        
        rsAddWertung.MoveNext
    Loop
    
    rsAddWertung.Close
    
    ' Wertungen löschen, die nicht rein gehören
    sqlcmd = "select distinct pr.pr_id from Paare_Rundenqualifikation pr, Auswertung a where a.pr_id=pr.pr_id and pr.rt_id=" & rtid & " and anwesend_Status<>1"
    Set Qualifikation = dbs.OpenRecordset(sqlcmd)
    Do While (Not Qualifikation.EOF())
        
        updCmd = "Delete from Auswertung where pr_id=" & Qualifikation!PR_ID
        
        dbs.Execute (updCmd)
        
        Qualifikation.MoveNext
    Loop
    
    Qualifikation.Close
    
    ' Wertungen noch in die richtige Reihenfolge bringen
    sqlcmd = "select * from Paare_Rundenqualifikation pr where rt_id=" & rtid & " and anwesend_Status=1 and rundennummer is not null"
    
    Set rstauswertung = dbs.OpenRecordset(sqlcmd)
    Do While (Not rstauswertung.EOF())
        
        updCmd = "Update Auswertung a set reihenfolge=" & rstauswertung!Rundennummer
        updCmd = updCmd & " where a.pr_id=" & rstauswertung!PR_ID
        
        dbs.Execute (updCmd)
        
        rstauswertung.MoveNext
    Loop
    
    rstauswertung.Close
    
    [Form_Wertung aufnehmen1 Unterformular].Requery
    Form_Paare_ohne_Punkte_UF.Requery
    Call Wertung_aufnehmen1_Unterformular_Enter
End Sub

Function tanzrunde_selected()
    If IsNull(Me!Tanzrunde) Or Me!Tanzrunde = 0 Then
       MsgBox ("Bitte Tanzrunde auswählen!")
       tanzrunde_selected = True
    End If
End Function

