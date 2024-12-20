Option Explicit
    Dim dbs As Database
    Dim stDocName As String

Private Sub Berechnen_Click()   ' holt Anzahl Paare und trägt sie in die jeweils erste Runde ein
    Dim re  As Recordset
    Dim res As Recordset
    Dim paa As Recordset
    Dim strsql As String
    Dim anz As Integer
    Me.Requery
    Set dbs = CurrentDb
    Set re = Me.RecordsetClone
    strsql = "SELECT Rundentab.RT_ID, Rundentab.Turniernr, Rundentab.Runde, Rundentab.Startklasse, Rundentab.Anz_Paare, Rundentab.getanzt, Rundentab.Rundenreihenfolge, Rundentab.Startzeit, Rundentab.Paare, Rundentab.Dauer, Rundentab.WB, Rundentab.HTML, Rundentab.RT_Stat, Rundentab.ranking_anzeige, MSys__Tanz_Runden_fix.InAuswertung FROM Rundentab INNER JOIN MSys__Tanz_Runden_fix ON Rundentab.Runde = MSys__Tanz_Runden_fix.Runde WHERE (((Rundentab.Turniernr)=1)) ORDER BY Rundentab.Rundenreihenfolge;"
    Set res = dbs.OpenRecordset(strsql)
    re.MoveFirst
    Do Until re.EOF
        
        Set paa = dbs.OpenRecordset("SELECT Count(startkl) AS Anz FROM Paare GROUP BY Paare.Startkl, Paare.Anwesent_Status HAVING ((Paare.Startkl=""" & re!Startklasse & """) AND (Paare.Anwesent_Status=1));")
        If Not paa.EOF And re!Rundenreihenfolge < 999 Then
        res.FindFirst "(InAuswertung or Runde =""Vor_r_Fuß"") AND Startklasse=""" & re!Startklasse & """"
            re.Edit
            If res!RT_ID = re!RT_ID And Not res.NoMatch Then
                re!Paare = paa!anz
            End If
            re.Update
        End If
        re.MoveNext
    Loop
    Call Dauer_DblClick(0)
End Sub

Private Sub Feld81_AfterUpdate()
    Const stklassen = "RR_J, RR_S, RR_S1, RR_S2"
    If InStr(stklassen, Me!Feld81) > 0 And DLookup("Mehrkampfstationen", "Turnier", "Turniernum = 1") <> "" Then
        Me!Mehrkampf_eintragen.Visible = True
    Else
        Me!Mehrkampf_eintragen.Visible = False
    End If

    Select Case Me!Feld81
        Case "*"
            Me.RecordSource = "SELECT Rundentab.RT_ID, Rundentab.Turniernr, Rundentab.Runde, Rundentab.Startklasse, Rundentab.Anz_Paare, Rundentab.getanzt, Rundentab.Rundenreihenfolge, Rundentab.Startzeit, Rundentab.Paare, Rundentab.Dauer, Rundentab.WB, Rundentab.HTML, Rundentab.RT_Stat, Rundentab.ranking_anzeige, Rundentab.punkte_anzeige FROM Rundentab WHERE (((Rundentab.Turniernr)=" & get_aktTNr() & ")) ORDER BY Rundentab.Rundenreihenfolge;"
        Case "<"
            Me.RecordSource = "SELECT Rundentab.RT_ID, Rundentab.Turniernr, Rundentab.Runde, Rundentab.Startklasse, Rundentab.Anz_Paare, Rundentab.getanzt, Rundentab.Rundenreihenfolge, Rundentab.Startzeit, Rundentab.Paare, Rundentab.Dauer, Rundentab.WB, Rundentab.HTML, Rundentab.RT_Stat, Rundentab.ranking_anzeige, Rundentab.punkte_anzeige FROM Rundentab WHERE Rundentab.Rundenreihenfolge < 1000 ORDER BY Rundentab.Rundenreihenfolge;"
        Case Else
            Me.RecordSource = "SELECT Rundentab.RT_ID, Rundentab.Turniernr, Rundentab.Runde, Rundentab.Startklasse, Rundentab.Anz_Paare, Rundentab.getanzt, Rundentab.Rundenreihenfolge, Rundentab.Startzeit, Rundentab.Paare, Rundentab.Dauer, Rundentab.WB, Rundentab.HTML, Rundentab.RT_Stat, Rundentab.ranking_anzeige, Rundentab.punkte_anzeige FROM Rundentab WHERE (Rundentab.Startklasse=""" & Me!Feld81 & """ AND Rundentab.Turniernr= " & get_aktTNr() & ") ORDER BY Rundentab.Rundenreihenfolge;"
    End Select
    Requery
End Sub

Private Sub Feld81_DblClick(Cancel As Integer)
    Me!Feld81 = "*"
    Feld81_AfterUpdate
End Sub

Private Sub Form_Close()
    If get_properties("EWS") = "EWS3" And get_properties("Externer_Pfad") = 0 Then
        make_wr_zeitplan
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim re As Recordset
    Dim i, t As Integer
    If KeyCode = 90 Then  ' z
        Set re = Me.RecordsetClone
        re.MoveFirst
        For i = 1 To Me.Form.SelTop - 1
            re.MoveNext
        Next
        For i = 1 To Me.Form.SelHeight
            t = t + Nz(re!Dauer)
            re.MoveNext
        Next
        MsgBox t & " Minuten", , "Zeitdauer"
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim re As Recordset
    Dim sqlwhere, sqlstr As String
    
    Set dbs = CurrentDb
    Set re = dbs.OpenRecordset("Startklasse_Turnier")
    If re.RecordCount = 0 Then
        MsgBox "Es wurden noch keine Startklassen angelegt!"
        Exit Sub
    End If
    
    If get_properties("EWS") = "EWS1" Then
        Me!Kontrollkästchen84.Visible = True        ' HTML Seiten
        Me!Bezeichnungsfeld82.Visible = True        ' HTML Seiten
        Me!ranking.Visible = False
        Me!Bezeichnungsfeld91.Visible = False       ' ranking
        Me!Punkte.Visible = False
        Me!PunkteF.Visible = False
        Select Case Forms![A-Programmübersicht]!Turnierausw.Column(8)
            Case "D"
    
            Case Else
                Me!hochladen.Visible = False

        End Select
    Else
        Me!Kontrollkästchen84.Visible = False
        Me!Bezeichnungsfeld82.Visible = False
        Select Case Forms![A-Programmübersicht]!Turnierausw.Column(8)
            Case "D"
            
            Case "BW"
                Me!hochladen.Visible = False
                Me!ranking.Visible = True
                Me!Bezeichnungsfeld91.Visible = True
                Me!Punkte.Visible = True
                Me!PunkteF.Visible = True
            Case "BY"
                Me!hochladen.Visible = False
            Case Else
                Me.hochladen.Visible = False
        End Select
    End If
    
    Set re = dbs.OpenRecordset("SELECT * FROM Turnier;")
    
    Select Case re!MehrkampfStationen
        Case "Bodenturnen und Trampolin"
            sqlwhere = " UNION SELECT Runde, Rundentext, Rundenreihenfolge, MitStartklasse, R_IS_ENDRUNDE FROM Tanz_Runden_fix WHERE Runde Like 'MK_3*' Or Runde Like 'MK_4*' Or Runde Like 'MK_5*'"
        Case "Kondition und Koordination"
            Set re = dbs.OpenRecordset("SELECT * FROM (SELECT MK_11 FROM Turnier UNION  SELECT MK_12 FROM Turnier UNION SELECT MK_13 FROM Turnier UNION SELECT MK_21 FROM Turnier UNION SELECT MK_22 FROM Turnier UNION SELECT MK_23 FROM Turnier) WHERE NOT ISNULL([MK_11]) and [MK_11] <>'' ORDER BY MK_11;")
            If re.RecordCount = 0 Then
                sqlwhere = ""
            Else
                If re.RecordCount = 0 Then
                    sqlwhere = ""
                Else
                    re.MoveFirst
                    Do Until re.EOF
                        sqlwhere = sqlwhere & IIf(Len(sqlwhere) > 0, " OR", "") & " Runde=""" & re!MK_11 & """"
                        re.MoveNext
                    Loop
                    sqlwhere = " UNION SELECT Runde, Rundentext, Rundenreihenfolge, MitStartklasse, R_IS_ENDRUNDE FROM Tanz_Runden_fix WHERE" & sqlwhere & " OR Runde =""MK_5_TNZ"""
                End If
            End If
        Case "Breitensportwettbewerb"
            sqlwhere = " UNION SELECT Runde, Rundentext, Rundenreihenfolge, MitStartklasse, R_IS_ENDRUNDE FROM Tanz_Runden_fix WHERE Runde Like 'MK_6*' Or Runde Like 'MK_5*'"
        Case ""
            sqlwhere = ""
    End Select
    sqlstr = "SELECT Tanz_Runden_fix.Runde, Tanz_Runden_fix.Rundentext, Tanz_Runden_fix.Rundenreihenfolge, Tanz_Runden_fix.MitStartklasse, Tanz_Runden_fix.R_IS_ENDRUNDE FROM Tanz_Runden_fix WHERE Tanz_Runden_fix.Runde NOT LIKE 'MK_*'"
    sqlstr = sqlstr & sqlwhere
    sqlstr = sqlstr & " UNION SELECT Runde, Rundentext, Rundenreihenfolge, MitStartklasse, R_IS_ENDRUNDE FROM Tanz_Runden_erg ORDER BY Rundenreihenfolge;"
    
    Me!Kombinationsfeld53.RowSource = sqlstr
    
End Sub

Private Sub hochladen_Click()
    send_zeitplan Forms![A-Programmübersicht]!Turnier_Nummer
End Sub

Private Sub Kombinationsfeld51_Enter()
    If get_properties("highlight_ablauf") Then
        Me!Kombinationsfeld51.FormatConditions(1).Modify acFieldValue, acEqual, Chr(34) & Me!Kombinationsfeld51 & Chr(34), ""
    End If
End Sub

Private Sub Kombinationsfeld51_Exit(Cancel As Integer)
    Me!Kombinationsfeld51.FormatConditions(1).Modify acFieldValue, acEqual, """", ""
End Sub

Private Sub Kombinationsfeld53_Enter()
    If get_properties("highlight_ablauf") Then
        If Not IsNull(Me!Kombinationsfeld53) Then
            Me!Kombinationsfeld53.FormatConditions(0).Modify acExpression, acBetween, "[runde] Like """ & Me!Kombinationsfeld53 & "*""", ""
        End If
    End If
End Sub

Private Sub Kombinationsfeld53_Exit(Cancel As Integer)
    Me!Kombinationsfeld53.FormatConditions(0).Modify acExpression, acBetween, """", ""
End Sub

Private Sub Mehrkampf_eintragen_Click()
    Dim re, neu As Recordset
    Dim i, j As Integer
    Dim max_reihe As Integer
    Dim Runde
    Dim TNR As Integer
    
        TNR = get_aktTNr
        Set dbs = CurrentDb
        Set re = dbs.OpenRecordset("SELECT Max([Rundenreihenfolge]) AS Max_Reihe FROM Rundentab;")
        max_reihe = Int((Nz(re!max_reihe / 10) + 1)) * 10
        
        Set re = dbs.OpenRecordset("SELECT * FROM Turnier WHERE Turniernum=" & TNR & ";")
        Set neu = Me.RecordsetClone
        If Not re.EOF() Then
            Select Case DLookup("Mehrkampfstationen", "Turnier", "Turniernum = 1")
                Case "Kondition und Koordination"
                    For i = 1 To 2
                        For j = 1 To 3
                            If re("MK_" & i & j) <> "" And Not IsNull(re("MK_" & i & j)) Then _
                                Runde = Runde & re("MK_" & i & j) & ", "
                        Next
                    Next
                    Runde = Runde & "MK_5_TNZ, End_r, Sieger"
                Case "Bodenturnen und Trampolin"
                    Runde = "MK_3_BOT, MK_4_TRA, MK_5_TNZ, End_r, Sieger"
                Case Else
                    MsgBox "Diese Mehrkampfart wurde nicht definiert!"
                    Exit Sub
            End Select
            Runde = Split(Runde, ", ")
            make_rde Me!Feld81, Runde, ""
        End If
    DoCmd.Requery
End Sub

Private Sub Kombinationsfeld51_AfterUpdate()
    If left(Me!Kombinationsfeld53, 4) = "End_" Then Me!Kombinationsfeld64 = 1
    If left(Me!Kombinationsfeld51, 2) = "F_" Then Me!Kombinationsfeld64 = 1
End Sub

Private Sub Kombinationsfeld51_DblClick(Cancel As Integer)
    Me!Feld81 = Me!Kombinationsfeld51
    Feld81_AfterUpdate
End Sub

Private Sub Kontrollkästchen84_Click()
    If Me!Kontrollkästchen84 = False Then Me!RT_Stat = 0
End Sub

Private Sub Kombinationsfeld51_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub Kombinationsfeld64_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub Kombinationsfeld53_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub Dauer_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub Kontrollkästchen84_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub Paare_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub Punkte_Click()
    If Me.Kombinationsfeld53 = "Sieger" Then Exit Sub
    [punkte_anzeige] = Not [punkte_anzeige]
    DoCmd.Requery "punkte"
End Sub

Private Sub Punkte_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
    If KeyCode = 32 Then Punkte_Click
End Sub

Private Sub ranking_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
    If KeyCode = 32 Then ranking_Click
End Sub

Private Sub ranking_Click()
    If Me.Kombinationsfeld53 = "Sieger" Then Exit Sub
    [ranking_anzeige] = Not [ranking_anzeige]
    DoCmd.Requery "ranking"
End Sub

Private Sub Rundenreihenfolge_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub Startzeit_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub runden_ergaenzen_Click()
    Dim dbs As Database
    Dim rde As Recordset
    Dim rst As Recordset
    Dim tu As Recordset
    Dim stmt As String
    Dim Runde As Variant
    Dim msg As String
    Dim i, j As Integer
    Dim Startklasse_text As String
    
    check_ablaufzeit
    
    Set dbs = CurrentDb
    stmt = "SELECT DISTINCT Startklasse FROM Rundentab WHERE Startklasse <> '';"
    Set rde = dbs.OpenRecordset(stmt)
    If rde.RecordCount > 0 Then rde.MoveFirst
    Do Until rde.EOF
        stmt = "Select * from Startklasse where Startklasse='" & rde!Startklasse & "';"
        Set rst = dbs.OpenRecordset(stmt)
        If rst.RecordCount <> 0 Then
            Startklasse_text = rst!Startklasse_text
        
            Select Case get_bs_erg(rde!Startklasse, 5)
                Case "BW_MA"
                    stmt = "Select count(*) as anzahl from rundentab where turniernr=" & [Form_A-Programmübersicht]![Akt_Turnier] & " and Startklasse='" & rde!Startklasse & "' and Runde like 'Vor_r*';"
                    Set rst = dbs.OpenRecordset(stmt)
                    If rst!Anzahl > 0 Then      ' wenn eine geteile Vorrunde müssen beide da sein
                        Runde = Array("Vor_r_lang", "Vor_r_schnell", "Hoff_r")
                        If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & " Vorrunde, " & vbCrLf
                        dbs.Execute "DELETE * from rundentab WHERE turniernr=" & [Form_A-Programmübersicht]![Akt_Turnier] & " AND Startklasse='" & rde!Startklasse & "' AND Runde='Vor_r';"
                    End If
                    
                    dbs.Execute "DELETE * from rundentab WHERE turniernr=" & [Form_A-Programmübersicht]![Akt_Turnier] & " AND Startklasse='" & rde!Startklasse & "' AND Runde='End_r';"
                    Runde = Array("End_r_lang", "End_r_schnell", "Sieger")
                    If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & " Endrunde, " & vbCrLf
                
                Case "BW_SA"
                    stmt = "Select count(*) as anzahl from rundentab where turniernr=" & [Form_A-Programmübersicht]![Akt_Turnier] & " and Startklasse='" & rde!Startklasse & "' and Runde like 'Vor_r*';"
                    Set rst = dbs.OpenRecordset(stmt)
                    If rst!Anzahl > 0 Then      ' eine geteile Vorrunde darf nicht sein
                        Runde = Array("Vor_r", "Hoff_r")
                        If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & " Vorrunde, " & vbCrLf
                        dbs.Execute "DELETE * from rundentab WHERE turniernr=" & [Form_A-Programmübersicht]![Akt_Turnier] & " AND Startklasse='" & rde!Startklasse & "' AND Runde like 'Vor_r_*';"
                    End If
                    
                    dbs.Execute "DELETE * from rundentab WHERE turniernr=" & [Form_A-Programmübersicht]![Akt_Turnier] & " AND Startklasse='" & rde!Startklasse & "' AND Runde='End_r';"
                    Runde = Array("End_r_lang", "End_r_schnell", "Sieger")
                    If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & " Endrunde, " & vbCrLf
                                        
                Case "BW_MB", "BW_SB"
                    Runde = Array("Musik", "End_r", "Sieger")
                    If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & " Endrunde, " & vbCrLf
                
                Case "BW_JA"
                    stmt = "Select count(*) as anzahl from rundentab where turniernr=" & [Form_A-Programmübersicht]![Akt_Turnier] & " and Startklasse='" & rde!Startklasse & "' and Runde like 'Vor_r*';"
                    Set rst = dbs.OpenRecordset(stmt)
                    If rst!Anzahl > 0 Then      ' eine geteile Vorrunde darf nicht sein
                        Runde = Array("Vor_r", "Hoff_r")
                        If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & " Vorrunde, " & vbCrLf
                        dbs.Execute "DELETE * from rundentab WHERE turniernr=" & [Form_A-Programmübersicht]![Akt_Turnier] & " AND Startklasse='" & rde!Startklasse & "' AND Runde like 'Vor_r_*';"
                    End If
                    
                    dbs.Execute "DELETE * from rundentab WHERE turniernr=" & [Form_A-Programmübersicht]![Akt_Turnier] & " AND Startklasse='" & rde!Startklasse & "' AND Runde='End_r_*';"
                    Runde = Array("End_r", "Sieger")
                    If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & " Endrunde, " & vbCrLf
                                        
                Case "RR_A", "RR_B"
                    Runde = Array("Startbuchabgabe", "End_r_Fuß", "End_r_Akro", "Sieger")
                    If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & ", " & vbCrLf
                    
                    ' Löschen von End_r
                    dbs.Execute "DELETE * from rundentab WHERE turniernr=" & [Form_A-Programmübersicht]![Akt_Turnier] & " AND Startklasse='" & rde!Startklasse & "' AND Runde='End_r';"
                    
                Case "RR_C"
                    Runde = Array("Startbuchabgabe", "End_r", "Sieger")
                    If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & ", " & vbCrLf
                    
                    ' Löschen von End_r_
                    dbs.Execute "DELETE * from rundentab WHERE turniernr=" & [Form_A-Programmübersicht]![Akt_Turnier] & " AND Startklasse='" & rde!Startklasse & "' AND Runde Like 'End_r_*';"
                    
                Case "RR_S", "RR_J"
                    Runde = get_mk("Startbuchabgabe, End_r, Sieger")
                    If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & ", " & vbCrLf
                    
                    ' Löschen von End_r_
                    dbs.Execute "DELETE * from rundentab WHERE turniernr=" & [Form_A-Programmübersicht]![Akt_Turnier] & " AND Startklasse='" & rde!Startklasse & "' AND Runde Like 'End_r_*';"
                    
                Case "RR_S1", "RR_S2"
                    Runde = get_mk("Startbuchabgabe, Sieger")
                    
                    If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & ", " & vbCrLf
                     ' Löschen von End_r
                    dbs.Execute "DELETE * from rundentab WHERE turniernr=" & [Form_A-Programmübersicht]![Akt_Turnier] & " AND Startklasse='" & rde!Startklasse & "' AND Runde Like 'End_r*';"
                   
                Case "F_RR_", "F_BW_"
                    Runde = Array("Startbuchabgabe", "End_r", "Sieger")
                    If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & ", " & vbCrLf
                    
                    ' Löschen von End_r_
                    dbs.Execute "DELETE * from rundentab WHERE turniernr=" & [Form_A-Programmübersicht]![Akt_Turnier] & " AND Startklasse='" & rde!Startklasse & "' AND Runde Like 'End_r_*';"
                    
                Case "BWBS_", "SLBS_", "BYBS_"
                    stmt = "Select count(*) as anzahl from rundentab where turniernr=" & [Form_A-Programmübersicht]![Akt_Turnier] & " and Startklasse='" & rde!Startklasse & "' and Runde like 'End_r_*';"
                    Set rst = dbs.OpenRecordset(stmt)
                    If rst!Anzahl > 0 Then      ' wenn eine geteile Endrunde müssen beide da sein
                        Runde = Array("End_r_1", "End_r_2")
                        If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & ", " & vbCrLf
                        dbs.Execute "DELETE * from rundentab WHERE turniernr=" & [Form_A-Programmübersicht]![Akt_Turnier] & " AND Startklasse='" & rde!Startklasse & "' AND Runde='End_r';"
                    End If
                    Runde = Array("Sieger")
                    If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & ", " & vbCrLf
            End Select
        End If
        rde.MoveNext
    Loop
    If msg <> "" Then
        MsgBox "Es wurden bei " & vbCrLf & left(msg, Len(msg) - 3) & vbCrLf & "die fehlende(n) Runde(n) automatisch ergänzt."
    End If
    DoCmd.Requery
End Sub

Function make_rde(klasse, rde, Startklasse_text) As Boolean
    Dim dbs As Database
    Set dbs = CurrentDb
    Dim rst As Recordset
    Dim stmt As String
    Dim Reihenfolge As Integer
    Dim j As Integer
    
    For j = 0 To UBound(rde)
        stmt = "Select count(*) as anzahl from rundentab where turniernr=" & [Form_A-Programmübersicht]![Akt_Turnier] & " and Startklasse='" & klasse & "' and Runde='" & rde(j) & "';"
        Set rst = dbs.OpenRecordset(stmt)
        If rst!Anzahl = 0 Then
            Set rst = dbs.OpenRecordset("Select max(Rundenreihenfolge) as reihenfolge from rundentab WHERE Rundenreihenfolge < 999 AND Turniernr = " & [Form_A-Programmübersicht]![Akt_Turnier] & ";")
            If rst.EOF Then
                Reihenfolge = 1
            Else
                Reihenfolge = Nz(rst!Reihenfolge) + IIf(rst!Reihenfolge > 998, 0, 1)
            End If

            Set rst = dbs.OpenRecordset("Rundentab")
        
            rst.AddNew
            If rde(j) = "Startbuchabgabe" Then
                rst!Rundenreihenfolge = 1
            Else
                rst!Rundenreihenfolge = Reihenfolge
            End If
            rst!Turniernr = get_aktTNr
            rst!Startklasse = klasse
            rst!Runde = rde(j)
            rst!Anz_Paare = 2
            If InStr(1, rde(j), "End_r") > 0 Or left(rde(j), 3) = "MK_" Or rde(j) = "Startbuchabgabe" Or rde(j) = "Sieger" Then
                rst!Anz_Paare = 1
            End If
            rst.Update
            
            make_rde = True
        End If
    Next
            
    Set rst = Nothing
    Set dbs = Nothing
End Function

Private Sub Dauer_DblClick(Cancel As Integer)   ' berechnet  alle Zeiten neu
    Dim re As Recordset
    Dim next_t, next_h
    Dim st As Boolean
    Set dbs = CurrentDb
    Set re = Me.RecordsetClone
    Me.Requery
    re.MoveFirst
    Do Until re.EOF
        If re!Runde <> "Startbuchabgabe" And re!Runde <> "WR_Besp" Then
            If st Then
                If Not re.EOF Then
                    re.Edit
                    re!Startzeit = next_t + next_h
                    re.Update
                End If
            End If
            next_t = re!Startzeit
            next_h = (re!Dauer / 1440)
            If Not IsNull(re!Startzeit) Then st = True
        End If
        re.MoveNext
    Loop
    If get_properties("EWS") = "EWS3" Then _
        make_wr_zeitplan
        
End Sub

Private Sub schliesssen_Click()
    DoCmd.Close
    
End Sub

Private Sub Rundenplanung_Click()
    stDocName = "Rundenplanung"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub btnAblaufplanung_Click()
    stDocName = "Ablaufplanung"
    DoCmd.OpenReport stDocName, acPreview
    
End Sub

Private Sub btnAktualisieren_Click()
    Me.OrderBy = "[ablaufplanung].[Rundenreihenfolge]"
    Me.Requery
End Sub

Private Sub Kombinationsfeld53_AfterUpdate()        'Runde
    Dim rde As String
    If Not IsNull(Me!Kombinationsfeld53) Then
        rde = Kombinationsfeld53.Column(3)
        If rde = 0 Then
            Me!Startklasse = Null
            Me!Anz_Paare = 0
        End If
        
        rde = Kombinationsfeld53
        If InStr(1, rde, "End_r") > 0 Or InStr(1, rde, "MK_") > 0 Or rde = "Startbuchabgabe" Or rde = "Sieger" Then
            Me!Anz_Paare = 1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If get_properties("check_runden_on_close") = 1 Then
        runden_ergaenzen_Click
    End If
End Sub

Private Sub Zeitplan_Click()
    Dim out As Object
    Dim line As String
    Dim ht_pfad As String
    Dim st As String
    
    If get_properties("EWS") = "EWS1" Then
        ht_pfad = getBaseDir & "Apache2\htdocs\beamer\"
        line = make_beamer_zeitplan(RT_ID)
        line = Replace(line, "x__zoom", "")                  ' "style=""padding:200px""")
        
        Set out = file_handle(ht_pfad & "index.html")
        out.WriteLine (line)
        out.Close
    Else
        st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=beamer_zeitplan&text=" & Me!RT_ID)
    End If
End Sub

Public Function get_mk(rnd)        ' Mehrkampfstationen sammeln
    Dim db As Database
    Dim re As Recordset
    Dim i As Integer
    Set db = CurrentDb
    Set re = db.OpenRecordset("SELECT * FROM Turnier;")
    Select Case re!MehrkampfStationen
        Case "Bodenturnen und Trampolin"
            rnd = rnd & ", MK_3_BOT, MK_4_TRA, MK_5_TNZ"
        Case "Kondition und Koordination"
            Set re = db.OpenRecordset("SELECT * FROM (SELECT MK_11 FROM Turnier UNION  SELECT MK_12 FROM Turnier UNION SELECT MK_13 FROM Turnier UNION SELECT MK_21 FROM Turnier UNION SELECT MK_22 FROM Turnier UNION SELECT MK_23 FROM Turnier) WHERE NOT ISNULL([MK_11]) and [MK_11] <>'' ORDER BY MK_11;")
            If re.RecordCount = 0 Then
                get_mk = ""
            Else
                re.MoveFirst
                i = 0
                rnd = rnd & ", MK_5_TNZ"
                Do Until re.EOF
                    rnd = rnd & IIf(Len(rnd) > 0, ", ", "") & re!MK_11
                    re.MoveNext
                Loop
            End If
        Case "Breitensportwettbewerb"
            rnd = rnd & ", MK_6_KFT, MK_6_BAL, MK_6_KON, MK_5_TNZ"
        Case ""
            get_mk = ""
    End Select
    get_mk = Split(rnd, ", ")
End Function

Function check_ablaufzeit()
    Dim re As Recordset
    Set dbs = CurrentDb
    Dim prev As Date
    ' prüfen ob uhrzeit konsistent
    Set re = dbs.OpenRecordset("SELECT * FROM Rundentab WHERE Rundenreihenfolge < 1000 ORDER BY Rundenreihenfolge;")
    If re.RecordCount > 0 Then re.MoveFirst
    prev = re!Startzeit
    re.MoveNext
    Do Until re.EOF()
        If re!Runde <> "Startbuchabgabe" And re!Runde <> "WR_Besp" Then
            If DateDiff("n", prev, re!Startzeit) > 0 Then
                prev = re!Startzeit
            Else
                MsgBox "Der Zeitablauf bei " & re!Rundenreihenfolge & " " & re!Runde & " ist nicht durchgängig"
            End If
        
        End If
        re.MoveNext
    Loop

End Function

