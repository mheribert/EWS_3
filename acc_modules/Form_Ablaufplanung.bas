Option Explicit
    Dim dbs As Database
    Dim stDocName As String

Private Sub Berechnen_Click()   ' holt Anzahl Paare und tr�gt sie in die jeweils erste Runde ein
    Dim re  As Recordset
    Dim res As Recordset
    Dim paa As Recordset
    Dim strSQL As String
    Dim anz As Integer
    Me.Requery
    Set dbs = CurrentDb
    Set re = Me.RecordsetClone
    strSQL = "SELECT Rundentab.RT_ID, Rundentab.Turniernr, Rundentab.Runde, Rundentab.Startklasse, Rundentab.Anz_Paare, Rundentab.getanzt, Rundentab.Rundenreihenfolge, Rundentab.Startzeit, Rundentab.Paare, Rundentab.Dauer, Rundentab.WB, Rundentab.HTML, Rundentab.RT_Stat, Rundentab.ranking_anzeige, MSys__Tanz_Runden_fix.InAuswertung FROM Rundentab INNER JOIN MSys__Tanz_Runden_fix ON Rundentab.Runde = MSys__Tanz_Runden_fix.Runde WHERE (((Rundentab.Turniernr)=1)) ORDER BY Rundentab.Rundenreihenfolge;"
    Set res = dbs.OpenRecordset(strSQL)
    re.MoveFirst
    Do Until re.EOF
        
        Set paa = dbs.OpenRecordset("SELECT Count(startkl) AS Anz FROM Paare GROUP BY Paare.Startkl, Paare.Anwesent_Status HAVING ((Paare.Startkl=""" & re!Startklasse & """) AND (Paare.Anwesent_Status=1));")
        If Not paa.EOF And re!Rundenreihenfolge < 999 Then
        res.FindFirst "(InAuswertung or Runde =""Vor_r_Fu�"") AND Startklasse=""" & re!Startklasse & """"
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
    If Me!Feld81 = "*" Then
        Me.RecordSource = "SELECT Rundentab.RT_ID, Rundentab.Turniernr, Rundentab.Runde, Rundentab.Startklasse, Rundentab.Anz_Paare, Rundentab.getanzt, Rundentab.Rundenreihenfolge, Rundentab.Startzeit, Rundentab.Paare, Rundentab.Dauer, Rundentab.WB, Rundentab.HTML, Tanz_Runden_fix.InAuswertung, Rundentab.RT_Stat, Rundentab.ranking_anzeige FROM Rundentab LEFT JOIN Tanz_Runden_fix ON Rundentab.Runde = Tanz_Runden_fix.Runde WHERE (((Rundentab.Turniernr)=" & get_aktTNr() & ")) ORDER BY Rundentab.Rundenreihenfolge;"
    Else
        Me.RecordSource = "SELECT Rundentab.RT_ID, Rundentab.Turniernr, Rundentab.Runde, Rundentab.Startklasse, Rundentab.Anz_Paare, Rundentab.getanzt, Rundentab.Rundenreihenfolge, Rundentab.Startzeit, Rundentab.Paare, Rundentab.Dauer, Rundentab.WB, Rundentab.HTML, Tanz_Runden_fix.InAuswertung, Rundentab.RT_Stat, Rundentab.ranking_anzeige FROM Rundentab LEFT JOIN Tanz_Runden_fix ON Rundentab.Runde = Tanz_Runden_fix.Runde WHERE (Rundentab.Startklasse=""" & Me!Feld81 & """ AND Rundentab.Turniernr= " & get_aktTNr() & ") ORDER BY Rundentab.Rundenreihenfolge;"
    End If
    Requery
End Sub

Private Sub Feld81_DblClick(Cancel As Integer)
    Me!Feld81 = "*"
    Feld81_AfterUpdate
End Sub

Private Sub Form_Open(Cancel As Integer)
    If Not Forms![A-Programm�bersicht]!Turnierausw.Column(8) = "D" Then
        Me.hochladen.Visible = False
        Me.Zeitplan.Visible = False
    End If
End Sub

Private Sub hochladen_Click()
    send_zeitplan Forms![A-Programm�bersicht]!Turnier_Nummer
End Sub

Private Sub Kombinationsfeld51_AfterUpdate()
    If left(Me!Kombinationsfeld53, 4) = "End_" Then Me!Kombinationsfeld64 = 1
    If left(Me!Kombinationsfeld51, 2) = "F_" Then Me!Kombinationsfeld64 = 1
End Sub

Private Sub Kombinationsfeld51_DblClick(Cancel As Integer)
    Me!Feld81 = Me!Kombinationsfeld51
    Feld81_AfterUpdate
End Sub

Private Sub Kontrollk�stchen84_Click()
    If Me!Kontrollk�stchen84 = False Then Me!RT_Stat = 0
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

Private Sub Kontrollk�stchen84_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub Paare_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub ranking_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
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
    Dim stmt As String
    Dim Runde As Variant
    Dim msg As String
    Dim i, j As Integer
    Dim Startklasse_text As String
    
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
                    stmt = "Select count(*) as anzahl from rundentab where turniernr=" & [Form_A-Programm�bersicht]![Akt_Turnier] & " and Startklasse='" & rde!Startklasse & "' and Runde like 'Vor_r*';"
                    Set rst = dbs.OpenRecordset(stmt)
                    If rst!Anzahl > 0 Then      ' wenn eine geteile Vorrunde m�ssen beide da sein
                        Runde = Array("Vor_r_lang", "Vor_r_schnell", "Hoff_r")
                        If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & " Vorrunde, " & vbCrLf
                        dbs.Execute "DELETE * from rundentab WHERE turniernr=" & [Form_A-Programm�bersicht]![Akt_Turnier] & " AND Startklasse='" & rde!Startklasse & "' AND Runde='Vor_r';"
                    End If
                    
                    dbs.Execute "DELETE * from rundentab WHERE turniernr=" & [Form_A-Programm�bersicht]![Akt_Turnier] & " AND Startklasse='" & rde!Startklasse & "' AND Runde='End_r';"
                    Runde = Array("End_r_lang", "End_r_schnell", "Sieger")
                    If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & " Endrunde, " & vbCrLf
                
                Case "BW_SA"
                    stmt = "Select count(*) as anzahl from rundentab where turniernr=" & [Form_A-Programm�bersicht]![Akt_Turnier] & " and Startklasse='" & rde!Startklasse & "' and Runde like 'Vor_r*';"
                    Set rst = dbs.OpenRecordset(stmt)
                    If rst!Anzahl > 0 Then      ' eine geteile Vorrunde darf nicht sein
                        Runde = Array("Vor_r", "Hoff_r")
                        If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & " Vorrunde, " & vbCrLf
                        dbs.Execute "DELETE * from rundentab WHERE turniernr=" & [Form_A-Programm�bersicht]![Akt_Turnier] & " AND Startklasse='" & rde!Startklasse & "' AND Runde like 'Vor_r_*';"
                    End If
                    
                    dbs.Execute "DELETE * from rundentab WHERE turniernr=" & [Form_A-Programm�bersicht]![Akt_Turnier] & " AND Startklasse='" & rde!Startklasse & "' AND Runde='End_r';"
                    Runde = Array("End_r_lang", "End_r_schnell", "Sieger")
                    If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & " Endrunde, " & vbCrLf
                                        
                Case "BW_MB", "BW_SB"
                    Runde = Array("End_r", "Sieger")
                    If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & " Endrunde, " & vbCrLf
                
                Case "BW_JA"
                    stmt = "Select count(*) as anzahl from rundentab where turniernr=" & [Form_A-Programm�bersicht]![Akt_Turnier] & " and Startklasse='" & rde!Startklasse & "' and Runde like 'Vor_r*';"
                    Set rst = dbs.OpenRecordset(stmt)
                    If rst!Anzahl > 0 Then      ' eine geteile Vorrunde darf nicht sein
                        Runde = Array("Vor_r", "Hoff_r")
                        If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & " Vorrunde, " & vbCrLf
                        dbs.Execute "DELETE * from rundentab WHERE turniernr=" & [Form_A-Programm�bersicht]![Akt_Turnier] & " AND Startklasse='" & rde!Startklasse & "' AND Runde like 'Vor_r_*';"
                    End If
                    
                    dbs.Execute "DELETE * from rundentab WHERE turniernr=" & [Form_A-Programm�bersicht]![Akt_Turnier] & " AND Startklasse='" & rde!Startklasse & "' AND Runde='End_r_*';"
                    Runde = Array("End_r", "Sieger")
                    If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & " Endrunde, " & vbCrLf
                                        
                Case "RR_A", "RR_B"
                    Runde = Array("End_r_Fu�", "End_r_Akro", "Sieger")
                    If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & ", " & vbCrLf
                    
                    ' L�schen von End_r
                    dbs.Execute "DELETE * from rundentab WHERE turniernr=" & [Form_A-Programm�bersicht]![Akt_Turnier] & " AND Startklasse='" & rde!Startklasse & "' AND Runde='End_r';"
                    
                '****** ToDo Einf�gen Mehrkampf, wenn Option gesetzt! ****
                '****** ToDo Endrunden nur bei mehr als sieben Paaren ****
                
                Case "RR_S", "RR_J", "RR_C"
                    Runde = Array("End_r", "Sieger")
                    If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & ", " & vbCrLf
                    
                    ' L�schen von End_r_
                    dbs.Execute "DELETE * from rundentab WHERE turniernr=" & [Form_A-Programm�bersicht]![Akt_Turnier] & " AND Startklasse='" & rde!Startklasse & "' AND Runde Like 'End_r_*';"
                    
                Case "BWBS_", "SLBS_"
                    stmt = "Select count(*) as anzahl from rundentab where turniernr=" & [Form_A-Programm�bersicht]![Akt_Turnier] & " and Startklasse='" & rde!Startklasse & "' and Runde like 'End_r_*';"
                    Set rst = dbs.OpenRecordset(stmt)
                    If rst!Anzahl > 0 Then      ' wenn eine geteile Endrunde m�ssen beide da sein
                        Runde = Array("End_r_1", "End_r_2")
                        If make_rde(rde!Startklasse, Runde, Startklasse_text) Then msg = msg & Startklasse_text & ", " & vbCrLf
                        dbs.Execute "DELETE * from rundentab WHERE turniernr=" & [Form_A-Programm�bersicht]![Akt_Turnier] & " AND Startklasse='" & rde!Startklasse & "' AND Runde='End_r';"
                    End If
            End Select
        End If
        rde.MoveNext
    Loop
    If msg <> "" Then
        MsgBox "Es wurden bei " & vbCrLf & left(msg, Len(msg) - 3) & vbCrLf & "die fehlende(n) Runde(n) automatisch erg�nzt."
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
        stmt = "Select count(*) as anzahl from rundentab where turniernr=" & [Form_A-Programm�bersicht]![Akt_Turnier] & " and Startklasse='" & klasse & "' and Runde='" & rde(j) & "';"
        Set rst = dbs.OpenRecordset(stmt)
        If rst!Anzahl = 0 Then
            Set rst = dbs.OpenRecordset("Select max(Rundenreihenfolge) as reihenfolge from rundentab WHERE Rundenreihenfolge < 999 AND Turniernr = " & [Form_A-Programm�bersicht]![Akt_Turnier] & ";")
            If rst.EOF Then
                Reihenfolge = 1
            Else
                Reihenfolge = rst!Reihenfolge + IIf(rst!Reihenfolge > 998, 0, 1)
            End If

            Set rst = dbs.OpenRecordset("Rundentab")
        
            rst.AddNew
            rst!Rundenreihenfolge = Reihenfolge
            rst!Turniernr = get_aktTNr
            rst!Startklasse = klasse
            rst!Runde = rde(j)
            rst!Anz_Paare = IIf(InStr(1, rde(j), "End_r") > 0, 1, 2)
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
    runden_ergaenzen_Click
    Me.Requery
End Sub

Private Sub Kombinationsfeld53_AfterUpdate()        'Runde
    If (Kombinationsfeld53.Column(3) = 0) Then
        Me!Startklasse = Null
        Me!Anz_Paare = 0
    End If
    If InStr(1, Kombinationsfeld53.Column(1), "Endrunde") > 0 Then
        Me!Anz_Paare = 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    runden_ergaenzen_Click
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
        
        Set out = file_handle(ht_pfad & "anzeige.html")
        out.writeline (line)
        out.Close
    Else
        st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=beamer_zeitplan&text=" & Me!RT_ID)
    End If
    
        

End Sub