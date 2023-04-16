Option Compare Database
Option Explicit

Sub lese_Auswerteunterlagen(Startklasse, st_kl)
    Dim oXLSApp As Object
    Dim oXLSWKB As Object
    Dim db As Database
    Dim wr, re, rt As Recordset
    Dim WR_ID
    Dim wr_anz As Integer
    Dim spalte
    Dim s_rd, PR_ID As Integer
    Dim y_off, x_off As Integer
    Dim RT_ID As Integer
    Dim i, csp As Integer
    Dim cgi As String
    Dim Runde As String
    Set db = CurrentDb()

'    startklassen = Array("RR_S1", "Schüler 1", "RR_S2", "Schüler 2", "RR_S", "Schüler", "RR_J", "Junioren", "RR_C", "C-Klasse")
    
    Set oXLSApp = CreateObject("Excel.Application")
    oXLSApp.Visible = True
    oXLSApp.DisplayAlerts = False
    If get_mk() = "Kondition und Koordination" Then
        Set oXLSWKB = oXLSApp.Workbooks.Open(getBaseDir & "Turn und Athletik-WB\2_Auswertungsunterlagen Tanz-Koordination-Kondition.xlsx", 3)
    Else                                                                                            ' nicht ändern wegen verknüpfungen
        Set oXLSWKB = oXLSApp.Workbooks.Open(getBaseDir & "Turn und Athletik-WB\1_Auswertungsunterlagen Tanz-Bodentunen-Trampolin.xlsx", 3)
    End If

    
    Set wr = db.OpenRecordset("SELECT Startklasse_Wertungsrichter.*, WR_Kuerzel FROM Wert_Richter INNER JOIN Startklasse_Wertungsrichter ON Wert_Richter.WR_ID = Startklasse_Wertungsrichter.WR_ID WHERE WR_function Like 'M*' AND Startklasse='" & st_kl & "' ORDER BY WR_Kuerzel;")
    i = 1
    If wr.RecordCount = 0 Then
        MsgBox "Für " & Startklasse & " sind keine WR eingeteilt!"
    Else
        wr.MoveFirst
        ReDim WR_ID(0)
        Do Until wr.EOF
            If wr!WR_function = "MA" Then
                WR_ID(0) = wr!WR_ID
            End If
            If wr!WR_function = "MB" Then
                ReDim Preserve WR_ID(i)
                WR_ID(i) = wr!WR_ID
                i = i + 1
            End If
            wr.MoveNext
        Loop
        Set wr = Nothing
        csp = 2
        For s_rd = 1 To 7
            Set re = db.OpenRecordset("SELECT * FROM Auswertung;")
            If get_mk() = "Kondition und Koordination" Then
                spalte = Array(4, "", 13, "", 16, "", 19, "", 24, "", 27, "", 30, "", 100, "")
                x_off = spalte(csp)
                Runde = ""
                If oXLSWKB.Worksheets(Startklasse).cells(11, x_off) <> "" Then
                    If x_off = 4 Then
                        Runde = "MK_5_TNZ"
                        wr_anz = UBound(WR_ID)
                    Else
                        Runde = DLookup("Runde", "Tanz_Runden_fix", "Rundentext ='MK_" & oXLSWKB.Worksheets(Startklasse).cells(11, x_off) & "'")
                        wr_anz = 0
                    End If
                End If
            Else
                If csp = 0 Then         '  And st_kl <> "RR_S1" And st_kl <> "RR_S2" Then
                    csp = 2
                End If
                spalte = Array(4, "MK_5_TNZ", 13, "MK_3_BOT", 24, "MK_4_TRA", 100, "")
                x_off = spalte(csp)
                wr_anz = UBound(WR_ID)
                Runde = spalte(csp + 1)
            End If
            If spalte(csp) = 100 Then Exit For
            csp = csp + 2
            Set rt = db.OpenRecordset("SELECT rt_id FROM  Rundentab WHERE Startklasse='" & st_kl & "' AND Runde='" & Runde & "';")
            If rt.RecordCount > 0 Then
                For y_off = 14 To 73 Step 2
                    cgi = lese_4(re, 0, oXLSWKB.Worksheets(Startklasse), y_off, x_off, WR_ID(0), rt!RT_ID)
                    If cgi = False Then Exit For
                    If wr_anz > 0 Then cgi = lese_4(re, 4, oXLSWKB.Worksheets(Startklasse), y_off, x_off + 1, WR_ID(1), rt!RT_ID)
                    If wr_anz > 1 Then cgi = lese_4(re, 4, oXLSWKB.Worksheets(Startklasse), y_off, x_off + 3, WR_ID(2), rt!RT_ID)
                    If wr_anz > 2 Then cgi = lese_4(re, 4, oXLSWKB.Worksheets(Startklasse), y_off, x_off + 5, WR_ID(3), rt!RT_ID)
                Next
                For i = 0 To wr_anz
                    sort_rnd rt!RT_ID, WR_ID(i)
                Next
            End If
        Next
    End If
     
    oXLSWKB.Save
    oXLSApp.DisplayAlerts = True
    oXLSWKB.Close
    oXLSApp.Quit
    Set oXLSWKB = Nothing
    Set oXLSApp = Nothing
    
End Sub

Function lese_4(re, MA, xsheet, y, x, WR_ID, RT_ID)
    Dim pu As String
    Dim pr
    Dim TP_ID As Integer
    Dim w_dis As Boolean     ' bei fehlendee Wertung Paar nicht angetreten
    If xsheet.cells(y, 1) <> "" Then
        lese_4 = True
        If xsheet.cells(y, x) = "" Or xsheet.cells(y + 1, x) = "" Then w_dis = True
        TP_ID = xsheet.cells(y, 1)
        pu = "PR_ID1=" & TP_ID & "&rt_ID=" & RT_ID & "&"
        pu = pu & "wmk_td1=" & xsheet.cells(y, x) & "&"
        pu = pu & IIf(xsheet.cells(y + 1, x) = "", "", "wmk_th1=" & xsheet.cells(y + 1, x) & "&")
        If MA > 0 Then
            If xsheet.cells(y, x + 1) = "" Or xsheet.cells(y + 1, x + 1) = "" Then w_dis = True
            pu = pu & "wmk_dd1=" & xsheet.cells(y, x + 1) & "&"
            pu = pu & "wmk_dh1=" & xsheet.cells(y + 1, x + 1) & "&"
        End If
        pu = pu & "WR_ID=" & WR_ID & "&"
        pu = pu & "Punkte1=" & celltoZahl(xsheet.cells(y, x)) + celltoZahl(xsheet.cells(y + 1, x)) + IIf(MA = 0, 0, celltoZahl(xsheet.cells(y, x + 1)) + celltoZahl(xsheet.cells(y + 1, x + 1)))
        If w_dis Then
            pu = pu & "&w_dis=1"
        End If
        pr = DLookup("PR_ID", "Paare_Rundenqualifikation", "TP_ID=" & TP_ID & " AND RT_ID=" & RT_ID & "")
        re.FindFirst ("pr_id=" & pr & " AND WR_ID=" & WR_ID & "")
        If re.NoMatch Then   'Abfrage vorhanden, wenn nicht neu
            re.AddNew
            re!PR_ID = pr
            re!WR_ID = WR_ID
        Else            ' oder edit
            re.Edit
        End If
        re!Punkte = celltoZahl(xsheet.cells(y, x)) + celltoZahl(xsheet.cells(y + 1, x))
        If MA > 0 Then
            re!Punkte = re!Punkte + celltoZahl(xsheet.cells(y, x + 1)) + celltoZahl(xsheet.cells(y + 1, x + 1))
        End If
        If w_dis Then re!Punkte = Null
    '    re!Reihenfolge = rh
        re!Cgi_Input = pu
        re!Platz = Null
        re.Update

    Else
        lese_4 = False
    End If

End Function

Private Function celltoZahl(cell)
    celltoZahl = 0
    If cell <> "" Then
        celltoZahl = cell
    End If
End Function

Function sort_rnd(rt, WR_ID)
    Dim db As Database
    Dim re As Recordset
    Dim pl, pl_m, pl_a As Integer
    Dim rd, zeit, so As String
    Dim stmt As String
    Set db = CurrentDb
    rd = DLookup("Runde", "Rundentab", "RT_ID=" & rt)
    zeit = "MK_1_KLE, MK_1_STL, MK_2_KAS, MK_2_KOO, MK_2_SCH, MK_2_STE"
    If InStr(zeit, rd) = 0 Then
        so = " DESC"
    Else
        so = ""
    End If
'hier muss die Reihenfoge gedreht werden bei Zeit stop
    stmt = "SELECT * from Auswertung a"
    stmt = stmt & " where a.wr_id=" & WR_ID & " and exists (select 1 from Paare_Rundenqualifikation pr where pr.pr_id=a.pr_id and pr.rt_id=" & rt & ")"
    stmt = stmt & " order by a.punkte " & so

    Set re = db.OpenRecordset(stmt)
    If re.RecordCount = 0 Then
        MsgBox "Es gibt noch keine Wertungen in dieser Tanzrunde!"
    Else
        re.MoveLast
        re.MoveFirst
        pl = 0
        pl_m = 0
        pl_a = 0
        Do Until re.EOF
            re.Edit
            If Not IsNull(re!Punkte) Then
                If pl_m = re!Punkte Then
                    pl_a = pl_a + 1
                    re!Platz = pl
                Else
                    pl = pl + 1 + pl_a
                    pl_m = re!Punkte
                        re!Platz = pl
                    pl_a = 0
                End If
            End If
            re.Update
            re.MoveNext
        Loop
    End If
End Function

Sub schreibe_Auswerteunterlagen()
    Dim oXLSApp As Object
    Dim oXLSWKB As Object
    Dim xsheet  As Object
    Dim db As Database
    Dim tu, re, wr As Recordset
    Dim t As Integer
    Dim retl As Integer
    Dim s1, s2, s, j, c As Integer
    Dim y_off As Integer
    Dim st_kl As String
    
    Set db = CurrentDb
    Set tu = db.OpenRecordset("Turnier")
    
    If get_mk = "" Then Exit Sub
    Set oXLSApp = CreateObject("Excel.Application")
    oXLSApp.Visible = True
    oXLSApp.DisplayAlerts = False
    If tu!MehrkampfStationen = "Kondition und Koordination" Then
        Set oXLSWKB = oXLSApp.Workbooks.Open(getBaseDir & "Turn und Athletik-WB\2_Auswertungsunterlagen Tanz-Koordination-Kondition.xlsx", 0)
    Else
        Set oXLSWKB = oXLSApp.Workbooks.Open(getBaseDir & "Turn und Athletik-WB\1_Auswertungsunterlagen Tanz-Bodentunen-Trampolin.xlsx", 0)
    End If
    
    Set xsheet = oXLSWKB.Worksheets("Turnierangaben")
    
    retl = 6
    If xsheet.cells(3, 3) = tu!Turnier_Name Then
        retl = MsgBox("Die EXCEL-Datei wurde schon befüllt, wirklich nochmal befüllen?", vbYesNo, " Achtung!")
    End If
    
    If retl = vbYes Then
        xsheet.cells(3, 3) = tu!Turnier_Name
        xsheet.cells(4, 3) = CStr(tu!T_Datum)
        xsheet.cells(5, 3) = tu!Veranst_Ort
        xsheet.cells(6, 3) = tu!Veranst_Name
        xsheet.cells(7, 3) = ""
        If tu!MehrkampfStationen = "Kondition und Koordination" Then
            xsheet.cells(8, 3) = Mid(DLookup("Rundentext", "Tanz_Runden_fix", "Runde='" & tu!MK_21 & "'"), 4)
            xsheet.cells(9, 3) = Mid(DLookup("Rundentext", "Tanz_Runden_fix", "Runde='" & tu!MK_22 & "'"), 4)
            xsheet.cells(10, 3) = Mid(DLookup("Rundentext", "Tanz_Runden_fix", "Runde='" & tu!MK_23 & "'"), 4)
            xsheet.cells(11, 3) = Mid(DLookup("Rundentext", "Tanz_Runden_fix", "Runde='" & tu!MK_11 & "'"), 4)
            xsheet.cells(12, 3) = Mid(DLookup("Rundentext", "Tanz_Runden_fix", "Runde='" & tu!MK_12 & "'"), 4)
            xsheet.cells(13, 3) = Mid(DLookup("Rundentext", "Tanz_Runden_fix", "Runde='" & tu!MK_13 & "'"), 4)
        End If
        
        Set re = db.OpenRecordset("SELECT * FROM paare WHERE Startkl<>'RR_A' AND Startkl<>'RR_B' AND (Anwesent_Status = 1 OR Anwesent_Status = 2) ORDER BY startkl, startnr;")
        re.MoveFirst
        t = 2
        Do Until re.EOF
        
            If left(re!Startkl, 4) = "RR_S" Or re!Startkl = "RR_J" Or re!Startkl = "RR_C" Then
                Set xsheet = oXLSWKB.Worksheets("Teilnehmer")
                xsheet.Range("B" & t).Value = re!Startkl
                xsheet.Range("C" & t).Value = re!Startnr
                xsheet.Range("D" & t).Value = re!Da_Vorname
                xsheet.Range("E" & t).Value = re!Da_NAchname
                xsheet.Range("F" & t).Value = re!He_Vorname
                xsheet.Range("G" & t).Value = re!He_Nachname
                xsheet.Range("H" & t).Value = re!Verein_nr
                xsheet.Range("I" & t).Value = re!Verein_Name
                xsheet.Range("J" & t).Value = re!Name_Team
                xsheet.Range("K" & t).Value = re!Startbuch
                xsheet.Range("L" & t).Value = re!TP_ID
                Set wr = db.OpenRecordset("SELECT Count(*) AS Ausdr1 FROM Startklasse_Wertungsrichter WHERE Startklasse='" & re!Startkl & "' AND WR_function='MB';")
    
                Select Case re!Startkl
                    Case "RR_S1"
                        st_kl = "Schüler 1"
                        y_off = s1 * 2 + 14
                        s1 = s1 + 1
                    Case "RR_S2"
                        st_kl = "Schüler 2"
                        y_off = s2 * 2 + 14
                        s2 = s2 + 1
                    Case "RR_S"
                        st_kl = "Schüler"
                        y_off = s * 2 + 14
                        s = s + 1
                    Case "RR_J"
                        st_kl = "Junioren"
                        y_off = j * 2 + 14
                        j = j + 1
                    Case "RR_C"
                        st_kl = "C-Klasse"
                        y_off = c * 2 + 14
                        c = c + 1
                End Select
                'oXLSWKB.Worksheets(st_kl).cells(y_off, 1) = re!TP_ID
                oXLSWKB.Worksheets(st_kl).cells(y_off, 2) = re!Startnr
                If re!Startkl = "RR_S1" Or re!Startkl = "RR_S2" Then oXLSWKB.Worksheets(st_kl).cells(10, 6) = wr!Ausdr1
                If get_mk <> "Kondition und Koordination" Then
                    oXLSWKB.Worksheets(st_kl).cells(10, 15) = wr!Ausdr1
                    oXLSWKB.Worksheets(st_kl).cells(10, 26) = wr!Ausdr1
                End If
                t = t + 1
            End If
            re.MoveNext
        Loop
      
    End If
    oXLSWKB.Save
    oXLSApp.DisplayAlerts = True
    oXLSWKB.Close
    oXLSApp.Quit
    Set oXLSWKB = Nothing
    Set oXLSApp = Nothing
 
End Sub

Function befülle_dateien()
    Dim oXLSApp
    Dim oXLSWKB
    Dim dateien
    Dim i As Integer
    
    If get_mk() = "Kondition und Koordination" Then
        dateien = Array("2_Wertung Kondition", "2_Wertung Koordination", "2_Wertung Tanzen - A-WR", "2_Wertung Tanzen - B-WR 1", "2_Wertung Tanzen - B-WR 2", "2_Wertung Tanzen - B-WR 3")
    Else
        dateien = Array("1_Wertung Bodenturnen", "1_Wertung Tanzen - A-WR", "1_Wertung Tanzen - B-WR 1", "1_Wertung Tanzen - B-WR 2", "1_Wertung Tanzen - B-WR 3")
    End If
    Set oXLSApp = CreateObject("Excel.Application")
    oXLSApp.Visible = True
    oXLSApp.DisplayAlerts = False
    For i = 0 To UBound(dateien) - 1
        Set oXLSWKB = oXLSApp.Workbooks.Open(getBaseDir & "Turn und Athletik-WB\" & dateien(i) & ".xlsx", 3)
        oXLSWKB.Save
        oXLSWKB.Close
    Next

End Function
