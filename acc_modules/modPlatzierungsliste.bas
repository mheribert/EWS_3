Option Compare Database

Public Sub showReport_Platzierte_Paare()
' Zuerst die Daten in die Tabelle 'Report_Platzierte_Paare' schreiben
    Dim dbs As Database
    Set dbs = CurrentDb
    Dim rst, rstWR As Recordset
    Dim rstZiel As Recordset
    Dim Verein As String
    Dim Anzahl As Integer
    Call dbs.Execute("DELETE FROM Report_Platzierte_Paare")
    
    Dim stmt As String
    Dim i As Integer
    Dim RT_ID As Integer, Turniernr As Integer
    Dim Startkl As String
    
    RT_ID = [Form_A-Programmübersicht].Report_RT_ID
    Turniernr = [Form_A-Programmübersicht]![Akt_Turnier]
    
    ' Startklasse ermitteln
    stmt = "Select Startklasse from rundentab where rt_id=" & RT_ID
    Set rst = dbs.OpenRecordset(stmt)
    If (rst.EOF) Then
        MsgBox "Falsche Runde!"
        Exit Sub
    End If
    
    Startkl = rst!Startklasse
    rst.Close
    
    Call fillReport_Platzierte_Paare(dbs, 0, RT_ID, "Moderation", "")
    '*****AB***** V13_05C - NEU die WR nur dann anlegen, wenn nicht Rock'n'Roll Einzel
    If InStr(Startkl, "RR_") <> 1 And InStr(Startkl, "F_RR") <> 1 And left(Startkl, 3) <> "BW_" Then
        ' Schleife über alle WR
        stmt = "SELECT sw.Startklasse, wr.*"
        stmt = stmt & " FROM Wert_Richter wr, Startklasse_Wertungsrichter sw, Rundentab rt"
        stmt = stmt & " WHERE wr.WR_ID=sw.WR_ID and sw.startklasse=rt.startklasse"
        stmt = stmt & " and rt.rt_id=" & RT_ID
        stmt = stmt & " and wr.Turniernr=" & Turniernr
        '*****AB***** V13_04 - weitere WHERE_CLAUSE eingefügt, damit nur Wertungsrichter durchlaufen werden
        stmt = stmt & " and (sw.WR_function = 'X' or sw.WR_function = 'Ft' or sw.WR_function = 'Ak')"
        stmt = stmt & " ORDER BY wr.WR_Kuerzel"
        
        Dim WR_Name As String
        Set rst = dbs.OpenRecordset(stmt)
        Dim wr_count As Integer
        wr_count = 1
        '*****AB***** V13_05C - wegen mehr als sieben WR im neuen System beim RnR auf sieben begrenzt!
        Do While (Not rst.EOF And wr_count < 8)
            WR_Name = rst!WR_Vorname & " " & rst!WR_Nachname
            
            ' Jetzt für diesen WR die Wertung schreiben
            Call fillReport_Platzierte_Paare(dbs, wr_count, RT_ID, WR_Name, "Wertungsrichter " & rst!WR_Kuerzel)
            
            wr_count = wr_count + 1
            rst.MoveNext
        Loop
        rst.Close
    End If
    dbs.Close

End Sub

Public Sub fillReport_Platzierte_Paare(dbs As Database, Count As Integer, RT_ID As Integer, WR_Name As String, WR_Kurz As String)
    
    Dim stmt, stmt2 As String
    ' Jetzt für diesen WR die Wertung schreiben
    stmt = "SELECT t.Turniernum, r.R_NAME_ABLAUF, p.Startkl, p.Startklasse_text, p.Startnr, p.Name, p.Verein_Name, m.Platz, m.MajoritaetKurz, t.Turnier_Name, t.T_Datum, t.Veranst_Name, t.Veranst_Ort, r.Rundentext, m.disqualifiziert, m.punktabzug, m.anmerkung"
    If (Count = 0) Then
        stmt = stmt & Chr(13) & " , m.MajoritaetKurz as PlatzWR"
    Else
'***** AB ***** V 13.05 - hier m.Platz eingefügt, da m.WR8 nicht verfügbar war
        'stmt = stmt & Chr(13) & " , m.Platz as PlatzWR"
'***** AB ***** V 13.05C - hier m.WR + count eingefügt, da m.Platz nicht funktionierte, zusätzlich begrenzt in showReport_Platzierte_Paare() auf 7 WR, da dies den Fehler auslöste!
        stmt = stmt & Chr(13) & " , m.WR" & Count & " as PlatzWR"
    End If
    stmt = stmt & Chr(13) & " FROM View_Runden AS r, View_Majoritaet AS m, View_Paare AS p, Turnier AS t"
    stmt = stmt & Chr(13) & " WHERE m.TP_ID = p.TP_ID And p.turniernr = t.Turniernum And r.rt_id = m.rt_id and m.rt_id_weiter is null"
    stmt = stmt & Chr(13) & " AND m.RT_ID=" & RT_ID
    Set rstWR = dbs.OpenRecordset(stmt)
    Set rstZiel = dbs.OpenRecordset("Report_Platzierte_Paare")
    
    Do While (Not rstWR.EOF)
        rstZiel.AddNew
        
        rstZiel!Turniernr = rstWR!Turniernum
        rstZiel!Startkl = rstWR!Startkl
        rstZiel!Startnr = rstWR!Startnr
        rstZiel!Name = rstWR!Name
        rstZiel!Startklasse_text = rstWR!Startklasse_text
        rstZiel!Platz = rstWR!Platz
        rstZiel!Turnier_Name = rstWR!Turnier_Name
        rstZiel!T_Datum = rstWR!T_Datum
        rstZiel!Veranst_Name = rstWR!Veranst_Name
        rstZiel!Veranst_Ort = rstWR!Veranst_Ort
        rstZiel!Rundentext = rstWR!Rundentext
        rstZiel!Verein_Name = rstWR!Verein_Name
        rstZiel!disqualifiziert = (rstWR!disqualifiziert > 0)
        rstZiel!R_NAME_ABLAUF = rstWR!R_NAME_ABLAUF
        rstZiel!Majoritaet = rstWR!MajoritaetKurz
        rstZiel!Platz_WR = rstWR!PlatzWR
        rstZiel!Punktabzug = rstWR!Punktabzug
        rstZiel!Punktabzug_Anmerkung = rstWR!Anmerkung
        
        rstZiel!RT_ID = RT_ID
        rstZiel!Name_WR = WR_Name
        rstZiel!WR_Kurz = WR_Kurz
        
        rstZiel.Update
        rstWR.MoveNext
    Loop
    rstZiel.Close
    rstWR.Close
End Sub
