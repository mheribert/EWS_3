Option Compare Database
Option Explicit

Public Sub Geteilte_Endrunde_Fuellen(RT_ID As Integer, RT_ID_FT As Integer, RT_ID_AK As Integer, anz As Integer, fact)
    ' HM 20.02.1012
    ' umgestellt auf factor
    ' automatische Auswertung der geteilten Endrunde -> ohne Regelverstoß und Disqualifikation
    Dim dbs As Database
    Set dbs = CurrentDb
    
    Dim rstft As Recordset
    Dim rstak As Recordset
    Dim rstABER As Recordset
    Dim stmt As String
    Dim RegelV_FT, RegelV_AK As Integer
    Dim DQ_ID As Integer
    Dim x_wr As Integer
    
    'bisherige Datensätze aus der Tabelle geteilte_Endrunde löschen wenn gleiche RT_ID (Klasse)
    dbs.Execute ("Delete from geteiltes_Endergebnis where rt_id=" & RT_ID)
    
    'Datensätze FT in Recordset
    Set rstft = dbs.OpenRecordset("select * from Majoritaet where rt_id=" & RT_ID_FT & " ORDER BY TP_ID")
    
    'Datensätze Akro in Recordset
    Set rstak = dbs.OpenRecordset("select * from Majoritaet where rt_id=" & RT_ID_AK & " ORDER BY TP_ID")
    
    'Tabelle geteiltes_Endergebnis in Recordset
    Set rstABER = dbs.OpenRecordset("geteiltes_Endergebnis")
    
    ' schauen ob alle Endrunden ausgewertet sind, wenn nicht -> machen, mit Hinweis
    If rstft.RecordCount = 0 Or rstak.RecordCount = 0 Then
        MsgBox ("Es wurden die geteilten Endrunden nicht einzeln ausgewertet. Wenn es Regelverstöße gegeben hat, müssen diese in der Einzelauswertung eingetragen werden. Ohne Regelverstöße erfolgt eine korrekte Gesamtauswertung.")
        If rstft.RecordCount = 0 Then
            Call msystem(RT_ID_FT, "", "", "End_r", anz, True)  'da ausgewertet wurde, neu lesen
            Set rstft = dbs.OpenRecordset("select * from Majoritaet where rt_id=" & RT_ID_FT & " ORDER BY TP_ID")
        End If
        If rstak.RecordCount = 0 Then
            Call msystem(RT_ID_AK, "", "", "End_r", anz, True)  'da ausgewertet wurde, neu lesen
            Set rstak = dbs.OpenRecordset("select * from Majoritaet where rt_id=" & RT_ID_AK & " ORDER BY TP_ID")
        End If
    End If
    
    Do While (Not rstft.EOF)
        If Not rstak!TP_ID = rstft!TP_ID Then
            MsgBox ("Hier stimmt was mit der Reihenfolge in der FT (" & rstft!TP_ID & " / Akro " & rstak!TP_ID & " Majorität nicht")
        End If
        RegelV_FT = rstft!PA_ID
        DQ_ID = rstft!DQ_ID
        RegelV_AK = rstak!PA_ID
        If rstak!DQ_ID > 0 Then
           If DQ_ID = 0 Then
                    'Verstoß nur übernehmen wenn noch keiner eingetragen wurde
              DQ_ID = rstak!DQ_ID
           End If
        End If
        For x_wr = 1 To anz
            rstABER.AddNew
            rstABER!RT_ID = RT_ID
            rstABER!WR_ID = rstft("wr" & x_wr & "_id")
            rstABER!TP_ID = rstft!TP_ID
            rstABER!Punkte_FT = rstft("wr" & x_wr & "_Punkte")
            rstABER!RegelV_FT = RegelV_FT
            rstABER!Platz_ft = rstft("wr" & x_wr & "_Platz")
            
            rstABER!Punkte_AK = rstak("wr" & x_wr & "_Punkte")
            rstABER!RegelV_AK = RegelV_AK
            rstABER!Platz_ak = rstak("wr" & x_wr & "_Platz")
            rstABER!Platz_Summe = rstABER!Platz_ft + (rstABER!Platz_ak * fact)
            rstABER!DQ_ID = DQ_ID
            rstABER.Update
        Next x_wr
        rstft.MoveNext
        rstak.MoveNext
    Loop
    rstft.Close
    rstABER.Close
    rstak.Close
    
    ' Und jetzt noch die Gesamtplatzierungen pro WR eintragen
    Dim rstGER As Recordset
    Dim Platz As Integer
    Dim currWR_ID As Integer
    currWR_ID = -1
    
    ' Gesamtplatz vergeben, Daten aus geteiltes_Endergebnis holen und nach Platz_Summe un Punkte_AK sortieren
    ' durch Sortierung ist keine Bonusverteilung nötig!
    
    stmt = "select * from geteiltes_Endergebnis where rt_id=" & RT_ID & " order by wr_id, platz_summe, Platz_ak"
    Set rstGER = dbs.OpenRecordset(stmt)

    Do While (Not rstGER.EOF)
        If (currWR_ID <> rstGER!WR_ID) Then
            Platz = 1
            currWR_ID = rstGER!WR_ID
        End If
        
        rstGER.Edit
        
        rstGER!gesamt_Platz = Platz
        
        rstGER.Update
        
        rstGER.MoveNext
        Platz = Platz + 1
    Loop
    rstGER.Close
    
    ' Alle Einträge in der Tabelle Auswertung zu dieser Runde löschen
    stmt = "delete FROM Auswertung a where exists (select 1 from Paare_Rundenqualifikation pr where pr.pr_id=a.pr_id and pr.rt_id=" & RT_ID & ")"
    dbs.Execute (stmt)
    
    ' Einträge aus der geteilten Endrunde in die Tabelle Auswertung übernehmen
    Dim rst As Recordset
    Dim rstPaareRundenquali As Recordset
    stmt = "select RT_ID, ger.WR_ID, TP_ID, gesamt_platz, Platz_Summe from geteiltes_Endergebnis ger, Wert_Richter wr where ger.wr_id=wr.wr_id and rt_id=" & RT_ID & " order by tp_id, wr_kuerzel"
    Set rstGER = dbs.OpenRecordset(stmt)
    Set rst = dbs.OpenRecordset("Auswertung")
    Set rstPaareRundenquali = dbs.OpenRecordset("Paare_Rundenqualifikation")
    
    Do While (Not rstGER.EOF)
    '*** für den folgenden Teil fehlen Einträge in der Tabelle Paare_Rundenqualifikation, diese müssen wohl noch angelegt werden
    '*** da sonst keine Werte in die Auswertungstabelle eingetragen werden, und es keine Auswertung gibt!!!
        rstPaareRundenquali.FindFirst ("RT_ID=" & RT_ID & " and TP_ID=" & rstGER!TP_ID)
        
        If (Not rstPaareRundenquali.NoMatch) Then
            rst.AddNew
            rst!PR_ID = rstPaareRundenquali!PR_ID
            rst!WR_ID = rstGER!WR_ID
            rst!Punkte = 80 - rstGER!Platz_Summe
            rst!Platz = rstGER!gesamt_Platz
            rst.Update
        End If
        rstGER.MoveNext
    Loop
    
    rstGER.Close
    rstPaareRundenquali.Close
    rst.Close
End Sub
