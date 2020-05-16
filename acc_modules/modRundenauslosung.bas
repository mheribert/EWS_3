Option Compare Database

Function Rundenauslosung(RT_ID, Anz_Paare)
    Dim dbs As Database
    Dim rstauslosung, rstErsatzrunde, rs As Recordset
    Set dbs = CurrentDb
   
    '  Anfang
    '  verhindern, dass mehrere Paare aus dem gleichen Verein in der gleichen Runde tanzen
    
    Dim found, abort As Boolean
    found = True
    abort = False
    Dim Count As Integer
    Count = 1
    Do While ((found) And (Not abort) And (Count < 50)) ' Wiederholen, bis keine Runde mehr mit einem doppelten Verein gefunden wurde
        Dim sqlString As String
        Dim runde1, runde2 As Integer
        Dim verein1 As String
        Dim tpid1, tpid2 As Integer
        
        Dim kritisch As Boolean
        If (Count > 40) Then
            kritisch = True
        End If
        
        sqlString = "SELECT Rundennummer, Verein_Name FROM Paare_Rundenqualifikation where rt_id=" & RT_ID & " and rundennummer is not null"
        sqlString = sqlString + " group by Verein_Name, Rundennummer having count(*)>1"
        
        Set rstauslosung = dbs.OpenRecordset(sqlString)
        
        If rstauslosung.EOF() Then
           found = False
           rstauslosung.Close
        Else
            rstauslosung.MoveFirst
            verein1 = rstauslosung!Verein_Name
            runde1 = rstauslosung!Rundennummer
            rstauslosung.Close
            
            verein1 = Replace(verein1, "'", "''")
            
            ' Hole erstes Tanzpaar von diesem Verein
            sqlString = "Select TP_ID from Paare_Rundenqualifikation where rt_id=" & RT_ID & " and Rundennummer=" & runde1 & " and Verein_Name='" & verein1 & "'"
            Set rs = dbs.OpenRecordset(sqlString)
            
            rs.MoveFirst
            tpid1 = rs!TP_ID
            rs.Close
            
            ' Suche Runde, in der von diesem Verein kein Paar tanzt
            sqlString = "SELECT distinct Rundennummer FROM Paare_Rundenqualifikation pr1 where pr1.rundennummer is not null and pr1.rt_id=" & RT_ID
            sqlString = sqlString & " and not exists (Select 1 from Paare_Rundenqualifikation pr2 where pr2.rt_id=pr1.rt_id and pr2.Verein_Name='" & verein1 & "' and pr2.rundennummer=pr1.rundennummer and pr2.rundennummer is not null)"
            
            Set rstErsatzrunde = dbs.OpenRecordset(sqlString)
            
            If (rstErsatzrunde.EOF()) Then
                abort = True
                Call rundenauslosung2(RT_ID, Anz_Paare)
                MsgBox "In einer oder mehreren Runden starten mehrere Paare eines Vereins zusammen, da ein Verein mehr Paare stellt als Runden getanzt werden."
                rstErsatzrunde.Close
            Else
                runde2 = rstErsatzrunde!Rundennummer
                rstErsatzrunde.Close
                
                ' Suche noch eine Startnummer aus der Ersatzrunde
                sqlString = "Select TP_ID from Paare_Rundenqualifikation where rt_id=" & RT_ID & " and rundennummer=" & runde2
                
                Set rs = dbs.OpenRecordset(sqlString)
                rs.MoveFirst
                tpid2 = rs!TP_ID
                rs.Close
                
                ' Tausche nun die beiden Paare in den gefundenen Runden
                
                sqlString = "Update Paare_Rundenqualifikation set rundennummer=" & runde2
                sqlString = sqlString & " where rt_id=" & RT_ID & " and TP_ID=" & tpid1
                
                dbs.Execute (sqlString)
                
                sqlString = "Update Paare_Rundenqualifikation set rundennummer=" & runde1
                sqlString = sqlString & " where rt_id=" & RT_ID & " and TP_ID=" & tpid2
                
                dbs.Execute (sqlString)
                
            End If
       
        End If
        Count = Count + 1
    Loop
    
    If (Count >= 50) Then
        Call rundenauslosung2(RT_ID, Anz_Paare)
        MsgBox "Es konnte nicht sichergestellt werden, dass Paare aus einem Verein nicht zusammen gelost wurden."
    End If
    
    dbs.Close
    
End Function

' -------------------------------------------------------------------------------------
' Alternative Rundenauslosung, wenn ein Verein zu viele Paare am Start hat, so dass
' von diesem Verein mehr Paare am Start sind, als Runden getanzt werden
' -------------------------------------------------------------------------------------
Function rundenauslosung2(RT_ID, Anz_Paare)
    
    Dim dbs As Database
    Set dbs = CurrentDb
    Dim rs As Recordset
    
    Dim sqlString, sql2 As String
    sqlString = "SELECT pr1.Verein_Name, count(*) as Anzahl FROM Paare_Rundenqualifikation pr1 where rt_id=" & RT_ID
    sqlString = sqlString & " group by Verein_Name order by 2 desc"
    
    Set rs = dbs.OpenRecordset(sqlString)
    
    Dim grossVerein As String
    
    If (rs.EOF()) Then
        rs.Close
        dbs.Close
        Exit Function
    End If
    
    grossVerein = rs!Verein_Name
    
    rs.Close
    
    sqlString = "SELECT count(*) as Anzahl FROM Paare_Rundenqualifikation pr1 where pr1.rt_id=" & RT_ID
    sqlString = sqlString & " and pr1.anwesend_Status=1"
    
    Set rs = dbs.OpenRecordset(sqlString)
    Dim Paaranzahl As Integer
    
    Paaranzahl = rs!Anzahl
    rs.Close
    
    sqlString = "SELECT pr1.TP_ID, pr1.Verein_Name FROM Paare_Rundenqualifikation pr1 where pr1.rt_id=" & RT_ID
    sqlString = sqlString & " and pr1.anwesend_Status=1"
    sqlString = sqlString & " order by 2 desc"
    
    Set rs = dbs.OpenRecordset(sqlString)
    
    
    Dim maxRunde As Integer
    maxRunde = Int((Paaranzahl + Anz_Paare - 1) / Anz_Paare)
        
    Dim Runde, TP_ID As Integer
    Runde = 1
    
    Do While (Not rs.EOF())
        TP_ID = rs!TP_ID
        
        rs.MoveNext
        Runde = (Runde Mod maxRunde) + 1
        Dim updateStr As String
        
        updateStr = "UPDATE Paare_Rundenqualifikation pr1 "
        updateStr = updateStr & " Set TP_ID=" & TP_ID
        updateStr = updateStr & " where pr1.rt_id=" & RT_ID
        
        dbs.Execute (updateStr)
    Loop
    
    rs.Close
    dbs.Close
        
End Function

Public Sub UpdateRundenqualifikation(RT_ID_Quelle As Integer, RT_ID_Ziel As Integer, AuslosungUebernehmen As Boolean)
' -----------------------------------------------------------------
' Aktualisiert eine Rundeneinteilung nach Vorlage von einer anderen
' -----------------------------------------------------------------
    Dim dbs As Database
    Set dbs = CurrentDb
    Dim rst_quelle As Recordset
    Dim rst_ziel As Recordset
    Dim TP_ID As Integer
    ' 1. Alle Tanzpaare löschen, die nicht in der Quelle enthalten sind
    Set rst_ziel = dbs.OpenRecordset("Select * from Paare_Rundenqualifikation where rt_id=" & RT_ID_Ziel)
    Set rst_quelle = dbs.OpenRecordset("Select * from Paare_Rundenqualifikation where rt_id=" & RT_ID_Quelle)
    Do While (Not rst_ziel.EOF)
        TP_ID = rst_ziel!TP_ID
        
        rst_quelle.FindFirst ("TP_ID=" & TP_ID)
        If (rst_quelle.NoMatch) Then
            rst_ziel.Delete
        End If
        rst_ziel.MoveNext
    Loop
    ' 2. Alle Tanzpaare hinzufügen, die noch fehlen
    If rst_quelle.RecordCount = 0 Then Exit Sub
    rst_quelle.MoveFirst
    Do While (Not rst_quelle.EOF)
        TP_ID = rst_quelle!TP_ID
        
        rst_ziel.FindFirst ("TP_ID=" & TP_ID)
        If (rst_ziel.NoMatch) Then
            rst_ziel.AddNew
            rst_ziel!TP_ID = TP_ID
            rst_ziel!RT_ID = RT_ID_Ziel
            rst_ziel!Verein_Name = rst_quelle!Verein_Name
            rst_ziel.Update
        End If
        
        rst_quelle.MoveNext
    Loop
    
    ' 3. Rundeneinteilung aktualisieren
    rst_quelle.MoveFirst
    Do While (Not rst_quelle.EOF)
        TP_ID = rst_quelle!TP_ID
        
        rst_ziel.FindFirst ("TP_ID=" & TP_ID)
        
        rst_ziel.Edit
        If (AuslosungUebernehmen) Then
            rst_ziel!Rundennummer = rst_quelle!Rundennummer
            rst_ziel!Auslosung = rst_quelle!Auslosung
        End If
        rst_ziel!Anwesend_Status = rst_quelle!Anwesend_Status
        rst_ziel.Update
        
        rst_quelle.MoveNext
    Loop
    
End Sub
