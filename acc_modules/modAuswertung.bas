Option Compare Database
Option Explicit                       ' alle Variablen muessen deklariert werden
    Dim db As Database
    
Public Sub RR_Auswertung(rt, TNR, ft_id, st_kl)
    Dim ft_ak As Recordset
    Dim pr As Recordset
    Dim maj As Recordset
    Dim ft_pu, ak_pu, ftft_pu, bs_pu, bw_pu, bw_verst, ges_Punkte As Double
    Dim sqlstm As String
    Dim Runde As String
    Dim is_wertung As Boolean
    
    Set db = CurrentDb
    
    Set pr = db.OpenRecordset("SELECT * FROM rundentab WHERE RT_ID=" & rt & ";")
    Runde = pr!Runde
    Set pr = db.OpenRecordset("SELECT * FROM Paare_Rundenqualifikation WHERE RT_ID=" & rt & " AND Anwesend_Status=1;")
    
    pr.MoveFirst
    Do Until pr.EOF
        is_wertung = False
        ft_pu = 0
        ak_pu = 0
        ftft_pu = 0
        bw_pu = 0
        bw_verst = 0
        sqlstm = "SELECT Wert_Richter.WR_Kuerzel, Auswertung.* FROM Startklasse_Wertungsrichter INNER JOIN (Wert_Richter INNER JOIN Auswertung ON Wert_Richter.WR_ID = Auswertung.WR_ID) ON Startklasse_Wertungsrichter.WR_ID = Wert_Richter.WR_ID " & _
                 "WHERE (Turniernr=" & TNR & " AND Auswertung.PR_ID=" & pr!PR_ID & " AND WR_function ='Ft' AND WR_Azubi=False AND Startklasse='" & st_kl & "') ORDER BY Auswertung.PR_ID, Auswertung.Punkte;"
        Set ft_ak = db.OpenRecordset(sqlstm) 'Fußtechnik Runde
        If Not ft_ak.EOF Then
            ft_pu = get_mittel(ft_ak)
            is_wertung = True
        End If
    
        sqlstm = "SELECT Wert_Richter.WR_Kuerzel, Auswertung.* FROM Startklasse_Wertungsrichter INNER JOIN (Wert_Richter INNER JOIN Auswertung ON Wert_Richter.WR_ID = Auswertung.WR_ID) ON Startklasse_Wertungsrichter.WR_ID = Wert_Richter.WR_ID " & _
                 "WHERE (Turniernr=" & TNR & " AND Auswertung.PR_ID=" & pr!PR_ID & " AND WR_function ='Ak' AND WR_Azubi=False AND Startklasse='" & st_kl & "') ORDER BY Auswertung.PR_ID, Auswertung.Punkte;"
        Set ft_ak = db.OpenRecordset(sqlstm) ' Akrorunde
        If Not ft_ak.EOF Then
            ak_pu = get_mittel(ft_ak)
            is_wertung = True
        End If
        
        sqlstm = "SELECT Wert_Richter.WR_Kuerzel, Auswertung.* FROM Startklasse_Wertungsrichter INNER JOIN (Wert_Richter INNER JOIN Auswertung ON Wert_Richter.WR_ID = Auswertung.WR_ID) ON Startklasse_Wertungsrichter.WR_ID = Wert_Richter.WR_ID " & _
                 "WHERE (Turniernr=" & TNR & " AND Auswertung.PR_ID=" & pr!PR_ID & " AND WR_function ='X' AND WR_Azubi=False AND Startklasse='" & st_kl & "') ORDER BY Auswertung.PR_ID, Auswertung.Punkte;"
        Set ft_ak = db.OpenRecordset(sqlstm)    'NewJudgingSystem
        If Not ft_ak.EOF Then
            bw_pu = get_mittel(ft_ak)
            If InStr(1, Runde, "schnell") > 0 Then bw_pu = bw_pu * 1.1
            is_wertung = True
        End If
        
        sqlstm = "SELECT Wert_Richter.WR_Kuerzel, Auswertung.* FROM Startklasse_Wertungsrichter INNER JOIN (Wert_Richter INNER JOIN Auswertung ON Wert_Richter.WR_ID = Auswertung.WR_ID) ON Startklasse_Wertungsrichter.WR_ID = Wert_Richter.WR_ID " & _
                 "WHERE (Turniernr=" & TNR & " AND Auswertung.PR_ID=" & pr!PR_ID & " AND WR_function ='Ob' AND WR_Azubi=False AND Startklasse='" & st_kl & "') ORDER BY Auswertung.PR_ID, Auswertung.Punkte;"
        Set ft_ak = db.OpenRecordset(sqlstm)    ' Observerabzüge NewJudgingSystem
        If Not ft_ak.EOF Then
            bw_verst = IIf(Nz(ft_ak!Punkte) = "", 0, ft_ak!Punkte)
        End If
        
        If ft_id > 0 Then   ' falls geteilte Endr 1.Runde holen
            sqlstm = "SELECT Majoritaet.* FROM Paare_Rundenqualifikation INNER JOIN Majoritaet ON Paare_Rundenqualifikation.TP_ID = Majoritaet.TP_ID" & _
                     " WHERE Paare_Rundenqualifikation.PR_ID=" & pr!PR_ID & " AND Majoritaet.RT_ID=" & ft_id & ";"
            Set ft_ak = db.OpenRecordset(sqlstm)
            If ft_ak.RecordCount = 1 Then
                ftft_pu = ft_ak!WR7
            Else
                MsgBox "Fehler bei zweigeteilter Runde holen! 1. Runde nochmal auswerten."
                Exit Sub
            End If
        End If
        
        If is_wertung Then
            Set maj = db.OpenRecordset("SELECT * FROM Majoritaet WHERE RT_ID=" & rt & " AND tp_id=" & pr!TP_ID & ";")
            If maj.EOF Then
                maj.AddNew
                maj!RT_ID = rt
                maj!TP_ID = pr!TP_ID
            Else
                maj.Edit
            End If
            maj!WR1_Punkte = IIf(ft_id < 0, 0, ftft_pu)
            maj!WR2_Punkte = ft_pu
            maj!WR3_Punkte = ak_pu
            maj!WR5_Punkte = bw_pu
            maj!WR6_Punkte = bw_verst
            
            maj!WR1 = FormatNumber(IIf(ft_id < 0, 0, ftft_pu), 2)
            maj!WR2 = Format(ft_pu, "##0.00")
            maj!WR3 = Format(ak_pu, "##0.00")
            maj!WR5 = Format(bw_pu, "##0.00")
            maj!WR6 = Format(bw_verst, "-##0.00")
            
            ges_Punkte = ftft_pu + IIf(ft_pu + ak_pu < 0, 0, ft_pu + ak_pu) + bw_pu - bw_verst - maj!PA_ID * get_verstoss(st_kl)
            If ges_Punkte < 0 Then ges_Punkte = 0
            maj!WR7 = Format(ges_Punkte, "##0.00")
            maj!WR7_Punkte = ges_Punkte
            maj.Update
        End If
        pr.MoveNext
    Loop
    If Runde = "KO_r" Then
        Call RR_KO_Sieger_ermitteln(rt)
    End If
    Call RR_platz_vergeben(rt)
End Sub

Public Sub RR_Punkteabzug(rt As Integer, stkl As String, TP As Integer, Anzahl_Abzuege As Integer, Runde As String)
    Dim stmt As String
    Dim maj As Recordset
    
    Set db = CurrentDb
    stmt = "Select * from majoritaet where rt_id=" & rt & " and tp_id=" & TP
    Set maj = db.OpenRecordset(stmt)
    maj.Edit
    '*****AB***** V13.06 WR1 fehlte in der Addition, daher gab es zu wenig Punkte bei Regelverstoß, hier eingefügt
    maj!WR7 = Format(maj!WR1 + maj!WR2 + maj!WR3 - (get_verstoss(stkl) * Anzahl_Abzuege), "##0.00")
    If maj!WR7 < 0 Then maj!WR7 = 0
    maj.Update
    
    If Runde = "KO_r" Then
        Call RR_KO_Sieger_ermitteln(rt)
    End If
    Call RR_platz_vergeben(rt)
End Sub

Public Sub RR_platz_vergeben(rt)
    Set db = CurrentDb
    Dim maj As Recordset
    Dim pl, pl_m, pl_a As Integer
    
    Set maj = db.OpenRecordset("SELECT * FROM Majoritaet WHERE RT_ID=" & rt & " ORDER BY KO_Sieger, DQ_ID, WR7 DESC;")
    If maj.RecordCount = 0 Then
        MsgBox "Es gibt noch keine Wertungen in dieser Tanzrunde!"
    Else
        maj.MoveLast
        maj.MoveFirst
        pl = 0
        pl_m = 0
        pl_a = 0
        Do Until maj.EOF
            maj.Edit
            If pl_m = maj!WR7 Then
                pl_a = pl_a + 1
                maj!Platz = pl
                maj!Platz_Orig = pl
            Else
                pl = pl + 1 + pl_a
                pl_m = maj!WR7
'                If maj!DQ_ID = 0 Then
                    maj!Platz = pl
'                    If maj!Platz_Orig = 0 Then maj!Platz_Orig = pl
'                Else
'                    maj!Platz = maj.RecordCount
'                    If maj!Platz_Orig = 0 Then maj!Platz_Orig = maj.RecordCount
'                End If
                pl_a = 0
            End If
            maj.Update
            maj.MoveNext
        Loop
    End If
End Sub

Public Sub RR_KO_Sieger_ermitteln(rt)
    '*****AB***** V14.02 neue Funktion zum Auswerten von KO Runden - Sieger der Runde ermitteln
    Set db = CurrentDb
    Dim maj As Recordset
    Dim Punkte_Paar1, Punkte_Paar2 As Double
    Dim Runde_Paar1 As Integer
    Set maj = db.OpenRecordset("SELECT Paare_Rundenqualifikation.*, * FROM Majoritaet INNER JOIN Paare_Rundenqualifikation ON (Majoritaet.TP_ID = Paare_Rundenqualifikation.TP_ID) AND (Majoritaet.RT_ID = Paare_Rundenqualifikation.RT_ID) WHERE (((Majoritaet.[RT_ID])=" & rt & ")) ORDER BY Paare_Rundenqualifikation.Rundennummer, Majoritaet.DQ_ID, Majoritaet.WR7 DESC;")
    maj.MoveLast
    maj.MoveFirst
    
    Punkte_Paar1 = 0
    Punkte_Paar2 = 0
    Runde_Paar1 = 0
    
    Do Until maj.EOF
        Punkte_Paar1 = maj!WR7
        Runde_Paar1 = maj!Rundennummer
        maj.MoveNext
        ' Falls ungerade Anzahl Paare in der Runde wird das letzte Paar als Sieger gesetzt
        If maj.EOF Then
                maj.MovePrevious
                maj.Edit
                maj![Majoritaet.Ko_Sieger] = True
                maj.Update
        Else
            If Runde_Paar1 = maj!Rundennummer Then
                If Punkte_Paar1 > maj!WR7 Then
                    maj.MovePrevious
                    maj.Edit
                    maj![Majoritaet.Ko_Sieger] = True
                    maj.Update
                    maj.MoveNext
                    maj.Edit
                    maj![Majoritaet.Ko_Sieger] = False
                    maj.Update
                Else
                    maj.MovePrevious
                    maj.Edit
                    maj![Majoritaet.Ko_Sieger] = False
                    maj.Update
                    maj.MoveNext
                    maj.Edit
                    maj![Majoritaet.Ko_Sieger] = True
                    maj.Update
                End If
            End If
        End If
        maj.MoveNext
    Loop
End Sub

' Neu wegen möglicher 6 oder 8 FT-WR
Function get_mittel(avr)
    Dim i(8) As Double
    Dim min As Integer
    Dim max As Integer
    Dim X As Integer
    
    avr.MoveLast
    For X = 1 To avr.RecordCount
        i(X) = Nz(avr!Punkte)
        avr.MovePrevious
    Next
    
    Select Case avr.RecordCount
        Case 2, 3
            min = 1
            max = avr.RecordCount
        Case 4
            min = 2
            max = 3
        Case 5
            min = 2
            max = 4
        Case 6
            min = 2
            max = 5
        Case 7
            min = 2
            max = 6
        Case 8
            min = 3
            max = 6
        Case Else
            MsgBox "Fehler in der Anzahl der WR!"
    End Select
    If Forms![A-Programmübersicht]!Turnierausw.Column(8) = "SL" Then
        min = 1
        max = avr.RecordCount
    End If
    For X = min To max
        i(0) = i(0) + i(X)
    Next
    get_mittel = i(0) / (max - min + 1)

End Function

Function get_verstoss(stkl)
    Dim rst1 As Recordset
    Dim stmt As String
    Set db = CurrentDb
    
    stmt = "Select * from startklasse where startklasse='" & stkl & "'"
    
    Set rst1 = db.OpenRecordset(stmt)
    rst1.MoveFirst
    get_verstoss = rst1!AbzugTSOVerstoss
    rst1.Close
End Function

Public Sub AuswertenundPlatzieren(StartklasseID As Integer, Startkl As String, AnzahlWR As Integer, Runde As String, IsEndrunde As Integer)
    ' Test, ob Majorität schon vorhanden, ggfs. löschen
    Dim dbs As Database
    Set dbs = CurrentDb
    Dim rst As Recordset
    Dim stmt As String
    Dim Count As Integer
    Dim ft_rt_id As Integer
    Dim result As Integer
    Dim stDocName As String
    
    stmt = "Select count(*) as anz from Majoritaet where rt_id=" & StartklasseID
    Set rst = dbs.OpenRecordset(stmt)
    Count = rst!anz
    rst.Close
        
    Dim Turniernr As Integer
    Turniernr = get_aktTNr
    
    Dim RT_ID_Teil1 As Integer
    Dim RT_ID_Teil2 As Integer
    Dim fact As Double

            If Runde = "End_r_Akro" Then
                ft_rt_id = check_1_Runde_von_2(Turniernr, Startkl, "End_r_Fuß")
                Call RR_Auswertung(StartklasseID, Turniernr, ft_rt_id, Startkl)
                
            ElseIf Runde = "Vor_r_schnell" Then
                ft_rt_id = check_1_Runde_von_2(Turniernr, Startkl, "Vor_r_lang")
                Call RR_Auswertung(StartklasseID, Turniernr, ft_rt_id, Startkl)
                
            ElseIf Runde = "End_r_schnell" Then
                ft_rt_id = check_1_Runde_von_2(Turniernr, Startkl, "End_r_lang")
                Call RR_Auswertung(StartklasseID, Turniernr, ft_rt_id, Startkl)
            
            ElseIf Runde = "End_r_2" Then
                ft_rt_id = check_1_Runde_von_2(Turniernr, Startkl, "End_r_1")
                Call RR_Auswertung(StartklasseID, Turniernr, ft_rt_id, Startkl)
            
            Else
'                If Left(Startkl, 3) = "BS_" Then
'                    Call msystem(StartklasseID, Turniernr, Startkl, Runde, AnzahlWR, False)
'                Else
                    Call RR_Auswertung(StartklasseID, Turniernr, 0, Startkl)
'                End If
            End If
'        Else
'            '*****AB***** V14.02 kompletten Block verschoben, da bei RR_Auswertung die Tabelleaktualisiert wird
'            If (count > 0) Then
'                result = MsgBox("Es besteht schon eine Auswertung. Wollen Sie die Runde neu errechnen?", vbYesNo)
'                If (result = vbNo) Then
'                    Exit Sub
'                Else
'                    stmt = "Delete from Majoritaet where rt_id=" & StartklasseID
'                    dbs.Execute (stmt)
'                End If
'            End If
'            '*****AB***** V14.02 Block ENDE
'
'            Call msystem(StartklasseID, Turniernr, Startkl, Runde, AnzahlWR, False)
'        End If
'    End If
    
    ' Wenn Endrunde erreicht, dann die Platzierung sofort aufrufen; ansonsten werden die Paare in Paare weiter nehmen platziert 16.6.04 HK
    If IsEndrunde = 1 Then
        Call PaarePlatzieren(StartklasseID, 1)
        Call WriteRundeReport(StartklasseID)
    End If
    
    Call UpdateAnzahl_Paare(StartklasseID)

    If Runde = "End_r" Or Runde = "End_r_Akro" Or Runde = "End_r_schnell" Or Runde = "End_r_2" Then
        make_a_siegerehrung StartklasseID    ' HTML-Moderation
    End If
    
End Sub

Function check_1_Runde_von_2(Turniernr As Integer, Startkl As String, Runde As String)
    Dim re As Recordset
    Set db = CurrentDb
    check_1_Runde_von_2 = getRT_ID(Turniernr, Startkl, Runde)
    Set re = db.OpenRecordset("SELECT Count(*) as anzahl FROM Majoritaet WHERE RT_ID=" & check_1_Runde_von_2 & ";")
    If re!Anzahl = 0 Then
        Call RR_Auswertung(check_1_Runde_von_2, Turniernr, 0, Startkl)
    End If
End Function

