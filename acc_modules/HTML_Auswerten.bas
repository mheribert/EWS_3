Option Compare Database
Option Explicit

    Public Type Formationswerte
      faktor    As Single           ' Faktor f�r Reduzierung bei Berechnung
      min       As Integer          ' minimum Anzahl T�nzer
      max       As Integer          ' maximum Anzahl T�nzer
    End Type

    Dim db As Database
    Dim fs, inp

Public Function get_wertungen(rt, s_kl, rde)  'Aus Server Input-Datei einlesen
    Dim re, pr, ana As Recordset
    Dim fs, inp
    Dim rh As Integer
    Dim fName As String
    Dim sqlcmd As String
    Dim strEWS2 As String
    Dim ft_rt As Integer
    Dim cgivar, WR_func
    Set db = CurrentDb
    
    WR_func = gen_wr_arr(s_kl)
    
    fName = getBaseDir & get_TerNr & "_RT" & CInt(rt) & "_raw.txt"   ' Dateiname erstellen
    Set fs = CreateObject("Scripting.FileSystemObject")
    If get_properties("EWS") = "EWS3" And Len(Dir(fName)) > 0 Then  ' bei EWS3 die rawWerte einlesen
        Set inp = fs.OpenTextFile(getBaseDir & get_TerNr & "_RT" & CInt(rt) & "_raw.txt", 1, 0)  'reading
        Set ana = db.OpenRecordset("Analyse", DB_OPEN_DYNASET)
        Do Until inp.AtEndOfStream
            cgivar = Split(inp.Readline, ";")
            
            ana.FindFirst "TP_ID = " & cgivar(0) & " And WR_ID = " & cgivar(1) & " And Cgi_Input = '" & cgivar(2) & "'"
            If ana.NoMatch Then
                ana.AddNew
                ana!TP_ID = cgivar(0)
                ana!WR_ID = cgivar(1)
                ana!RT_ID = rt
                ana!Cgi_Input = cgivar(2)
                ana.Update
            End If
            
        Loop
    End If
    
    If InStr(1, rde, "_Akro") Then      ' Sicherstellen dass Ft-Runde da ist
        Set re = db.OpenRecordset("SELECT * from RundenTab WHERE Startklasse = '" & s_kl & "' AND Runde = '" & left(rde, 3) & "_r_Fu�';", DB_OPEN_DYNASET)
        ft_rt = re!RT_ID
        get_wertungen re!RT_ID, s_kl, re!Runde      'Fu�technik rekursiv einlesen
    End If
    fName = getBaseDir & get_TerNr & "_RT" & CInt(rt) & ".txt"   ' Dateiname erstellen

    If get_properties("EWS") = "EWS2" And left(s_kl, 3) = "RR_" Then    ' Bei EWS2 RT-Datei herunterladen
        DoCmd.Hourglass True
        get_url_to_file "http://" & get_properties("EWS20_Adresse") & "/" & get_TerNr & "_RT" & CInt(rt) & ".txt", fName
        DoCmd.Hourglass False
    End If

    If Len(Dir(fName)) > 0 Then
        Set inp = fs.OpenTextFile(fName, 1, 0)  'reading
        Do Until inp.AtEndOfStream
            cgivar = Split(inp.Readline, ";")
            Set pr = db.OpenRecordset("SELECT PR_ID FROM Paare_Rundenqualifikation WHERE TP_ID=" & cgivar(0) & " AND RT_ID=" & rt & ";")
            If pr.RecordCount <> 0 Then
                Set re = db.OpenRecordset("SELECT * FROM Auswertung WHERE pr_id =" & pr!PR_ID & " AND wr_id =" & cgivar(1) & ";")
                If re.EOF Then  'Abfrage vorhanden, wenn nicht neu
                    re.AddNew
                    re!PR_ID = pr!PR_ID
                    re!WR_ID = cgivar(1)
                Else            ' oder edit
                    re.Edit
                End If
                If cgivar(1) = 99 Then
                    re!Punkte = rechne_abzuege(cgivar(0), cgivar(2))
                Else
                    re!Punkte = rechne_punkte(cgivar(0), cgivar(2), s_kl, rh, rde, ft_rt, WR_func)   ' Punkte Klassen- und Rundenabh�ngig ausrechen
                End If
                re!Reihenfolge = rh
                re!Cgi_Input = cgivar(2)
                re!Platz = 0
                re.Update
            Else
                db.Execute "INSERT INTO Analyse (CGI_Input,zeit) VALUES ('Fehler in der RT-Datei " & rt & "', '" & Time & "')"
            End If
        Loop
    Else
        get_wertungen = True
    End If
    
    ' Wertungen l�schen, die nicht rein geh�ren z.B. unentschuldigt
    sqlcmd = "select distinct pr.pr_id from Paare_Rundenqualifikation pr, Auswertung a where a.pr_id=pr.pr_id and pr.rt_id=" & rt & " and anwesend_Status<>1;"
    Set re = db.OpenRecordset(sqlcmd)
    Do Until re.EOF
        sqlcmd = "Delete from Auswertung where pr_id=" & re!PR_ID
        db.Execute (sqlcmd)
        re.MoveNext
    Loop

End Function

Private Function gen_wr_arr(s_kl)
    Dim re As Recordset
    Set re = db.OpenRecordset("SELECT * FROM Startklasse_Wertungsrichter WHERE Startklasse ='" & s_kl & "' ORDER BY wr_id DESC;")
    Dim back
    ReDim back(re!WR_ID)
    re.MoveFirst
    Do Until re.EOF
        back(re!WR_ID) = re!WR_function
        re.MoveNext
    Loop
    gen_wr_arr = back
End Function

Private Function rechne_abzuege(PR_ID, inp)
    Dim vars
    Dim i, rh, x As Integer
    Dim Punkte As Double
    Dim verst
    verst = Array("beobachter_zukurz", "beobachter_zulang", "beobachter_Makeup", "beobachter_schmuck", "beobachter_requsit")

    Set vars = zerlege(inp)
    i = eins_zwei(PR_ID, vars)
    rh = vars.Item("rh" & i)
    
    For x = 0 To UBound(verst)
        If vars.Item(verst(x)) <> "" Then
            Punkte = Punkte + CSng(vars.Item(verst(x)))
        End If
    
    Next
    rechne_abzuege = Punkte
End Function

Private Function rechne_punkte(PR_ID, inp, s_kl, rh, rde, ft_rt, WR_func)
    'Punkte Klassen- und Rundenabh�ngig ausrechen
    Dim vars
    Dim i, t As Integer
    Dim Punkte As Double
    Dim ft_pte As Double
    Dim kl_punkte, flds As Variant
    
    If left(s_kl, 3) = "RR_" Then
        inp = Replace(inp, "_1=", "1=")
        inp = Replace(inp, "_2=", "2=")
    End If
    
    Set vars = zerlege(inp)
    i = eins_zwei(PR_ID, vars)
    rh = vars.Item("rh1")
    
    Select Case left(s_kl, 3)
    
        Case "F_B"
            If vars.Item("wtk1") <> "" Then
                kl_punkte = Punkteverteilung(s_kl, ch_runde(rde), rde)
                Punkte = Punkte + CSng(vars.Item("wtk1")) * kl_punkte(0) / 10 + CSng(vars.Item("wch1")) * kl_punkte(1) / 10
                Punkte = Punkte + CSng(vars.Item("wtf1")) * kl_punkte(2) / 10 + CSng(vars.Item("wab1")) * kl_punkte(4) / 10
                Punkte = Punkte + CSng(vars.Item("waw1")) * kl_punkte(5) / 10 + CSng(vars.Item("waf1")) * kl_punkte(6) / 10
            End If
        Case "F_R"
            If vars.Item("tfl" & i & "a20") <> "" Then
                Punkte = Punkte + CSng(vars.Item("wfl" & i & "a20"))
            Else
                If WR_func(vars.Item("WR_ID")) <> "Ob" Then
                    Punkte = add_akro(vars, i, t)
                    If t = 0 Then t = 1
                    '"wtk1", "wch1", "wtf1", "wab1", "waw1", "waf1" , "wak11", "wak12", "wak13", "wak14", "wak15", "wak16", "wak17", "wak18"
                    If s_kl = "F_RR_M" Then ' Master RR
        '                Punkte = Punkte / IIf(t < 7, 6, t) * 5
                    Else
                        Punkte = 0
                    End If
                    If vars.Item("wtk1") <> "" Then
                        kl_punkte = Punkteverteilung(s_kl, ch_runde(rde), rde)
                        Punkte = Punkte + CSng(vars.Item("wtk1")) * kl_punkte(0) / 10 + CSng(vars.Item("wch1")) * kl_punkte(1) / 10
                        Punkte = Punkte + CSng(vars.Item("wtf1")) * kl_punkte(2) / 10 + CSng(vars.Item("wab1")) * kl_punkte(4) / 10
                        Punkte = Punkte + CSng(vars.Item("waw1")) * kl_punkte(5) / 10 + CSng(vars.Item("waf1")) * kl_punkte(6) / 10
                    End If
                    Punkte = (Punkte * Form_abzuege(PR_ID, s_kl)) - Val(Nz(Replace(vars.Item("wfl1"), ".", ",")))
                End If
            End If
        Case "RR_"
            If vars.Item("tfl" & i & "a20") <> "" Then
                Punkte = Punkte + CSng(vars.Item("wfl" & i & "a20"))
            Else
                If WR_func(vars.Item("WR_ID")) = "Ob" Then
                    If Not vars.exists("Obs_check1") Then
                        Punkte = CSng(vars.Item("wmk_th" & i))
                    End If
                Else
                    Punkte = add_akro(vars, i, t)
                    If t = 0 Then t = 1     '  Wegen Runden in denen keine Akro ist
                    If InStr(1, rde, "_Fu�") > 0 Then Punkte = 0
                    If vars.Item("wsh" & i) <> "" Then
                        kl_punkte = Punkteverteilung(s_kl, ch_runde(rde), rde)
                        If ch_runde(rde) = "ER" Then
                            Punkte = Punkte + CSng(vars.Item("wsh" & i)) * kl_punkte(0) / 10 + CSng(vars.Item("wth" & i)) * kl_punkte(1) / 10
                            Punkte = Punkte + CSng(vars.Item("wsd" & i)) * kl_punkte(2) / 10 + CSng(vars.Item("wtd" & i)) * kl_punkte(3) / 10
                            Punkte = Punkte + CSng(vars.Item("wch" & i)) * kl_punkte(4) / 10 + CSng(vars.Item("wtf" & i)) * kl_punkte(5) / 10
                            Punkte = Punkte + CSng(vars.Item("wda" & i)) * kl_punkte(6) / 10
                        Else
                            Punkte = Punkte + CSng(vars.Item("wsh" & i)) * kl_punkte(0) / 10 * 2
                            Punkte = Punkte + CSng(vars.Item("wsd" & i)) * kl_punkte(2) / 10 * 2
                            Punkte = Punkte + CSng(vars.Item("wch" & i)) * kl_punkte(4) / 10 + CSng(vars.Item("wch" & i)) * kl_punkte(5) / 10
                            Punkte = Punkte + CSng(vars.Item("wch" & i)) * kl_punkte(6) / 10
                        End If
                        Punkte = Punkte - Replace(vars.Item("wfl" & i), ".", ",")
                    End If
                    If vars.Item("wmk_th" & i) <> "" Then
                        Punkte = Punkte + to_zahl(vars, "wmk_th" & i)
                        Punkte = Punkte + to_zahl(vars, "wmk_dh" & i)
                        Punkte = Punkte + to_zahl(vars, "wmk_td" & i)
                        Punkte = Punkte + to_zahl(vars, "wmk_dd" & i)
                    End If
                End If
            End If
        Case "BW_"
            If vars.exists("wng_tth" & i) Then
                Punkte = CSng(vars.Item("Punkte_err" & i))     ' newGuidelines         Kategorien-Streichverfahren
            Else
                kl_punkte = Punkteverteilung(s_kl, ch_runde(rde), rde)
                If ch_runde(rde) = "ER" Then
                    Punkte = Punkte + CSng(vars.Item("wgs" & i)) * kl_punkte(0) / 10 + CSng(vars.Item("wbd" & i)) * kl_punkte(1) / 10
                    Punkte = Punkte + CSng(vars.Item("wtf" & i)) * kl_punkte(2) / 10 + CSng(vars.Item("win" & i)) * kl_punkte(4) / 10
                    Punkte = Punkte + CSng(vars.Item("wsi" & i)) * kl_punkte(5) / 10 + CSng(vars.Item("wdp" & i)) * kl_punkte(6) / 10
                Else
                    Punkte = Punkte + CSng(vars.Item("wgs" & i)) * kl_punkte(0) / 10 + CSng(vars.Item("wgs" & i)) * kl_punkte(1) / 10
                    Punkte = Punkte + CSng(vars.Item("wtf" & i)) * kl_punkte(2) / 10 + CSng(vars.Item("win" & i)) * kl_punkte(4) / 10
                    Punkte = Punkte + CSng(vars.Item("win" & i)) * kl_punkte(5) / 10 + CSng(vars.Item("wdp" & i)) * kl_punkte(6) / 10
                End If
            End If
            Punkte = Punkte + add_verstoesse(vars, i)
            If Punkte < 0 Then Punkte = 0
            
        Case "LH_"
            Punkte = CSng(vars.Item("wsh" & i)) + CSng(vars.Item("wsd" & i)) + CSng(vars.Item("wtd" & i))
            Punkte = Punkte + CSng(vars.Item("wfg" & i)) - CSng(vars.Item("wfl" & i))
            If Punkte < 0 Then Punkte = 0
        
        Case Else
            Select Case DLookup("Land", "Startklasse", "Startklasse ='" & s_kl & "'")
                Case "BY"
                    kl_punkte = Punkteverteilung(s_kl, ch_runde(rde), rde)
                    If vars.exists("wgs" & i) Then
                        Punkte = CSng(vars.Item("wgs" & i)) * kl_punkte(0) / 10
                        Punkte = Punkte + CSng(vars.Item("wbd" & i)) * kl_punkte(1) / 10
                        Punkte = Punkte + CSng(vars.Item("wtf" & i)) * kl_punkte(2) / 10
                        Punkte = Punkte + CSng(vars.Item("win" & i)) * kl_punkte(3) / 10
                    Else
                        Punkte = 0
                    End If
                Case "BW"
                    If vars.exists("wth" & i) Then
                        Punkte = CSng(vars.Item("wth" & i)) + CSng(vars.Item("wtd" & i)) + CSng(vars.Item("wta" & i))
                        Punkte = Punkte + CSng(vars.Item("wak" & i)) - Replace(vars.Item("wfe" & i), ".", ",")
                    Else
                        Punkte = 0
                    End If
                Case "SL"
                    Punkte = CSng(vars.Item("wth" & i)) + CSng(vars.Item("wtd" & i)) + CSng(vars.Item("wta" & i))
                
                Case "D"
                    Punkte = CSng(vars.Item("wgs" & i))
'                    Punkte = CSng(vars.Item("Punkte" & i))
                    
                Case "NBS_"
                    Select Case left(s_kl, 6)
                        Case "BS_RR_"
                            Punkte = Punkte + CSng(vars.Item("wsh" & i)) + CSng(vars.Item("wsd" & i))
                            Punkte = Punkte + CSng(vars.Item("wth" & i)) - Replace(vars.Item("wfl" & i), ".", ",")
                     
                        Case "BS_F_R"
                            Punkte = Punkte + CSng(vars.Item("wsh" & i)) + CSng(vars.Item("wth" & i))
                            Punkte = Punkte + CSng(vars.Item("wch" & i)) - Val(Nz(Replace(vars.Item("wfl" & i), ".", ",")))
                              
                        Case "BS_"
                            Punkte = CSng(vars.Item("wsh" & i))
                              
                    End Select
                    If Punkte < 0 Then Punkte = 0
                
                Case Else
                    MsgBox "Fehler bei der Punkteberechnung" & vbCrLf & "Startklasse wurde nicht erkannt."

            End Select
            If Punkte < 0 Then Punkte = 0
    End Select
'    If Punkte < 0 Then Punkte = 0
    rechne_punkte = FormatNumber(Punkte, 2)
End Function

Private Function to_zahl(vars, wert)
    to_zahl = 0
    If vars.exists(wert) Then
        If vars.Item(wert) <> "" Then
            to_zahl = CSng(vars.Item(wert))
        End If
    End If
End Function

Public Function get_platzierung(rt)
    Dim re, pr As Recordset
    Dim rh As Integer
    Dim fName As String
    Dim t As Integer
    Dim cgivar
    Set db = CurrentDb
    ' Hier werden die Platzierungen eingelesen und in tbl Auswertung geschrieben
    fName = getBaseDir & get_TerNr & "_RT" & CLng(rt * 1000) & ".txt"
    If Len(Dir(fName)) > 0 Then
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set inp = fs.OpenTextFile(fName, 1, 0)  'reading
        Do Until inp.AtEndOfStream
            cgivar = Split(inp.Readline, ";")
            Set cgivar = zerlege(cgivar(2))
            For t = 1 To 8
                If cgivar.exists("wpt" & t) Then
                    Set pr = db.OpenRecordset("SELECT PR_ID FROM Paare_Rundenqualifikation WHERE TP_ID=" & cgivar.Item("PR_ID" & t) & " AND RT_ID=" & rt & ";")
                    Set re = db.OpenRecordset("SELECT * FROM Auswertung WHERE pr_id =" & pr!PR_ID & " AND wr_id =" & cgivar.Item("WR_ID") & ";")
                    re.Edit
                    re!Platz = cgivar.Item("wpt" & t)
                    re.Update
        
                End If
            Next
        Loop
    Else
        get_platzierung = 2
    End If
End Function

' neu newJudgingSystem
Function add_verstoesse(vars, i)
    Dim verst
    Dim x As Integer
    verst = Array("wsidebysidevw", "wakrovw", "whighlightvw", "wjuniorvw", "wkleidungvw", "wtanzbereichvw", "wtanzzeitvw", "waufrufvw")
    For x = 0 To UBound(verst)
        If vars.Item(verst(x) & i) <> "" Then
            add_verstoesse = add_verstoesse + CSng(vars.Item(verst(x) & i))
        End If
    Next
End Function

Function Punkteverteilung(Startklasse, rd, rde)
    Dim punkte_verteilung
    Select Case Startklasse
        Case "F_RR_ST"  ' Showteam
            punkte_verteilung = Array(15, 25, 20, 0, 7.5, 7.5, 25)
        Case "F_RR_GF"  ' Girl RR
            punkte_verteilung = Array(20, 20, 20, 0, 10, 10, 20)
        Case "F_RR_LF"  ' Lady RR
            punkte_verteilung = Array(20, 20, 20, 0, 10, 10, 20)
        Case "F_RR_J"   ' Jugend
            punkte_verteilung = Array(20, 20, 20, 0, 10, 10, 20)
        Case "F_RR_Q"   ' Quattro RR
            punkte_verteilung = Array(20, 20, 20, 0, 10, 10, 20)
        Case "F_RR_M"   ' Master RR
            punkte_verteilung = Array(20, 20, 20, 0, 10, 10, 20)
        Case "F_BW_M"   ' Master BW
            punkte_verteilung = Array(25, 25, 25, 0, 9, 8, 8)
        Case "RR_S"
            punkte_verteilung = Array(4.5, 4.5, 4.5, 4.5, 6.3, 6.3, 5.4)
        Case "RR_J"
            punkte_verteilung = Array(6, 6, 6, 6, 8.4, 8.4, 7.2)
        Case "RR_C"
            punkte_verteilung = Array(6, 6, 6, 6, 8.4, 8.4, 7.2)
        Case "RR_A", "RR_B"
            Select Case rd
                Case "VR", "ZR"
                    punkte_verteilung = Array(6.25, 6.25, 6.25, 6.25, 8.75, 8.75, 7.5)
                Case "ER"
                    punkte_verteilung = Array(4.375, 4.375, 4.375, 4.375, 6.125, 6.125, 5.25)
            End Select
            If rde = "Semi" Then punkte_verteilung = Array(8.75, 8.75, 8.75, 8.75, 12.25, 12.25, 10.5)
        ' Boogie NJS TSO1.8
        Case "BW_MA"
            punkte_verteilung = Array(15, 7.5, 10, 0, 15, 10, 7.5)
        Case "BW_SA"
            punkte_verteilung = Array(15, 7.5, 10, 0, 15, 10, 7.5)
        Case "BW_JA"
            punkte_verteilung = Array(15, 7.5, 10, 0, 15, 10, 7.5)
        Case "BW_MB"
            punkte_verteilung = Array(15, 7.5, 10, 0, 15, 10, 7.5)
        Case "BW_SB"
            punkte_verteilung = Array(15, 7.5, 10, 0, 15, 10, 7.5)
        Case "BW_NG":
            punkte_verteilung = Array(1.5, 1.5, 1, 1, 1, 1, 1, 1, 3, 3, 3)
        ' Breitensport Bayern
        Case "BS_BY_BJ", "BS_BY_BE", "BS_BY_BS", "BS_BY_S1":
            punkte_verteilung = Array(15, 10, 10, 30, 0, 0, 0)
        Case Else
            punkte_verteilung = Array(10, 10, 10, 10, 10, 10, 10, 10, 10)
    End Select
    Punkteverteilung = punkte_verteilung
End Function

'********* HM V14.03 ****************
' Faktor f�r Berechnung wegen fehlender T�nzer
' minumum und maximum Anzahl T�nzer f�r Startklasse
Function Faktor_Formation_Abzuege(Startklasse) As Formationswerte
    Select Case Startklasse
        Case "F_RR_ST"          ' Showteam
            Faktor_Formation_Abzuege.faktor = 0
            Faktor_Formation_Abzuege.min = 4
            Faktor_Formation_Abzuege.max = 16
        Case "F_RR_GF"          ' Girl RR
            Faktor_Formation_Abzuege.faktor = 1.75
            Faktor_Formation_Abzuege.min = 8
            Faktor_Formation_Abzuege.max = 12
        Case "F_RR_LF"          ' Lady RR
            Faktor_Formation_Abzuege.faktor = 1.25
            Faktor_Formation_Abzuege.min = 8
            Faktor_Formation_Abzuege.max = 16
        Case "F_RR_J"           ' Jugend
            Faktor_Formation_Abzuege.faktor = 1.25
            Faktor_Formation_Abzuege.min = 8
            Faktor_Formation_Abzuege.max = 12
        Case "F_RR_Q"           ' Quattro RR
            Faktor_Formation_Abzuege.faktor = 0
            Faktor_Formation_Abzuege.min = 8
            Faktor_Formation_Abzuege.max = 8
        Case "F_RR_M"           ' Master RR
            Faktor_Formation_Abzuege.faktor = 1.25
            Faktor_Formation_Abzuege.min = 8
            Faktor_Formation_Abzuege.max = 12
        Case "F_BW_M"           ' Master Boogie
            Faktor_Formation_Abzuege.faktor = 0
            Faktor_Formation_Abzuege.min = 8
            Faktor_Formation_Abzuege.max = 12
        Case Else
            Faktor_Formation_Abzuege.faktor = 0
            Faktor_Formation_Abzuege.min = 0
            Faktor_Formation_Abzuege.max = 0
    End Select
End Function

' Berechnung wegen fehlender T�nzer
Function Form_abzuege(PR_ID, s_kl)
    Dim db As Database
    Dim re As Recordset
    Dim f As Formationswerte
    Set db = CurrentDb
    Set re = db.OpenRecordset("SELECT Anz_Taenzer FROM Paare WHERE TP_ID= " & PR_ID & ";")
    f = Faktor_Formation_Abzuege(s_kl)
    
    Form_abzuege = (100 - ((f.max - re!Anz_Taenzer) * f.faktor)) / 100
End Function

Private Function get_ft(WR_ID, PR_ID, RT_ID) ' bei A/B Vorrunde FT miteinrechnen
    Dim re As Recordset
    Set db = CurrentDb()
    
    Set re = db.OpenRecordset("SELECT Punkte FROM Paare_Rundenqualifikation INNER JOIN Auswertung ON Paare_Rundenqualifikation.PR_ID = Auswertung.PR_ID WHERE (Paare_Rundenqualifikation.TP_ID=" & PR_ID & " AND WR_ID=" & WR_ID & " AND RT_ID=" & RT_ID & ");", DB_OPEN_DYNASET)
    If re.EOF Then
        '   MsgBox "Fehler! Es existiert keine Fu�technik-Wertung"
    Else
        get_ft = re!Punkte
    End If
    re.Close
End Function

Function add_akro(scr_elem, i, t)
    Dim Punkte As Single
    Dim z As Integer
    For z = 1 To 8
        If scr_elem.Item("wak" & i & z) <> "" Then
            Punkte = Punkte + CSng(scr_elem.Item("wak" & i & z))
            t = t + 1
        End If
        If scr_elem.Item("wflak" & i & z) <> "" Then
            Punkte = Punkte - CSng(scr_elem.Item("wflak" & i & z))
        End If
        If scr_elem.Item("wfl" & i & "_ak" & i & z) <> "" Then
            Punkte = Punkte - CSng(scr_elem.Item("wfl" & i & "_ak" & i & z))
        End If
    Next
    add_akro = Punkte
End Function

Function zerlege(inp)
    Dim vars, var, back
    Dim i As Integer
    Set back = CreateObject("Scripting.Dictionary")
    vars = Split(inp, "&")
    For i = 0 To UBound(vars)
        If InStr(1, vars(i), "wertung_in") = 0 Then
            var = Split(vars(i), "=")
            back.Add var(0), Replace(var(1), ".", ",")
        End If
    Next
    Set zerlege = back
End Function

Public Function eins_zwei(p_id, vars)
    If CSng(vars.Item("PR_ID1")) = p_id Then
        eins_zwei = 1
    ElseIf vars.exists("PR_ID2") Then
        eins_zwei = 2
    End If

End Function

'*****AB***** V13.02 neue Funktion zum Einlesen der RT_Daten
Public Function Str_to_Sng(AuswerteString As String)

'*** Wandelt einen �bergebenen String in einen CSingle Wert um, wenn der String nicht leer ist, sonst gibt er Null zur�ck

If IsEmpty(AuswerteString) Or AuswerteString = "" Then
    Str_to_Sng = Null
Else
    Str_to_Sng = CSng(AuswerteString)
End If
End Function

'*****AB***** V13.02 neue Funktion zum Einlesen der RT_Daten
Public Sub Import_RT_txt(RundenTab_ID)

On Error GoTo RT_Import_Fehler_Err
'*** ben�tigte �bergabewerte Runden_ID aus RTTabelle
' Parameter RundenTab_ID As Integer

Dim Werte_Array, Werte_Array_Zwischenergebnis, Werte_Assoz_Array
Dim SQL_String, SQL_Insert_Werte, SQL_Insert_Felder, inputSTR, fName As String
Dim n, Akroz�hler As Integer
Dim fs, inp, cgivar, Zeile, Testarray
Dim anzahl_paare As Integer
Dim AbgegebeneWertungen, rt, html_felder As Recordset
Dim db As Database

Set db = CurrentDb()
Set AbgegebeneWertungen = db.OpenRecordset("SELECT * from Abgegebene_Wertungen;", DB_OPEN_DYNASET)
Set Werte_Assoz_Array = CreateObject("Scripting.Dictionary")
Set rt = db.OpenRecordset("Select * from rundentab where rt_id = " & RundenTab_ID & ";", DB_OPEN_DYNASET)
Set html_felder = db.OpenRecordset("Select * from Wertungsb�gen where wb =""" & left(rt!Startklasse, 3) & """;", DB_OPEN_DYNASET)

    fName = getBaseDir & get_TerNr & "_RT" & CInt(RundenTab_ID) & ".txt"   ' Dateiname erstellen
    
    If Len(Dir(fName)) > 0 Then
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set inp = fs.OpenTextFile(fName, 1, 0)  'reading
        Do Until inp.AtEndOfStream
            cgivar = Split(inp.Readline, ";")
            Zeile = cgivar(2)
    
            If left(rt!Startklasse, 3) = "RR_" Then
                Zeile = Replace(Zeile, "_1=", "1=")
                Zeile = Replace(Zeile, "_2=", "2=")
            End If
            
            Werte_Array = Split(Zeile, "&")
            
            '*** pr�fen ob ein oder zwei Paare im String stehen
                Dim back, var
                Dim ki As Integer
                Set back = CreateObject("Scripting.Dictionary")
                For ki = 0 To UBound(Werte_Array)
                  var = Split(Werte_Array(ki), "=")
                  back.Add var(0), Replace(var(1), ".", ",")
                Next

            If back.exists("PR_ID2") Then
                If CSng(back.Item("PR_ID1")) > 0 Then
                    anzahl_paare = 2
                Else
                    anzahl_paare = 1
                End If
            Else
                anzahl_paare = 1
            End If
            
            Dim Paar_ID As Single
            Dim textakro As String
            
            'Werte aus dem Scripting.Dictionary raus holen, bei zwei Paaren nacheinander
            For n = 1 To anzahl_paare
                
                AbgegebeneWertungen.FindFirst ("Paar_ID = " & CSng(back.Item("PR_ID" & n)) & " AND RundenTab_ID = " & CSng(back.Item("rt_ID")) & "AND Wertungsrichter_ID = " & CSng(back.Item("WR_ID")))
                
                If AbgegebeneWertungen.NoMatch Then
                    AbgegebeneWertungen.AddNew
                Else
                    AbgegebeneWertungen.Edit
                End If
                AbgegebeneWertungen!Paar_ID = CSng(back.Item("PR_ID" & n))
                AbgegebeneWertungen!rh = CSng(back.Item("rh" & n))
                AbgegebeneWertungen!Wertungsrichter_ID = CSng(back.Item("WR_ID"))
                AbgegebeneWertungen!RundenTab_ID = CSng(back.Item("rt_ID"))
                 
                AbgegebeneWertungen!Herr_Grundtechnik = Str_to_Sng(back.Item(html_felder!Ber1 & n))
                AbgegebeneWertungen!Herr_Haltung_Drehtechnik = Str_to_Sng(back.Item(html_felder!Ber2 & n))
                AbgegebeneWertungen!Dame_Grundtechnik = Str_to_Sng(back.Item(html_felder!Ber3 & n))
                AbgegebeneWertungen!Dame_Haltung_Drehtechnik = Str_to_Sng(back.Item(html_felder!Ber4 & n))
                AbgegebeneWertungen!Choreographie = Str_to_Sng(back.Item(html_felder!Ber5 & n))
                AbgegebeneWertungen!Tanzfiguren = Str_to_Sng(back.Item(html_felder!Ber6 & n))
                AbgegebeneWertungen!T�nzerische_Darbietung = Str_to_Sng(back.Item(html_felder!Ber10 & n))
                AbgegebeneWertungen!Grobfehler_Text = back.Item(html_felder!Ber8 & n)
                AbgegebeneWertungen!Grobfehler_Summe = CSng(back.Item(html_felder!Ber9 & n))
                
                For Akroz�hler = 1 To 8
                    AbgegebeneWertungen("Akrobatik" & Akroz�hler) = Str_to_Sng(back.Item("wak" & n & Akroz�hler))
                    If back.exists("tflak" & n & Akroz�hler) Then
                        AbgegebeneWertungen("Akrobatik" & Akroz�hler & "_Grobfehler_Text") = back.Item("tfl" & "ak" & n & Akroz�hler)
                        AbgegebeneWertungen("Akrobatik" & Akroz�hler & "_Grobfehler_Summe") = Str_to_Sng(back.Item("wfl" & "ak" & n & Akroz�hler))
                    Else
                        AbgegebeneWertungen("Akrobatik" & Akroz�hler & "_Grobfehler_Text") = back.Item("tfl" & n & "_ak" & n & Akroz�hler)
                        AbgegebeneWertungen("Akrobatik" & Akroz�hler & "_Grobfehler_Summe") = Str_to_Sng(back.Item("wfl" & n & "_ak" & n & Akroz�hler))
                    End If
                Next
                AbgegebeneWertungen.Update
            Next
        Loop
    Else
        MsgBox "kein Import"
    End If
    
RT_Import_Fehler_Exit:
        'Funktion verlassen
        Exit Sub
        
RT_Import_Fehler_Err:
    Resume Next

End Sub

Function Observer_FT(st_kl, abge_Wertung, arr, rd)
    Dim st_klasse
    st_klasse = Punkteverteilung(st_kl, ch_runde(rd), rd)
    If IsNull(abge_Wertung) Or abge_Wertung = "" Then
        Observer_FT = "&nbsp;"
    Else
        If left(st_kl, 3) = "BW_" Then
            Observer_FT = abge_Wertung * st_klasse(arr) / 10
        ElseIf left(st_kl, 3) = "BS_" Then
            Observer_FT = abge_Wertung
        ElseIf left(st_kl, 3) = "F_B" Then
            Observer_FT = abge_Wertung
        Else
            Observer_FT = (10 - abge_Wertung) * 10
        End If
    End If
End Function

Public Sub ObserverHTML(trunde)
'*****AB***** V13.04 -  neue Funktion zum Anzeigen der Wertungen f�r den Observer unter IP_Adresse-Webserver/observer.html

    On Error GoTo ObserverHTML_Fehler_Err
    
    Dim HTML_Website, HTML_Paar_links, HTML_Paar_rechts As Variant
    Dim HTML_�berschrift As String
    Dim HTML_WR_Template As String
    Dim HTML_WR_Werte, sql, test, PaarLinks, PaarRechts As String
    Dim st_kl As String
    Dim AbgegebeneWertungen, Paar_Infos As Recordset
    Dim WR_Zaehler, x, A_WR, T_WR, Paar_ID, Letzte_Runde, letzte_Tanzrunde As Integer
    Dim i, t As Integer
    Dim T_WR_Reset, A_WR_Reset As Boolean
    Dim kl_punkte
    Dim db As Database
    Dim rd As String
    Dim GesamtPunkte As Double
    
    '*** Checken ob Website neu aufgebaut oder nur aktualisiert werden muss
    sql = "SELECT Abgegebene_Wertungen.RundenTab_ID, Abgegebene_Wertungen.rh, Abgegebene_Wertungen.Wertungsrichter_ID FROM Abgegebene_Wertungen;"
    
    Set db = CurrentDb()
    Set AbgegebeneWertungen = db.OpenRecordset(sql, DB_OPEN_DYNASET)
    rd = ch_runde(trunde)
    AbgegebeneWertungen.MoveLast
    
    Letzte_Runde = AbgegebeneWertungen!RundenTab_ID
    letzte_Tanzrunde = AbgegebeneWertungen!rh
    
    Set AbgegebeneWertungen = db.OpenRecordset("SELECT Wert_Richter.WR_Kuerzel, Wert_Richter.WR_Nachname, Startklasse_Wertungsrichter.WR_function, AW.* FROM (Startklasse_Wertungsrichter INNER JOIN (Rundentab INNER JOIN Abgegebene_Wertungen AS AW ON Rundentab.RT_ID = AW.RundenTab_ID) ON Startklasse_Wertungsrichter.Startklasse = Rundentab.Startklasse) INNER JOIN Wert_Richter ON (Startklasse_Wertungsrichter.WR_ID = Wert_Richter.WR_ID) AND (AW.Wertungsrichter_ID = Wert_Richter.WR_ID) WHERE (((AW.rh)=" & letzte_Tanzrunde & ") AND ((AW.RundenTab_ID)=" & Letzte_Runde & ")) ORDER BY AW.Paar_ID, Wert_Richter.WR_Kuerzel;", DB_OPEN_DYNASET)
    Set Paar_Infos = db.OpenRecordset("SELECT RT.RT_ID, PRQ.Rundennummer, PRQ.TP_ID, RT.Startklasse, Startklasse.Startklasse_text, RT.Runde, Paare.Startnr, Paare.Startnr, IIf([isTeam],[Name_Team],[Da_Nachname]) AS Ausdr1, Paare.He_Nachname FROM (Paare INNER JOIN (Rundentab AS RT INNER JOIN Paare_Rundenqualifikation AS PRQ ON RT.RT_ID = PRQ.RT_ID) ON Paare.TP_ID = PRQ.TP_ID) INNER JOIN Startklasse ON RT.Startklasse = Startklasse.Startklasse WHERE (((RT.RT_ID)=" & Letzte_Runde & ") AND ((PRQ.Rundennummer)=" & letzte_Tanzrunde & ")) ORDER BY RT.RT_ID, PRQ.Rundennummer, Paare.Startnr;", DB_OPEN_DYNASET)
    
    HTML_Website = ""
    If left(Paar_Infos!Startklasse, 3) = "BW_" Then ' Or Left(Paar_Infos!Startklasse, 3) = "F_B" Then
        HTML_�berschrift = "<table border='1' cellpadding='1' cellspacing='1' style='width: 1024px; text-align: center;'><tbody><tr bgcolor=#d0d0d0><td>Name</td><td>Grundschritt</td><td>Basic Dancing</td><td>Tanzfig</td><td>Interpret</td><td>Spontane Int</td><td>Dance Perf</td><td>Summe</td></tr><tr><td>TWR1</td></tr><tr><td>TWR2</td></tr><tr><td>TWR3</td></tr><tr><td>TWR4</td></tr><tr><td>TWR5</td></tr><tr><td>TWR6</td></tr><tr><td>TWR7</td></tr><tr></tr>"
        HTML_�berschrift = HTML_�berschrift & "</tbody></table>"
        HTML_WR_Template = "<td>WRNAME</td><td>WERT01</td><td>WERT02</td><td>WERT03</td><td>WERT05</td><td>WERT06</td><td>WERT07</td><td>WERT08</td>"
    Else
        HTML_�berschrift = "<table border='1' cellpadding='1' cellspacing='1' style='width: 1024px;'><tbody><tr bgcolor=#d0d0d0><td>Name</td><td>GT&nbsp;H</td><td>HD&nbsp;H</td><td>GT&nbsp;D</td><td>HD&nbsp;D</td><td>Chor</td><td>Tanzf.</td><td>T&auml;nzD</td><td>Summe</td><td>&nbsp;</td><td>Grobf</td><td>Abz�ge</td><td>&nbsp;</td><td>Punkte</td></tr><tr><td>TWR1</td></tr><tr><td>TWR2</td></tr><tr><td>TWR3</td></tr><tr><td>TWR4</td></tr><tr><td>TWR5</td></tr><tr></tr>"
        HTML_�berschrift = HTML_�berschrift & "<tr bgcolor=#d0d0d0><td>Name</td><td>Akro1</td><td>GF&nbsp;1</td><td>Akro2</td><td>GF&nbsp;2</td><td>Akro3</td><td>GF&nbsp;3</td><td>Akro4</td><td>GF&nbsp;4</td><td>Akro5</td><td>GF&nbsp;5</td><td>Akro6</td><td>GF&nbsp;6</td><td>Akro7</td><td>GF&nbsp;7</td><td>Akro8</td><td>GF&nbsp;8</td><td>Punkte</td></tr><tr><td>AWR1</td></tr><tr><td>AWR2</td></tr><tr><td>AWR3</td></tr><tr><td>AWR4</td></tr><tr></tr></tbody></table>"
        HTML_WR_Template = "<td>WRNAME</td><td>WERT01</td><td>WERT02</td><td>WERT03</td><td>WERT04</td><td>WERT05</td><td>WERT06</td><td>WERT07</td><td>WERT08</td><td>WERT09</td><td>WERT10</td><td>WERT11</td><td>WERT12</td><td>WERT13</td><td>WERT14</td><td>WERT15</td><td>WERT16</td><td>WERT17</td>"
    End If
    
    '**** neue Seite erzeugen
    HTML_Paar_links = HTML_�berschrift
    HTML_Paar_rechts = HTML_�berschrift
    
    AbgegebeneWertungen.MoveLast
    WR_Zaehler = AbgegebeneWertungen.RecordCount
    AbgegebeneWertungen.MoveFirst
    
    A_WR = 1
    T_WR = 1
    T_WR_Reset = False
    A_WR_Reset = False
    Paar_ID = AbgegebeneWertungen!Paar_ID
    st_kl = Paar_Infos!Startklasse
    
    For x = 1 To WR_Zaehler
        HTML_WR_Werte = HTML_WR_Template
    
        Paar_Infos.FindFirst "TP_ID = " & AbgegebeneWertungen!Paar_ID
    
        If AbgegebeneWertungen!WR_function = "Ft" Or AbgegebeneWertungen!WR_function = "X" Then
        
            GesamtPunkte = 0
            Select Case left(st_kl, 3)
                Case "BW_"
                    kl_punkte = Punkteverteilung(st_kl, rd, trunde)
                    GesamtPunkte = CSng(AbgegebeneWertungen!Herr_Grundtechnik) * kl_punkte(0) / 10
                    GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Herr_Haltung_Drehtechnik) * kl_punkte(1) / 10
                    GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Dame_Grundtechnik) * kl_punkte(2) / 10
                    GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Choreographie) * kl_punkte(4) / 10
                    GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Tanzfiguren) * kl_punkte(5) / 10
                    GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!T�nzerische_Darbietung) * kl_punkte(6) / 10
                Case "BS_"
                    GesamtPunkte = CSng(AbgegebeneWertungen!Herr_Grundtechnik)
                Case "F_B"
                    kl_punkte = Punkteverteilung(st_kl, rd, trunde)
                    GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Herr_Grundtechnik) * kl_punkte(0) / 10
                    GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Herr_Haltung_Drehtechnik) * kl_punkte(1) / 10
                    GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Dame_Grundtechnik) * kl_punkte(2) / 10
                    GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Choreographie) * kl_punkte(4) / 10
                    GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Tanzfiguren) * kl_punkte(5) / 10
                    GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!T�nzerische_Darbietung) * kl_punkte(6) / 10
                Case "F_R"
                    kl_punkte = Punkteverteilung(st_kl, rd, trunde)
                    GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Herr_Grundtechnik) * kl_punkte(0) / 10
                    GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Herr_Haltung_Drehtechnik) * kl_punkte(1) / 10
                    GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Dame_Grundtechnik) * kl_punkte(2) / 10
                    GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Choreographie) * kl_punkte(4) / 10
                    GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Tanzfiguren) * kl_punkte(5) / 10
                    GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!T�nzerische_Darbietung) * kl_punkte(6) / 10
                Case Else
                    GesamtPunkte = (CSng(AbgegebeneWertungen!Herr_Grundtechnik) / 2) + (CSng(Nz(AbgegebeneWertungen!Herr_Haltung_Drehtechnik)) / 2) + (CSng(AbgegebeneWertungen!Dame_Grundtechnik) / 2)
                    GesamtPunkte = GesamtPunkte + (CSng(Nz(AbgegebeneWertungen!Dame_Haltung_Drehtechnik)) / 2) + (CSng(AbgegebeneWertungen!Choreographie) * 6 / 10) + (CSng(Nz(AbgegebeneWertungen!Tanzfiguren) * 6 / 10)) + (CSng(AbgegebeneWertungen!T�nzerische_Darbietung) * 8 / 10)
            End Select
            If GesamtPunkte < 0 Then GesamtPunkte = 0
            
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WRNAME", AbgegebeneWertungen!WR_Nachname)
            'einen FT-WR einf�gen
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT01", Observer_FT(st_kl, AbgegebeneWertungen!Herr_Grundtechnik, 0, trunde))
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT02", Observer_FT(st_kl, AbgegebeneWertungen!Herr_Haltung_Drehtechnik, 1, trunde))
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT03", Observer_FT(st_kl, AbgegebeneWertungen!Dame_Grundtechnik, 2, trunde))
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT04", Observer_FT(st_kl, AbgegebeneWertungen!Dame_Haltung_Drehtechnik, 3, trunde))
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT05", Observer_FT(st_kl, AbgegebeneWertungen!Choreographie, 4, trunde))
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT06", Observer_FT(st_kl, AbgegebeneWertungen!Tanzfiguren, 5, trunde))
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT07", Observer_FT(st_kl, AbgegebeneWertungen!T�nzerische_Darbietung, 6, trunde))
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT08", Round(GesamtPunkte, 2))
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT09", "&nbsp;")
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT10", IIf(IsNull(AbgegebeneWertungen!Grobfehler_Text), "&nbsp;", AbgegebeneWertungen!Grobfehler_Text))
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT11", AbgegebeneWertungen!Grobfehler_Summe)
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT12", "&nbsp;")
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT13", Round(GesamtPunkte - AbgegebeneWertungen!Grobfehler_Summe, 2))
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT14", "&nbsp;")
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT15", "&nbsp;")
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT16", "&nbsp;")
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT17", "&nbsp;")
            
            If AbgegebeneWertungen!Paar_ID <> Paar_ID And T_WR_Reset = False Then
                T_WR = 1
                T_WR_Reset = True
            End If
            If AbgegebeneWertungen!Paar_ID = Paar_ID Then
                HTML_Paar_links = Replace(HTML_Paar_links, "<td>TWR" & T_WR & "</td>", HTML_WR_Werte)
                PaarLinks = Paar_Infos!Startnr & " " & Paar_Infos!Ausdr1 & " / " & Paar_Infos!He_Nachname
            Else
                HTML_Paar_rechts = Replace(HTML_Paar_rechts, "<td>TWR" & T_WR & "</td>", HTML_WR_Werte)
                PaarRechts = Paar_Infos!Startnr & " " & Paar_Infos!Ausdr1 & " / " & Paar_Infos!He_Nachname
            End If
           T_WR = T_WR + 1
        Else
            
            GesamtPunkte = 0
            For i = 1 To 8
                If Not IsNull(AbgegebeneWertungen!Akrobatik1) Then
                    GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen("Akrobatik" & i)) - CSng(AbgegebeneWertungen("Akrobatik" & i & "_Grobfehler_Summe"))
                    t = t + 1
                End If
            Next
            If st_kl = "F_RR_M" Then ' Master RR
                GesamtPunkte = GesamtPunkte / IIf(t < 7, 6, t) * 5
            Else
                GesamtPunkte = 0
            End If
            If GesamtPunkte < 0 Then GesamtPunkte = 0
            
            'einen AK-WR einf�gen
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WRNAME", AbgegebeneWertungen!WR_Nachname)
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT01", Round((100 - (AbgegebeneWertungen!Akrobatik1 * 100 / Get_Akropunkte(AbgegebeneWertungen!Paar_ID, Paar_Infos!Runde, 1))), 0) & " ")
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT02", AbgegebeneWertungen!Akrobatik1_Grobfehler_Text & " ")
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT03", Round((100 - (AbgegebeneWertungen!Akrobatik2 * 100 / Get_Akropunkte(AbgegebeneWertungen!Paar_ID, Paar_Infos!Runde, 2))), 0) & " ")
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT04", AbgegebeneWertungen!Akrobatik2_Grobfehler_Text & " ")
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT05", Round((100 - (AbgegebeneWertungen!Akrobatik3 * 100 / Get_Akropunkte(AbgegebeneWertungen!Paar_ID, Paar_Infos!Runde, 3))), 0) & " ")
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT06", AbgegebeneWertungen!Akrobatik3_Grobfehler_Text & " ")
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT07", Round((100 - (AbgegebeneWertungen!Akrobatik4 * 100 / Get_Akropunkte(AbgegebeneWertungen!Paar_ID, Paar_Infos!Runde, 4))), 0) & " ")
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT08", AbgegebeneWertungen!Akrobatik4_Grobfehler_Text & " ")
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT09", Round((100 - (AbgegebeneWertungen!Akrobatik5 * 100 / Get_Akropunkte(AbgegebeneWertungen!Paar_ID, Paar_Infos!Runde, 5))), 0) & " ")
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT10", AbgegebeneWertungen!Akrobatik5_Grobfehler_Text & " ")
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT11", Round((100 - (AbgegebeneWertungen!Akrobatik6 * 100 / Get_Akropunkte(AbgegebeneWertungen!Paar_ID, Paar_Infos!Runde, 6))), 0) & " ")
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT12", AbgegebeneWertungen!Akrobatik6_Grobfehler_Text & " ")
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT13", Round((100 - (AbgegebeneWertungen!Akrobatik7 * 100 / Get_Akropunkte(AbgegebeneWertungen!Paar_ID, Paar_Infos!Runde, 7))), 0) & " ")
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT14", AbgegebeneWertungen!Akrobatik7_Grobfehler_Text & " ")
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT15", Round((100 - (AbgegebeneWertungen!Akrobatik8 * 100 / Get_Akropunkte(AbgegebeneWertungen!Paar_ID, Paar_Infos!Runde, 8))), 0) & " ")
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT16", AbgegebeneWertungen!Akrobatik8_Grobfehler_Text & " ")
            HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT17", Round(GesamtPunkte, 2))
           
            If AbgegebeneWertungen!Paar_ID <> Paar_ID And A_WR_Reset = False Then
                A_WR = 1
                A_WR_Reset = True
            End If
            If AbgegebeneWertungen!Paar_ID = Paar_ID Then
                HTML_Paar_links = Replace(HTML_Paar_links, "<td>AWR" & A_WR & "</td>", HTML_WR_Werte)
            Else
                HTML_Paar_rechts = Replace(HTML_Paar_rechts, "<td>AWR" & A_WR & "</td>", HTML_WR_Werte)
            End If
            A_WR = A_WR + 1
        End If
    
        
        AbgegebeneWertungen.MoveNext
    
    Next x
        
    HTML_Website = "<H1>" & Paar_Infos!Startklasse_text & " " & Paar_Infos!Runde & "</H1><H2>" & PaarLinks & "</H2>" & HTML_Paar_links & "<br><br><H2>" & PaarRechts & "</H2>" & HTML_Paar_rechts & "<br><br>"
    
    Dim out
    Dim pfad As String
    Dim Server_IP
    
    'Server_IP = GetIpAddrTable()
    
    pfad = getBaseDir & "Apache2\htdocs\observer\index.html"
    Set out = file_handle(pfad)
    out.writeline ("<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3.org/TR/html4/loose.dtd""><html><head><title>" & Forms![A-Programm�bersicht]!Turnierbez & "Observer �bersicht" & "</title> <meta http-equiv=""refresh"" content=""5""; URL="" " & GetIpAddrTable() & "/observer.html""><meta http-equiv=""expires"" content=""0""></head><body>" & HTML_Website & "</body></html>")
    out.Close

ObserverHTML_Fehler_Exit:
        'Funktion verlassen
        Exit Sub
        
ObserverHTML_Fehler_Err:
    Resume Next



End Sub

Public Function Get_Akropunkte(TP_ID, Runde, Akronummer)
    Dim db As Database
    Dim Paare As Recordset
    Dim RundTxt, AkroText As String
    Set db = CurrentDb()
    
    Set Paare = db.OpenRecordset("select * from Paare where TP_ID = " & TP_ID, DB_OPEN_DYNASET)
    'Set paare = db.OpenRecordset("SELECT Paare.*, Paare.TP_ID FROM Paare WHERE (((Paare.TP_ID)=2));", DB_OPEN_DYNASET)
    
    If Runde = "Vor_r" Then
        RundTxt = "_VR"
    ElseIf Runde = "1_Zw_r" Or Runde = "2_Zw_r" Or Runde = "3_Zw_r" Then
        RundTxt = "_ZR"
    ElseIf Runde = "End_r" Or Runde = "End_r_Akro" Then
        RundTxt = "_ER"
  
    End If
    
    AkroText = "Wert" & Akronummer & RundTxt
    Get_Akropunkte = Paare(AkroText)
    
End Function
