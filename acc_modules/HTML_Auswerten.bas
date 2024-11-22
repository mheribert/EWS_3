Option Compare Database
Option Explicit

    Public Type Formationswerte
      faktor    As Single           ' Faktor für Reduzierung bei Berechnung
      min       As Integer          ' minimum Anzahl Tänzer
      max       As Integer          ' maximum Anzahl Tänzer
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
        Set re = db.OpenRecordset("SELECT * from RundenTab WHERE Startklasse = '" & s_kl & "' AND Runde = '" & left(rde, 3) & "_r_Fuß';", DB_OPEN_DYNASET)
        ft_rt = re!RT_ID
        get_wertungen re!RT_ID, s_kl, re!Runde      'Fußtechnik rekursiv einlesen
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
                    re!Punkte = rechne_punkte(cgivar(0), cgivar(2), s_kl, rh, rde, ft_rt, WR_func)   ' Punkte Klassen- und Rundenabhängig ausrechen
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
    
    ' Wertungen löschen, die nicht rein gehören z.B. unentschuldigt
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
    Dim i, rh, X As Integer
    Dim Punkte As Double
    Dim verst
    verst = Array("beobachter_zukurz", "beobachter_zulang", "beobachter_Makeup", "beobachter_schmuck", "beobachter_requsit")

    Set vars = zerlege(inp)
    i = eins_zwei(PR_ID, vars)
    rh = vars.Item("rh" & i)
    
    For X = 0 To UBound(verst)
        If vars.Item(verst(X)) <> "" Then
            Punkte = Punkte + CSng(vars.Item(verst(X)))
        End If
    
    Next
    rechne_abzuege = Punkte
End Function

Private Function rechne_punkte(PR_ID, inp, s_kl, rh, rde, ft_rt, WR_func)
    'Punkte Klassen- und Rundenabhängig ausrechen
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
            If vars.exists("wng_ttd" & i) Then   ' newGuidelines         Kategorien-Streichverfahren
                Punkte = CSng(vars.Item("Punkte_err" & i))
            Else        ' alte Bewertung
                If vars.Item("wtk1") <> "" Then
                    kl_punkte = Punkteverteilung(s_kl, ch_runde(rde), rde)
                    Punkte = Punkte + CSng(vars.Item("wtk1")) * kl_punkte(0) / 10 + CSng(vars.Item("wch1")) * kl_punkte(1) / 10
                    Punkte = Punkte + CSng(vars.Item("wtf1")) * kl_punkte(2) / 10 + CSng(vars.Item("wab1")) * kl_punkte(4) / 10
                    Punkte = Punkte + CSng(vars.Item("waw1")) * kl_punkte(5) / 10 + CSng(vars.Item("waf1")) * kl_punkte(6) / 10
                End If
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
                    If InStr(1, rde, "_Fuß") > 0 Then Punkte = 0
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
                    If vars.exists("wsh" & i) Or vars.exists("wverw" & i) Then
                        Punkte = CSng(vars.Item("wsh" & i)) * kl_punkte(0) / 10
                        Punkte = Punkte + CSng(vars.Item("wsd" & i)) * kl_punkte(1) / 10
                        Punkte = Punkte + CSng(vars.Item("wbd" & i)) * kl_punkte(2) / 10
                        Punkte = Punkte + CSng(vars.Item("wtf" & i)) * kl_punkte(3) / 10
                        Punkte = Punkte + CSng(vars.Item("win" & i)) * kl_punkte(4) / 10
                        Punkte = Punkte + add_verstoesse(vars, i)
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
    Dim X As Integer
    verst = Array("wsidebysidevw", "wakrovw", "whighlightvw", "wjuniorvw", "wkleidungvw", "wtanzbereichvw", "wtanzzeitvw", "waufrufvw", "wverw")
    For X = 0 To UBound(verst)
        If vars.Item(verst(X) & i) <> "" Then
            add_verstoesse = add_verstoesse + CSng(vars.Item(verst(X) & i))
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
        Case "BW_NG"
            punkte_verteilung = Array(1.5, 1.5, 1.5, 1.5, 1.5, 1, 1, 1, 2.5, 2.5, 2.5)
        ' Breitensport Bayern
        Case "BS_BY_BJ", "BS_BY_BE", "BS_BY_BS", "BS_BY_S1", "BS_BY_FU", "BS_BY_SH"
            punkte_verteilung = Array(7.5, 7.5, 15, 10, 25, 0, 0, 0)
        Case Else
            punkte_verteilung = Array(10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10)
    End Select
    Punkteverteilung = punkte_verteilung
End Function

'********* HM V14.03 ****************
' Faktor für Berechnung wegen fehlender Tänzer
' minumum und maximum Anzahl Tänzer für Startklasse
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

' Berechnung wegen fehlender Tänzer
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
        '   MsgBox "Fehler! Es existiert keine Fußtechnik-Wertung"
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

Public Function zerlege(inp)
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

'*** Wandelt einen übergebenen String in einen CSingle Wert um, wenn der String nicht leer ist, sonst gibt er Null zurück

If IsEmpty(AuswerteString) Or AuswerteString = "" Then
    Str_to_Sng = Null
Else
    Str_to_Sng = CSng(AuswerteString)
End If
End Function

'*****AB***** V13.02 neue Funktion zum Einlesen der RT_Daten
Public Sub Import_RT_txt(RundenTab_ID)
'*** benötigte Übergabewerte Runden_ID aus RTTabelle
' Parameter RundenTab_ID As Integer
On Error GoTo RT_Import_Fehler_Err

    If get_properties("EWS") <> "EWS1" Then Exit Sub


    Dim Werte_Array, Werte_Array_Zwischenergebnis, Werte_Assoz_Array
    Dim SQL_String, SQL_Insert_Werte, SQL_Insert_Felder, inputSTR, fName As String
    Dim n, Akrozähler As Integer
    Dim fs, inp, cgivar, Zeile, Testarray
    Dim anzahl_paare As Integer
    Dim AbgegebeneWertungen, rt, html_felder As Recordset
    Dim db As Database
    
    Set db = CurrentDb()
    Set AbgegebeneWertungen = db.OpenRecordset("SELECT * from Abgegebene_Wertungen;", DB_OPEN_DYNASET)
    Set Werte_Assoz_Array = CreateObject("Scripting.Dictionary")
    Set rt = db.OpenRecordset("Select * from rundentab where rt_id = " & RundenTab_ID & ";", DB_OPEN_DYNASET)
    If get_properties("EWS") = "EWS1" And left(rt!Startklasse, 3) = "BW_" Then
        Set html_felder = db.OpenRecordset("Select * from Wertungsbögen where wb ='DBW_alt';", DB_OPEN_DYNASET)
    Else
        Set html_felder = db.OpenRecordset("Select * from Wertungsbögen where wb ='D" & left(rt!Startklasse, 3) & "';", DB_OPEN_DYNASET)
    End If
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
            
            '*** prüfen ob ein oder zwei Paare im String stehen
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
                AbgegebeneWertungen!Tänzerische_Darbietung = Str_to_Sng(back.Item(html_felder!Ber10 & n))
                AbgegebeneWertungen!Grobfehler_Text = back.Item(html_felder!Ber8 & n)
                AbgegebeneWertungen!Grobfehler_Summe = CSng(back.Item(html_felder!Ber9 & n))
                
                For Akrozähler = 1 To 8
                    AbgegebeneWertungen("Akrobatik" & Akrozähler) = Str_to_Sng(back.Item("wak" & n & Akrozähler))
                    If back.exists("tflak" & n & Akrozähler) Then
                        AbgegebeneWertungen("Akrobatik" & Akrozähler & "_Grobfehler_Text") = back.Item("tfl" & "ak" & n & Akrozähler)
                        AbgegebeneWertungen("Akrobatik" & Akrozähler & "_Grobfehler_Summe") = Str_to_Sng(back.Item("wfl" & "ak" & n & Akrozähler))
                    Else
                        AbgegebeneWertungen("Akrobatik" & Akrozähler & "_Grobfehler_Text") = back.Item("tfl" & n & "_ak" & n & Akrozähler)
                        AbgegebeneWertungen("Akrobatik" & Akrozähler & "_Grobfehler_Summe") = Str_to_Sng(back.Item("wfl" & n & "_ak" & n & Akrozähler))
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
'*****AB***** V13.04 -  neue Funktion zum Anzeigen der Wertungen für den Observer unter IP_Adresse-Webserver/observer.html

    On Error GoTo ObserverHTML_Fehler_Err
    
    Dim HTML_Website, HTML_Paar_links, HTML_Paar_rechts As Variant
    Dim HTML_überschrift As String
    Dim HTML_WR_Template As String
    Dim HTML_WR_Werte, sql, test, PaarLinks, PaarRechts As String
    Dim st_kl As String
    Dim vars
    Dim AbgegebeneWertungen, Paar_Infos As Recordset
    Dim WR_Zaehler, X, Paar_ID, Letzte_Runde, letzte_Tanzrunde As Integer
    Dim Anz_Paare As Integer
    Dim A_WR(2), T_WR(2) As Integer
    Dim i, t As Integer
    Dim seite, seiten As Integer
    Dim T_WR_Reset, A_WR_Reset As Boolean
    Dim kl_punkte
    Dim db As Database
    Dim rd As String
    Dim GesamtPunkte As Double
    
    Set db = CurrentDb()
    sql = "SELECT * FROM Auswertung ORDER BY AUS_ID DESC;"
    Set AbgegebeneWertungen = db.OpenRecordset(sql, DB_OPEN_DYNASET)
    rd = ch_runde(trunde)

    Set vars = zerlege(Nz(AbgegebeneWertungen!Cgi_Input))

    Letzte_Runde = CSng(vars.Item("rt_ID"))
    letzte_Tanzrunde = CSng(vars.Item("rh1"))
    
    Set AbgegebeneWertungen = db.OpenRecordset("SELECT Wert_Richter.WR_Nachname, Wert_Richter.WR_func, Paare_Rundenqualifikation.TP_ID AS Paar_ID, Paare_Rundenqualifikation.PR_ID, Auswertung.Cgi_Input FROM (Auswertung INNER JOIN Wert_Richter ON Auswertung.WR_ID = Wert_Richter.WR_ID) INNER JOIN Paare_Rundenqualifikation ON Auswertung.PR_ID = Paare_Rundenqualifikation.PR_ID WHERE (((Auswertung.Cgi_Input) Like '*rh1=" & letzte_Tanzrunde & "*' AND (Auswertung.Cgi_Input) Like '*rt_ID=" & Letzte_Runde & "*'));", DB_OPEN_DYNASET)
'    Set AbgegebeneWertungen = db.OpenRecordset("SELECT Wert_Richter.WR_Kuerzel, Wert_Richter.WR_Nachname, Startklasse_Wertungsrichter.WR_function, AW.* FROM (Startklasse_Wertungsrichter INNER JOIN (Rundentab INNER JOIN Abgegebene_Wertungen AS AW ON Rundentab.RT_ID = AW.RundenTab_ID) ON Startklasse_Wertungsrichter.Startklasse = Rundentab.Startklasse) INNER JOIN Wert_Richter ON (Startklasse_Wertungsrichter.WR_ID = Wert_Richter.WR_ID) AND (AW.Wertungsrichter_ID = Wert_Richter.WR_ID) WHERE (((AW.rh)=" & letzte_Tanzrunde & ") AND ((AW.RundenTab_ID)=" & Letzte_Runde & ")) ORDER BY AW.Paar_ID, Wert_Richter.WR_Kuerzel;", DB_OPEN_DYNASET)
    Set Paar_Infos = db.OpenRecordset("SELECT RT.RT_ID, PRQ.Rundennummer, PRQ.TP_ID, RT.Startklasse, Startklasse.Startklasse_text, RT.Runde, Paare.Startnr, Paare.Startnr, IIf([isTeam],[Name_Team],[Da_Nachname]) AS Ausdr1, Paare.He_Nachname FROM (Paare INNER JOIN (Rundentab AS RT INNER JOIN Paare_Rundenqualifikation AS PRQ ON RT.RT_ID = PRQ.RT_ID) ON Paare.TP_ID = PRQ.TP_ID) INNER JOIN Startklasse ON RT.Startklasse = Startklasse.Startklasse WHERE (((RT.RT_ID)=" & Letzte_Runde & ") AND ((PRQ.Rundennummer)=" & letzte_Tanzrunde & ")) ORDER BY RT.RT_ID, PRQ.Rundennummer, Paare.Startnr;", DB_OPEN_DYNASET)

    A_WR(1) = 1
    T_WR(1) = 1
    A_WR(2) = 1
    T_WR(2) = 1

    T_WR_Reset = False
    A_WR_Reset = False
    Paar_ID = CSng(vars.Item("PR_ID1"))
    st_kl = Paar_Infos!Startklasse

    
    HTML_Website = ""
    If left(Paar_Infos!Startklasse, 3) = "BW_" Then ' Or Left(Paar_Infos!Startklasse, 3) = "F_B" Then
        HTML_überschrift = "<table border='1' cellpadding='1' cellspacing='1' style='width: 1024px; text-align: center;'><tbody><tr bgcolor=#d0d0d0><td>Name</td><td>Grundschritt</td><td>Basic Dancing</td><td>Tanzfig</td><td>Interpret</td><td>Spontane Int</td><td>Dance Perf</td><td>Summe</td></tr><tr><td>TWR1</td></tr><tr><td>TWR2</td></tr><tr><td>TWR3</td></tr><tr><td>TWR4</td></tr><tr><td>TWR5</td></tr><tr><td>TWR6</td></tr><tr><td>TWR7</td></tr><tr></tr>"
        HTML_überschrift = HTML_überschrift & "</tbody></table>"
        HTML_WR_Template = "<td>WRNAME</td><td>WERT01</td><td>WERT02</td><td>WERT03</td><td>WERT05</td><td>WERT06</td><td>WERT07</td><td>WERT08</td>"
    Else
        HTML_überschrift = "<table border='1' cellpadding='1' cellspacing='1' style='width: 1024px;'><tbody><tr bgcolor=#d0d0d0><td>Name</td><td>GT&nbsp;H</td><td>HD&nbsp;H</td><td>GT&nbsp;D</td><td>HD&nbsp;D</td><td>Chor</td><td>Tanzf.</td><td>T&auml;nzD</td><td>Summe</td><td>&nbsp;</td><td>Grobf</td><td>Abzüge</td><td>&nbsp;</td><td>Punkte</td></tr><tr><td>TWR1</td></tr><tr><td>TWR2</td></tr><tr><td>TWR3</td></tr><tr><td>TWR4</td></tr><tr><td>TWR5</td></tr><tr></tr>"
        HTML_überschrift = HTML_überschrift & "<tr bgcolor=#d0d0d0><td>Name</td><td>Akro1</td><td>GF&nbsp;1</td><td>Akro2</td><td>GF&nbsp;2</td><td>Akro3</td><td>GF&nbsp;3</td><td>Akro4</td><td>GF&nbsp;4</td><td>Akro5</td><td>GF&nbsp;5</td><td>Akro6</td><td>GF&nbsp;6</td><td>Akro7</td><td>GF&nbsp;7</td><td>Akro8</td><td>GF&nbsp;8</td><td>Punkte</td></tr><tr><td>AWR1</td></tr><tr><td>AWR2</td></tr><tr><td>AWR3</td></tr><tr><td>AWR4</td></tr><tr></tr></tbody></table>"
        HTML_WR_Template = "<td>WRNAME</td><td>WERT01</td><td>WERT02</td><td>WERT03</td><td>WERT04</td><td>WERT05</td><td>WERT06</td><td>WERT07</td><td>WERT08</td><td>WERT09</td><td>WERT10</td><td>WERT11</td><td>WERT12</td><td>WERT13</td><td>WERT14</td><td>WERT15</td><td>WERT16</td><td>WERT17</td>"
    End If
    
    '**** neue Seite erzeugen
    HTML_Paar_links = HTML_überschrift
    HTML_Paar_rechts = HTML_überschrift
    
    AbgegebeneWertungen.MoveLast
    WR_Zaehler = AbgegebeneWertungen.RecordCount
    AbgegebeneWertungen.MoveFirst
    
    
    For X = 1 To WR_Zaehler
        If Paar_ID = AbgegebeneWertungen!Paar_ID Then
            seite = 1
        Else
            seite = 2
        End If
        Set vars = zerlege(Nz(AbgegebeneWertungen!Cgi_Input))

            HTML_WR_Werte = HTML_WR_Template
        
            Paar_Infos.FindFirst "TP_ID = " & AbgegebeneWertungen!Paar_ID
        
             If AbgegebeneWertungen!WR_func = "Ft" Or AbgegebeneWertungen!WR_func = "X" Then
           
                GesamtPunkte = 0
                Select Case left(st_kl, 3)
                    Case "BW_"
                        kl_punkte = Punkteverteilung(st_kl, rd, trunde)
                        GesamtPunkte = CSng(AbgegebeneWertungen!Herr_Grundtechnik) * kl_punkte(0) / 10
                        GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Herr_Haltung_Drehtechnik) * kl_punkte(1) / 10
                        GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Dame_Grundtechnik) * kl_punkte(2) / 10
                        GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Choreographie) * kl_punkte(4) / 10
                        GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Tanzfiguren) * kl_punkte(5) / 10
                        GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Tänzerische_Darbietung) * kl_punkte(6) / 10
                    Case "RR_"
                        kl_punkte = Punkteverteilung(st_kl, rd, trunde)
                        If rd = "ER" Then
                            GesamtPunkte = GesamtPunkte + CSng(vars.Item("wsh" & seite)) * kl_punkte(0) / 10
                            GesamtPunkte = GesamtPunkte + CSng(vars.Item("wth" & seite)) * kl_punkte(1) / 10
                            GesamtPunkte = GesamtPunkte + CSng(vars.Item("wsd" & seite)) * kl_punkte(2) / 10
                            GesamtPunkte = GesamtPunkte + CSng(vars.Item("wtd" & seite)) * kl_punkte(3) / 10
                            GesamtPunkte = GesamtPunkte + CSng(vars.Item("wch" & seite)) * kl_punkte(4) / 10
                            GesamtPunkte = GesamtPunkte + CSng(vars.Item("wtf" & seite)) * kl_punkte(5) / 10
                            GesamtPunkte = GesamtPunkte + CSng(vars.Item("wda" & seite)) * kl_punkte(6) / 10
                        Else
                            GesamtPunkte = GesamtPunkte + CSng(vars.Item("wsh" & seite)) * kl_punkte(0) / 10 * 2
                            GesamtPunkte = GesamtPunkte + CSng(vars.Item("wsd" & seite)) * kl_punkte(2) / 10 * 2
                            GesamtPunkte = GesamtPunkte + CSng(vars.Item("wch" & seite)) * kl_punkte(4) / 10
                            GesamtPunkte = GesamtPunkte + CSng(vars.Item("wch" & seite)) * kl_punkte(5) / 10
                            GesamtPunkte = GesamtPunkte + CSng(vars.Item("wch" & seite)) * kl_punkte(6) / 10
                        End If
                        HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT01", Observer_FT(st_kl, vars.Item("wsh" & seite), 0, trunde))
                        HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT02", Observer_FT(st_kl, vars.Item("wth" & seite), 1, trunde))
                        HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT03", Observer_FT(st_kl, vars.Item("wsd" & seite), 2, trunde))
                        HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT04", Observer_FT(st_kl, vars.Item("wtd" & seite), 3, trunde))
                        HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT05", Observer_FT(st_kl, vars.Item("wch" & seite), 4, trunde))
                        HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT06", Observer_FT(st_kl, vars.Item("wtf" & seite), 5, trunde))
                        HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT07", Observer_FT(st_kl, vars.Item("wda" & seite), 6, trunde))
                        HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT10", vars.Item("tfl" & seite))
                        HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT11", CSng(vars.Item("wfl" & seite)))
                        HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT13", Round(GesamtPunkte - CSng(vars.Item("wfl" & seite)), 2))
                    Case "F_B"
                        kl_punkte = Punkteverteilung(st_kl, rd, trunde)
                        GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Herr_Grundtechnik) * kl_punkte(0) / 10
                        GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Herr_Haltung_Drehtechnik) * kl_punkte(1) / 10
                        GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Dame_Grundtechnik) * kl_punkte(2) / 10
                        GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Choreographie) * kl_punkte(4) / 10
                        GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Tanzfiguren) * kl_punkte(5) / 10
                        GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Tänzerische_Darbietung) * kl_punkte(6) / 10
                    Case "F_R"
                        kl_punkte = Punkteverteilung(st_kl, rd, trunde)
                        GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Herr_Grundtechnik) * kl_punkte(0) / 10
                        GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Herr_Haltung_Drehtechnik) * kl_punkte(1) / 10
                        GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Dame_Grundtechnik) * kl_punkte(2) / 10
                        GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Choreographie) * kl_punkte(4) / 10
                        GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Tanzfiguren) * kl_punkte(5) / 10
                        GesamtPunkte = GesamtPunkte + CSng(AbgegebeneWertungen!Tänzerische_Darbietung) * kl_punkte(6) / 10
                    Case Else
                        Select Case left(st_kl, 6)
                            Case "BS_BW_", "BS_F_R"
                                kl_punkte = Punkteverteilung(st_kl, rd, trunde)
                                GesamtPunkte = GesamtPunkte + CSng(vars.Item("wth" & seite)) * kl_punkte(0) / 10
                                GesamtPunkte = GesamtPunkte + CSng(vars.Item("wtd" & seite)) * kl_punkte(1) / 10
                                GesamtPunkte = GesamtPunkte + CSng(vars.Item("wta" & seite)) * kl_punkte(2) / 10
                                GesamtPunkte = GesamtPunkte + CSng(vars.Item("wak" & seite)) * kl_punkte(4) / 10
                                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT08", Round(GesamtPunkte, 2))
                                GesamtPunkte = GesamtPunkte - CSng(vars.Item("wfe" & seite))
                                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT01", Observer_FT(st_kl, vars.Item("wth" & seite), 0, trunde))
                                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT02", Observer_FT(st_kl, vars.Item("wtd" & seite), 1, trunde))
                                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT03", Observer_FT(st_kl, vars.Item("wta" & seite), 2, trunde))
                                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT04", Observer_FT(st_kl, vars.Item("wak" & seite), 3, trunde))
                                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT05", "&nbsp;")
                                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT06", "&nbsp;")
                                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT07", "&nbsp;")
                               
                                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT10", Observer_FT(st_kl, vars.Item("tfe" & seite), 11, trunde))
                                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT11", Observer_FT(st_kl, vars.Item("wfe" & seite), 10, trunde))
                                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT13", Round(GesamtPunkte, 2))
                            Case Else
                                GesamtPunkte = (CSng(AbgegebeneWertungen!Herr_Grundtechnik) / 2) + (CSng(Nz(AbgegebeneWertungen!Herr_Haltung_Drehtechnik)) / 2) + (CSng(AbgegebeneWertungen!Dame_Grundtechnik) / 2)
                                GesamtPunkte = GesamtPunkte + (CSng(Nz(AbgegebeneWertungen!Dame_Haltung_Drehtechnik)) / 2) + (CSng(AbgegebeneWertungen!Choreographie) * 6 / 10) + (CSng(Nz(AbgegebeneWertungen!Tanzfiguren) * 6 / 10)) + (CSng(AbgegebeneWertungen!Tänzerische_Darbietung) * 8 / 10)
                        End Select
                End Select
                If GesamtPunkte < 0 Then GesamtPunkte = 0
                
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WRNAME", AbgegebeneWertungen!WR_Nachname)
                'einen FT-WR einfügen
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT01", Observer_FT(st_kl, AbgegebeneWertungen!Herr_Grundtechnik, 0, trunde))
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT02", Observer_FT(st_kl, AbgegebeneWertungen!Herr_Haltung_Drehtechnik, 1, trunde))
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT03", Observer_FT(st_kl, AbgegebeneWertungen!Dame_Grundtechnik, 2, trunde))
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT04", Observer_FT(st_kl, AbgegebeneWertungen!Dame_Haltung_Drehtechnik, 3, trunde))
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT05", Observer_FT(st_kl, AbgegebeneWertungen!Choreographie, 4, trunde))
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT06", Observer_FT(st_kl, AbgegebeneWertungen!Tanzfiguren, 5, trunde))
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT07", Observer_FT(st_kl, AbgegebeneWertungen!Tänzerische_Darbietung, 6, trunde))
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
                
'                If AbgegebeneWertungen!Paar_ID <> Paar_ID And T_WR_Reset = False Then
'                    T_WR = 1
'                    T_WR_Reset = True
'                End If
                If AbgegebeneWertungen!Paar_ID = Paar_ID Then
                    HTML_Paar_links = Replace(HTML_Paar_links, "<td>TWR" & T_WR(1) & "</td>", HTML_WR_Werte)
                    PaarLinks = Paar_Infos!Startnr & " " & Paar_Infos!Ausdr1 & " / " & Paar_Infos!He_Nachname
                    T_WR(1) = T_WR(1) + 1
                Else
                    HTML_Paar_rechts = Replace(HTML_Paar_rechts, "<td>TWR" & T_WR(2) & "</td>", HTML_WR_Werte)
                    PaarRechts = Paar_Infos!Startnr & " " & Paar_Infos!Ausdr1 & " / " & Paar_Infos!He_Nachname
                    T_WR(2) = T_WR(2) + 1
                End If
            Else
                
                GesamtPunkte = 0
                For i = 1 To 8
'                    If Not IsNull(AbgegebeneWertungen!Akrobatik1) Then
                        GesamtPunkte = GesamtPunkte + CSng(vars.Item("wak" & seite & i)) - CSng(vars.Item("wfl" & seite & "_ak" & seite & i))
                        t = t + 1
'                    End If
                Next
                If st_kl = "F_RR_M" Then ' Master RR
                    GesamtPunkte = GesamtPunkte / IIf(t < 7, 6, t) * 5
                Else
                    GesamtPunkte = 0
                End If
                If GesamtPunkte < 0 Then GesamtPunkte = 0
                
                'einen AK-WR einfügen
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WRNAME", AbgegebeneWertungen!WR_Nachname)
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT01", Round((100 - (CSng(vars.Item("wak" & seite & "1")) * 100 / Get_Akropunkte(AbgegebeneWertungen!Paar_ID, Paar_Infos!Runde, 1))), 0) & " ")
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT02", vars.Item("wfl" & seite & "_ak" & seite & "1") & " ")
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT03", Round((100 - (CSng(vars.Item("wak" & seite & "2")) * 100 / Get_Akropunkte(AbgegebeneWertungen!Paar_ID, Paar_Infos!Runde, 2))), 0) & " ")
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT04", vars.Item("wfl" & seite & "_ak" & seite & "2") & " ")
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT05", Round((100 - (CSng(vars.Item("wak" & seite & "3")) * 100 / Get_Akropunkte(AbgegebeneWertungen!Paar_ID, Paar_Infos!Runde, 3))), 0) & " ")
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT06", vars.Item("wfl" & seite & "_ak" & seite & "3") & " ")
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT07", Round((100 - (CSng(vars.Item("wak" & seite & "4")) * 100 / Get_Akropunkte(AbgegebeneWertungen!Paar_ID, Paar_Infos!Runde, 4))), 0) & " ")
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT08", vars.Item("wfl" & seite & "_ak" & seite & "4") & " ")
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT09", Round((100 - (CSng(vars.Item("wak" & seite & "5")) * 100 / Get_Akropunkte(AbgegebeneWertungen!Paar_ID, Paar_Infos!Runde, 5))), 0) & " ")
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT10", vars.Item("wfl" & seite & "_ak" & seite & "5") & " ")
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT11", Round((100 - (CSng(vars.Item("wak" & seite & "6")) * 100 / Get_Akropunkte(AbgegebeneWertungen!Paar_ID, Paar_Infos!Runde, 6))), 0) & " ")
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT12", vars.Item("wfl" & seite & "_ak" & seite & "6") & " ")
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT13", Round((100 - (CSng(vars.Item("wak" & seite & "7")) * 100 / Get_Akropunkte(AbgegebeneWertungen!Paar_ID, Paar_Infos!Runde, 7))), 0) & " ")
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT14", vars.Item("wfl" & seite & "_ak" & seite & "7") & " ")
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT15", Round((100 - (CSng(vars.Item("wak" & seite & "8")) * 100 / Get_Akropunkte(AbgegebeneWertungen!Paar_ID, Paar_Infos!Runde, 8))), 0) & " ")
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT16", vars.Item("wfl" & seite & "_ak" & seite & "8") & " ")
                HTML_WR_Werte = Replace(HTML_WR_Werte, "WERT17", Round(GesamtPunkte, 2))
               
                If AbgegebeneWertungen!WR_func <> "Ob" Then
                    If AbgegebeneWertungen!Paar_ID = Paar_ID Then
                        HTML_Paar_links = Replace(HTML_Paar_links, "<td>AWR" & A_WR(1) & "</td>", HTML_WR_Werte)
                        A_WR(1) = A_WR(1) + 1
                    Else
                        HTML_Paar_rechts = Replace(HTML_Paar_rechts, "<td>AWR" & A_WR(2) & "</td>", HTML_WR_Werte)
                        A_WR(2) = A_WR(2) + 1
                    End If
                End If
            End If
    
'        Next seite
        AbgegebeneWertungen.MoveNext
    
    Next X
        
    HTML_Website = "<H1>" & Paar_Infos!Startklasse_text & " " & Paar_Infos!Runde & "</H1><H2>" & PaarLinks & "</H2>" & HTML_Paar_links & "<br><br><H2>" & PaarRechts & "</H2>" & HTML_Paar_rechts & "<br><br>"
    
    Dim out
    Dim pfad As String
    Dim Server_IP
    
    'Server_IP = GetIpAddrTable()
    
    pfad = getBaseDir & "Apache2\htdocs\observer\index.html"
    Set out = file_handle(pfad)
    out.WriteLine ("<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3.org/TR/html4/loose.dtd""><html><head><title>" & Forms![A-Programmübersicht]!Turnierbez & "Observer Übersicht" & "</title> <meta http-equiv=""refresh"" content=""5""; URL="" " & GetIpAddrTable() & "/observer.html""><meta http-equiv=""expires"" content=""0""></head><body>" & HTML_Website & "</body></html>")
    out.Close

ObserverHTML_Fehler_Exit:
        'Funktion verlassen
        Exit Sub
        
ObserverHTML_Fehler_Err:
    Resume Next



End Sub

Public Function Get_Akropunkte(TP_ID, Runde, akronummer)
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
    
    AkroText = "Wert" & akronummer & RundTxt
    Get_Akropunkte = Paare(AkroText)
    
End Function

Public Function output_for_BRBV()
    Dim db As Database
    Dim re As Recordset
    Dim sql As String
    Dim HTML_Seite As String
    Dim Paar As String
    Dim rde As String
    Dim i, j As Integer
    Dim Punkte As Single
    Dim vars
    Dim kl_punkte As Variant
    Dim wr As String
    Set db = CurrentDb
    
    sql = "SELECT Auswertung.AUS_ID, Auswertung.Cgi_Input, First(Auswertung.reihenfolge) AS ErsterWertvonreihenfolge, Paare_Rundenqualifikation.TP_ID, Auswertung.Punkte, Paare.Da_Nachname, Paare.He_Nachname, Startklasse.Startklasse_text, Tanz_Runden_fix.Rundentext, Auswertung.WR_ID, Majoritaet.Platz, Majoritaet.WR7, Paare.Startnr, Rundentab.Startklasse "
    sql = sql & "FROM Wert_Richter RIGHT JOIN ((Tanz_Runden_fix RIGHT JOIN (Startklasse RIGHT JOIN Rundentab ON Startklasse.Startklasse = Rundentab.Startklasse) ON Tanz_Runden_fix.Runde = Rundentab.Runde) INNER JOIN (Paare INNER JOIN (((Auswertung LEFT JOIN Startklasse_Wertungsrichter ON Auswertung.WR_ID = Startklasse_Wertungsrichter.WR_ID) INNER JOIN Paare_Rundenqualifikation ON Auswertung.PR_ID = Paare_Rundenqualifikation.PR_ID) INNER JOIN Majoritaet ON (Paare_Rundenqualifikation.RT_ID = Majoritaet.RT_ID) AND (Paare_Rundenqualifikation.TP_ID = Majoritaet.TP_ID)) ON Paare.TP_ID = Paare_Rundenqualifikation.TP_ID) ON Rundentab.RT_ID = Paare_Rundenqualifikation.RT_ID) ON Wert_Richter.WR_ID = Startklasse_Wertungsrichter.WR_ID "
    sql = sql & "WHERE (((Startklasse_Wertungsrichter.WR_function)<>'Ob') AND ((Wert_Richter.WR_Azubi)=False)) "
    sql = sql & "GROUP BY Auswertung.AUS_ID, Auswertung.Cgi_Input, Paare_Rundenqualifikation.TP_ID, Auswertung.Punkte, Paare.Da_Nachname, Paare.He_Nachname, Startklasse.Startklasse_text, Tanz_Runden_fix.Rundentext, Auswertung.WR_ID, Majoritaet.Platz, Majoritaet.WR7, Paare.Startnr, Rundentab.Startklasse, Rundentab.Rundenreihenfolge "
    sql = sql & "HAVING (((Rundentab.Startklasse) Like 'BS_BY_*')) "
    sql = sql & "ORDER BY Rundentab.Startklasse, Rundentab.Rundenreihenfolge, Majoritaet.Platz, Paare.Startnr, Auswertung.Punkte;"
    
    Set re = db.OpenRecordset(sql)
    re.MoveFirst
    Paar = 0
    rde = ""
    HTML_Seite = "<!DOCTYPE html><html><head><style>"
    HTML_Seite = HTML_Seite & "table {font-family: Arial; FONT-WEIGHT: bold; border-collapse: collapse; text-align: center; margin-left: auto; margin-right: auto;}" & vbCrLf
    HTML_Seite = HTML_Seite & "td {border: thin solid red; background-color: white;}"
    HTML_Seite = HTML_Seite & "td.paar_line { FONT-SIZE: 11pt; FONT-WEIGHT: bold; TEXT-ALIGN: center; color: black; BORDER-BOTTOM: #ccc 1px solid; padding: 5px;}" & vbCrLf
    HTML_Seite = HTML_Seite & ".paar_detail { FONT-SIZE: 10pt; color: #E86E0E;}"
    HTML_Seite = HTML_Seite & "td.werte { FONT-SIZE: 11pt; FONT-WEIGHT: bold; COLOR: navy; TEXT-ALIGN: center; BACKGROUND-COLOR: #acd; padding-left: 5px;padding-right: 5px; }" & vbCrLf
    HTML_Seite = HTML_Seite & "</style></head>"
    HTML_Seite = HTML_Seite & "<body style=""background-color: #347;""><table>"

'    HTML_Seite = HTML_Seite & "<tr><td colspan=""8"" style=""font-size: 45px;"">" & Umlaute_Umwandeln(Forms![A-Programmübersicht]!Turnierbez) & "</td></tr>"
    Do Until re.EOF
        If rde <> re!Rundentext Then
            HTML_Seite = HTML_Seite & "<tr><td colspan=""8"" height=""15px""></td></tr>"
            HTML_Seite = HTML_Seite & "<tr><td colspan=""8"">" & re!Startklasse_text & " - " & re!Rundentext & "</td></tr>"
            HTML_Seite = HTML_Seite & "<tr><td class=""werte"">Paar</td><td class=""werte"">Platz</td>x__wr<td class=""werte"">Punkte</td></tr>" & vbCrLf
            rde = re!Rundentext
            kl_punkte = Punkteverteilung("BS_BY_BJ", "ER", "")
        End If
        HTML_Seite = HTML_Seite & "<tr>"
        HTML_Seite = HTML_Seite & "<td class=""paar_line"">" & Umlaute_Umwandeln(re!Da_NAchname & " - " & re!He_Nachname) & "</td>"
        HTML_Seite = HTML_Seite & "<td class=""paar_line"">" & re!Platz & "</td>"
        Punkte = 0
        j = 0
        wr = ""
        Paar = re!Rundentext & re!TP_ID
        Do Until re!Rundentext & re!TP_ID <> Paar
            j = j + 1
            Set vars = zerlege(DLookup("cgi_input", "auswertung", "AUS_ID =" & re!AUS_ID))
            i = eins_zwei(re!TP_ID, vars)
            Punkte = re!WR7
            wr = wr & "<td class=""werte"">WR " & j & "</td>"
            HTML_Seite = HTML_Seite & "<td class=""paar_line"" >" & Format(CSng(vars.Item("Punkte" & i)), "##0.00") & "<br><span class=""paar_detail"">"
            Select Case left(re!Startklasse, 6)
                Case "BW_JA", "BW_MA", "BW_MB", "BW_SA", "BW_SB"
                    kl_punkte = Punkteverteilung("BW_NG", "ER", "")
                    HTML_Seite = HTML_Seite & CSng(vars.Item("wng_ttd" & i)) * kl_punkte(0) & "&#124;"
                    HTML_Seite = HTML_Seite & CSng(vars.Item("wng_tth" & i)) * kl_punkte(1) & "&#124;"
                    HTML_Seite = HTML_Seite & CSng(vars.Item("wng_bda" & i)) * kl_punkte(2) & "&#124;"
                    HTML_Seite = HTML_Seite & CSng(vars.Item("wng_dap" & i)) * kl_punkte(3) & "&#124;"
                    HTML_Seite = HTML_Seite & CSng(vars.Item("wng_bdb" & i)) * kl_punkte(4) & "&#124;"
                    HTML_Seite = HTML_Seite & CSng(vars.Item("wng_fta" & i)) * kl_punkte(5) & "&#124;"
                    HTML_Seite = HTML_Seite & CSng(vars.Item("wng_fts" & i)) * kl_punkte(6) & "&#124;"
                    HTML_Seite = HTML_Seite & CSng(vars.Item("wng_ftb" & i)) * kl_punkte(7) & "&#124;"
                    HTML_Seite = HTML_Seite & CSng(vars.Item("wng_inf" & i)) * kl_punkte(8) & "&#124;"
                    HTML_Seite = HTML_Seite & CSng(vars.Item("wng_ins" & i)) * kl_punkte(9) & "&#124;"
                    HTML_Seite = HTML_Seite & CSng(vars.Item("wng_inb" & i)) * kl_punkte(10)
                Case "BS_BY_"
                    kl_punkte = Punkteverteilung("BS_BY_BJ", "ER", "")
                    If vars.exists("wgs" & i) Then     ' für Turniere von dem 01.08.2024
                        HTML_Seite = HTML_Seite & CSng(vars.Item("wgs" & i)) * (kl_punkte(0) + kl_punkte(1)) / 10 & "&#124;"
                    Else
                        HTML_Seite = HTML_Seite & CSng(vars.Item("wsd" & i)) * kl_punkte(0) / 10 & "&#124;"
                        HTML_Seite = HTML_Seite & CSng(vars.Item("wsh" & i)) * kl_punkte(1) / 10 & "&#124;"
                    End If
                    HTML_Seite = HTML_Seite & CSng(vars.Item("wbd" & i)) * kl_punkte(2) / 10 & "&#124;"
                    HTML_Seite = HTML_Seite & CSng(vars.Item("wtf" & i)) * kl_punkte(3) / 10 & "&#124;"
                    HTML_Seite = HTML_Seite & CSng(vars.Item("win" & i)) * kl_punkte(4) / 10
                Case "BS_BW_", "BS_F_R"
                    kl_punkte = Punkteverteilung("BS_BY_BJ", "ER", "")
                    HTML_Seite = HTML_Seite & CSng(vars.Item("wth" & i)) & "&#124;"
                    HTML_Seite = HTML_Seite & CSng(vars.Item("wtd" & i)) & "&#124;"
                    HTML_Seite = HTML_Seite & CSng(vars.Item("wta" & i)) & "&#124;"
                    HTML_Seite = HTML_Seite & CSng(vars.Item("wak" & i)) & "&#124;"
                    HTML_Seite = HTML_Seite & CSng(vars.Item("wfe" & i)) * -1
            End Select
            HTML_Seite = HTML_Seite & "</span></td>"
            re.MoveNext
            If re.EOF Then Exit Do
        Loop
        HTML_Seite = HTML_Seite & "<td class=""paar_line"">" & Format(Punkte, "##0.00") & "</td>"
        HTML_Seite = HTML_Seite & "</tr>" & vbCrLf
    Loop
    HTML_Seite = HTML_Seite & "<tr><td colspan=""8"" height=""25px""></td></tr>"
    HTML_Seite = Replace(HTML_Seite, "x__wr", wr)
    re.MoveFirst
    Select Case left(re!Startklasse, 6)
        Case "BS_BY_"
            If vars.exists("wgs" & i) Then        ' für Turniere von dem 01.08.2024
                HTML_Seite = HTML_Seite & "<tr><td class=""paar_detail"">Wert 1</td><td colspan=""8"" class=""paar_detail"">Grundschritt (Rhythmus & Fu&szlig;technik)</td></tr>"
                HTML_Seite = HTML_Seite & "<tr><td class=""paar_detail"">Wert 2</td><td colspan=""8"" class=""paar_detail"">Basic Dancing, Lead & Follow, Harmonie</td></tr>"
                HTML_Seite = HTML_Seite & "<tr><td class=""paar_detail"">Wert 3</td><td colspan=""8"" class=""paar_detail"">Tanzfiguren (einfache, highlight)</td></tr>"
                HTML_Seite = HTML_Seite & "<tr><td class=""paar_detail"">Wert 4</td><td colspan=""8"" class=""paar_detail"">Interpretation (Figuren, spontane Interpretation)</td></tr>"
            Else
                HTML_Seite = HTML_Seite & "<tr><td class=""paar_detail"">Wert 1</td><td colspan=""8"" class=""paar_detail"">Grundschritt Follower (Rhythmus & Fu&szlig;technik)</td></tr>"
                HTML_Seite = HTML_Seite & "<tr><td class=""paar_detail"">Wert 2</td><td colspan=""8"" class=""paar_detail"">Grundschritt Leader (Rhythmus & Fu&szlig;technik)</td></tr>"
                HTML_Seite = HTML_Seite & "<tr><td class=""paar_detail"">Wert 3</td><td colspan=""8"" class=""paar_detail"">Basic Dancing, Lead & Follow, Harmonie</td></tr>"
                HTML_Seite = HTML_Seite & "<tr><td class=""paar_detail"">Wert 4</td><td colspan=""8"" class=""paar_detail"">Tanzfiguren (einfache, highlight)</td></tr>"
                HTML_Seite = HTML_Seite & "<tr><td class=""paar_detail"">Wert 5</td><td colspan=""8"" class=""paar_detail"">Interpretation (Figuren, spontane Interpretation)</td></tr>"
            End If
        Case "BS_BW_", "BS_F_R"
            HTML_Seite = HTML_Seite & "<tr><td class=""paar_detail"">Wert 1</td><td colspan=""8"" class=""paar_detail"">Technik Herr/Technik Formationen</td></tr>"
            HTML_Seite = HTML_Seite & "<tr><td class=""paar_detail"">Wert 2</td><td colspan=""8"" class=""paar_detail"">Technik Dame</td></tr>"
            HTML_Seite = HTML_Seite & "<tr><td class=""paar_detail"">Wert 3</td><td colspan=""8"" class=""paar_detail"">Tanz/Tanz Form</td></tr>"
            HTML_Seite = HTML_Seite & "<tr><td class=""paar_detail"">Wert 4</td><td colspan=""8"" class=""paar_detail"">Akrobatik</td></tr>"
            HTML_Seite = HTML_Seite & "<tr><td class=""paar_detail"">Wert 5</td><td colspan=""8"" class=""paar_detail"">Abz&#252;ge</td></tr>"
    End Select
    HTML_Seite = HTML_Seite & "<tr><td>Punkte</td><td colspan=""8"">bei geteilter Endrunde Summe der Hin- und R&#252;ckrunde</td></tr><tr>"
    HTML_Seite = HTML_Seite
    HTML_Seite = HTML_Seite
    HTML_Seite = HTML_Seite & "</table></body></html>"
    HTML_Seite = HTML_Seite
'    Debug.Print HTML_Seite
    
    Open gen_Ordner(getBaseDir & "_Versand\") & Forms![A-Programmübersicht]!Turnierbez & ".html" For Output As #1
    Print #1, HTML_Seite
    Close #1

End Function



