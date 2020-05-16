Option Compare Database
Option Explicit

    Public Const tr = "                   <tr>"
    Public Const trn = "                   </tr>"
    Dim a_check As String

    Private Declare Function GetIpAddrTable_API Lib "IpHlpApi" Alias "GetIpAddrTable" (pIPAddrTable As Any, pdwSize As Long, ByVal bOrder As Long) As Long
    
Sub build_html(pr, RT_nr, Runde)
    Dim out
    Dim line As String
    Dim db As Database
    Dim wr, re As Recordset
    Dim ht As Recordset
    Dim dNeu, dNext As String
    Dim tr_nr As String
    Dim a_check As String
    Dim ht_pfad, vb_pfad As String
    Dim rd_klasse As String
    Dim rd As Integer
    Dim sei_1, sei_2 As Integer
    Dim rmax As Integer
    Dim a_paare As Integer
    Dim ppr As Integer
    Dim IpAddrs
    Dim fil As Boolean
    

    ' Pfade generieren
    ht_pfad = getBaseDir & "Apache2\htdocs\"
    vb_pfad = getBaseDir & "Apache2\cgi-bin\"
    tr_nr = "T" & Forms![A-Programmübersicht]!Turnier_Nummer
    
    Set db = CurrentDb
    pr.MoveLast
'    If Left(rde, 3) = "BW_" Then
        Set wr = db.OpenRecordset("SELECT Rundentab.RT_ID, Wert_Richter.WR_ID, Wert_Richter.WR_Kuerzel, Wert_Richter.WR_Lizenznr, Left([WR_Nachname],1) & [WR_Vorname] AS Ausdr1, Wert_Richter.WR_tausch, WR_function, Startklasse.Startklasse_text, Startklasse.Startklasse, Rundentab.Anz_Paare, Tanz_Runden.Rundentext, Wert_Richter.WR_AzuBi FROM Wert_Richter LEFT JOIN (Startklasse RIGHT JOIN (Tanz_Runden RIGHT JOIN (Startklasse_Wertungsrichter LEFT JOIN Rundentab ON Startklasse_Wertungsrichter.Startklasse = Rundentab.Startklasse) ON Tanz_Runden.Runde = Rundentab.Runde) ON Startklasse.Startklasse = Startklasse_Wertungsrichter.Startklasse) ON Wert_Richter.WR_ID = Startklasse_Wertungsrichter.WR_ID WHERE (Rundentab.RT_ID=" & RT_nr & " OR Wert_Richter.WR_AzuBi=True);", DB_OPEN_DYNASET)
'    Else
'        Set wr = db.OpenRecordset("SELECT Rundentab.RT_ID, Wert_Richter.WR_ID, Wert_Richter.WR_Kuerzel, Wert_Richter.WR_Lizenznr, [WR_Nachname] & "" "" & [WR_Vorname] AS Ausdr1, Wert_Richter.WR_tausch, WR_function, Startklasse.Startklasse_text, Startklasse.Startklasse, Rundentab.Anz_Paare, Tanz_Runden.Rundentext, Wert_Richter.WR_AzuBi FROM Wert_Richter LEFT JOIN (Startklasse RIGHT JOIN (Tanz_Runden RIGHT JOIN (Startklasse_Wertungsrichter LEFT JOIN Rundentab ON Startklasse_Wertungsrichter.Startklasse = Rundentab.Startklasse) ON Tanz_Runden.Runde = Rundentab.Runde) ON Startklasse.Startklasse = Startklasse_Wertungsrichter.Startklasse) ON Wert_Richter.WR_ID = Startklasse_Wertungsrichter.WR_ID WHERE ((Rundentab.RT_ID=" & RT_nr & "  AND Startklasse_Wertungsrichter.WR_function<>'Ob') OR Wert_Richter.WR_AzuBi=True);", DB_OPEN_DYNASET)
'    End If
    If wr.RecordCount = 0 Then
        MsgBox "Die Wertungsrichtereinteilung ist noch nicht erfolgt!"
        Exit Sub
    End If
    
    a_paare = wr!Anz_Paare
    rd_klasse = wr!Rundentext & " - " & wr!Startklasse_text

    IpAddrs = GetIpAddrTable
    wr.MoveFirst
    
    
    Do Until wr.EOF
        rd = 0
        fil = False
        pr.Filter = "Anwesend_Status=1"
        Set re = pr.OpenRecordset
        re.MoveLast
        rmax = Int(re.RecordCount / a_paare) + re.RecordCount Mod a_paare
        re.MoveFirst
        Do Until re.EOF
            rd = rd + 1
            dNeu = tr_nr & "R" & wr!WR_Lizenznr & "_K" & RT_nr & "_" & rd     ' aktuelle HTML-Seite
            dNext = tr_nr & "R" & wr!WR_Lizenznr & "_K" & RT_nr & "_" & rd + 1 'Nächste HTML Seite
            Set out = file_handle(ht_pfad & dNeu & ".html")
            
            ppr = (a_paare = 1) Or ((rd * 2) - 1 = re.RecordCount) Or (((rd - rmax) * 2) - 1 = re.RecordCount)
            
            If re!Anz_Paare > 1 And (wr!WR_tausch = True And ppr = False) Then
                sei_1 = 2
                sei_2 = 1
            Else
                sei_1 = 1
                sei_2 = 2
            End If
            
            a_check = ""
            Select Case get_bs_erg(re!Startkl, 3)
                Case "F_B"  ' alle Boogie Formationen
                    line = get_line("BW", "Form", ppr)
                    line = Replace(line, "x__st1", re!Startnr)
                    line = Replace(line, "x__id1", re!TP_ID)
                    line = Replace(line, "x__rh1", rd)
'                    Alte Benennungen
'                    line = Replace(line, "x__sh1", kl_abstand & make_a_inpRRE("tk1", 10, "Tanztechnik - Grundschritt (Rhythmus und Fußtechnik) / Basic Dancing"))
'                    line = Replace(line, "x__th1", make_a_inpRRE("ch1", 10, "Tanzfiguren - komplexe Figuren, Highlightfiguren"))
'                    line = Replace(line, "x__sd1", make_a_inpRRE("tf1", 10, "Choreografie / Dance Performance (Aufbau, Musikinterpretation, Präsentation, Ausstrahlung)"))
'                    line = Replace(line, "x__td1", make_a_inpRRE("ab1", 10, "AF - Synchronität und Harmonie"))
'                    line = Replace(line, "x__ch1", make_a_inpRRE("aw1", 10, "AF - Bilder und Bildwechsel (Schwierigkeit, Ausführung)"))
'                    line = Replace(line, "x__da1", make_a_inpRRE("af1", 10, "AF - Formationsfiguren und Effekte") & kl_abstand)
                    
                    line = Replace(line, "x__sh1", kl_abstand & make_a_inpRRE("tk1", 10, "Technik - Grund, Haltungs und Drehtechnik"))
                    line = Replace(line, "x__th1", make_a_inpRRE("ch1", 10, "Tanz - Wertigkeit der Tanzfolge"))
                    line = Replace(line, "x__sd1", make_a_inpRRE("tf1", 10, "Tanz - Ausführung"))
                    line = Replace(line, "x__td1", make_a_inpRRE("ab1", 10, "AF - Wertigkeit"))
                    line = Replace(line, "x__ch1", make_a_inpRRE("aw1", 10, "AF - Ausführung)"))
                    line = Replace(line, "x__da1", make_a_inpRRE("af1", 10, "AF - Wirkung") & kl_abstand)
                    
                    Select Case wr!WR_function
                        Case "X"
                            line = Replace(line, "x__fuss", "")
                            line = Replace(line, "x__akro", " class=""in_vis""")
                            line = Replace(line, "x__obse", " class=""in_vis""")
                        Case "Ob"
                            line = Replace(line, "x__fuss", " class=""in_vis""")
                            line = Replace(line, "x__akro", " class=""in_vis""")
                            line = Replace(line, "x__obse", " class=""visi""")
                        Case Else
                        
                    End Select
                    line = Replace(line, "x__ck", "")
                
                Case "F_R", "F_F"   ' alle R'n'R Formationen
                    '   Ab hier neuer Bereich für getrennte Wertung
                    line = get_line("RR", "Form", ppr)
                    line = Replace(line, "x__st1", re!Startnr)
                    line = Replace(line, "x__id1", re!TP_ID)
                    line = Replace(line, "x__rh1", rd)
                    line = Replace(line, "x__sh1", kl_abstand & make_a_inpRRE("tk1", 10, "Technik - Grund-, Haltungs- und Drehtechnik"))
                    line = Replace(line, "x__th1", make_a_inpRRE("ch1", 10, "Tanz - Wert inkl. Formationsfiguren und Abstimmung zur Musik"))
                    line = Replace(line, "x__sd1", make_a_inpRRE("tf1", 10, "Tanz - Ausführung"))
                    line = Replace(line, "x__td1", make_a_inpRRE("ab1", 10, "AF - Wert der Bilder, Bildwechsel und Effekte"))
                    line = Replace(line, "x__ch1", make_a_inpRRE("aw1", 10, "AF - Ausführung"))
                    line = Replace(line, "x__da1", make_a_inpRRE("af1", 10, "Gesamtwirkung") & kl_abstand)
                    If wr!WR_function = "Ob" Then
                        line = Replace(line, "x__ak" & sei_1, make_akroOBS(re, Runde, sei_1, a_check, 8, a_paare))
                        line = Replace(line, "<input id=""absend"" type=""button"" class=""button_1"" value=""Absenden"" disabled ", "<input id=""absend"" type=""button"" class=""button_2"" value=""Absenden"" ")
                    Else
                        line = Replace(line, "x__ak" & sei_1, make_akroRRE(re, Runde, sei_1, a_check, 8, a_paare))
                    End If
                    line = Replace(line, "x__ber", akro_ber(re!Startkl, Runde))       ' für Berechnung in HTML-Seite
                    Select Case wr!WR_function
                        Case "Ft"
                            line = Replace(line, "x__fuss", "")
                            line = Replace(line, "x__akro", " class=""in_vis""")
                            line = Replace(line, "x__obse", " class=""in_vis""")
                        Case "Ak"
                            line = Replace(line, "x__fuss", " class=""in_vis""")
                            line = Replace(line, "x__akro", "")
                            line = Replace(line, "x__obse", " class=""in_vis""")
                        Case "Ob"
                            line = Replace(line, "x__fuss", " class=""in_vis""")
                            line = Replace(line, "x__akro", "")
                            line = Replace(line, "x__obse", " class=""in_vis""")
                        Case Else
                        
                    End Select
                    line = Replace(line, "x__ck", ", " & left(a_check, Len(a_check) - 2))
                
                Case "RR_"   ' alle RR einzel
                    ' ***** HM14.05 *****
                    ' kurze Wertung hinzugefügt
                    If a_paare = 1 Or wr!Rundentext = "Semifinale" Then
                        line = get_line("RR", "Seite", ppr)
                    Else
                        line = get_line("RR", "Seite_k", ppr)
                    End If
                    'x_nPg
                    line = Replace(line, "x__st" & sei_1, re!Startnr)
                    line = Replace(line, "x__id" & sei_1, re!TP_ID)
                    line = Replace(line, "x__rh" & sei_1, rd)
                    line = Replace(line, "x__sh1", kl_abstand & make_a_inpRRE("sh1", 10, "Technik Herr - Grundtechnik"))
                    line = Replace(line, "x__th1", make_a_inpRRE("th1", 10, "Technik Herr - Haltungs- und Drehtechnik"))
                    line = Replace(line, "x__sd1", make_a_inpRRE("sd1", 10, "Technik Dame - Grundtechnik"))
                    line = Replace(line, "x__td1", make_a_inpRRE("td1", 10, "Technik Dame - Haltungs- und Drehtechnik"))
                    line = Replace(line, "x__ch1", make_a_inpRRE("ch1", 10, "Tanz - Wertigkeit"))
                    line = Replace(line, "x__tf1", make_a_inpRRE("tf1", 10, "Tanz - Ausführung"))
                    line = Replace(line, "x__da1", make_a_inpRRE("da1", 10, "Tanz - Wirkung") & kl_abstand)
                    If wr!WR_function = "Ob" Then
                        line = Replace(line, "x__ak" & sei_1, make_akroOBS(re, Runde, sei_1, a_check, 6, a_paare))
                        line = Replace(line, "<input id=""absend"" type=""button"" class=""button_1"" value=""Absenden"" disabled ", "<input id=""absend"" type=""button"" class=""button_2"" value=""Absenden"" ")
                    Else
                        line = Replace(line, "x__ak" & sei_1, make_akroRRE(re, Runde, sei_1, a_check, 6, a_paare))
                    End If
                    line = Replace(line, "x__ber", akro_ber(re!Startkl, Runde))       ' für Berechnung in HTML-Seite
                    Select Case wr!WR_function
                        Case "Ft"
                            line = Replace(line, "x__fuss", "")
                            line = Replace(line, "x__akro", " class=""in_vis""")
                            line = Replace(line, "x__obse", " class=""in_vis""")
                        Case "Ak"
                            line = Replace(line, "x__fuss", " class=""in_vis""")
                            line = Replace(line, "x__akro", "")
                            line = Replace(line, "x__obse", " class=""in_vis""")
                        Case "Ob"
                            line = Replace(line, "x__fuss", " class=""in_vis""")
                            line = Replace(line, "x__akro", "")
                            line = Replace(line, "x__obse", " class=""in_vis""")
                        Case Else
                        
                    End Select
                    If Not ppr Then
                        re.MoveNext
                        line = Replace(line, "x__st" & sei_2, re!Startnr)
                        line = Replace(line, "x__id" & sei_2, re!TP_ID)
                        line = Replace(line, "x__rh" & sei_2, rd)
                        line = Replace(line, "x__sh2", kl_abstand & make_a_inpRRE("sh2", 10, "Technik Herr - Grundtechnik"))
                        line = Replace(line, "x__th2", make_a_inpRRE("th2", 10, "Technik Herr - Haltungs- und Drehtechnik"))
                        line = Replace(line, "x__sd2", make_a_inpRRE("sd2", 10, "Technik Dame - Grundtechnik"))
                        line = Replace(line, "x__td2", make_a_inpRRE("td2", 10, "Technik Dame - Haltungs- und Drehtechnik"))
                        line = Replace(line, "x__ch2", make_a_inpRRE("ch2", 10, "Tanz - Wertigkeit"))
                        line = Replace(line, "x__tf2", make_a_inpRRE("tf2", 10, "Tanz - Ausführung"))
                        line = Replace(line, "x__da2", make_a_inpRRE("da2", 10, "Tanz - Wirkung") & kl_abstand)
                        If wr!WR_function = "Ob" Then
                            line = Replace(line, "x__ak" & sei_2, make_akroOBS(re, Runde, sei_2, a_check, 6, a_paare))
                        Else
                            line = Replace(line, "x__ak" & sei_2, make_akroRRE(re, Runde, sei_2, a_check, 6, a_paare))
                        End If
                    End If
                    line = Replace(line, "x__ck", ", " & left(a_check, Len(a_check) - 2))
                Case "BW_"
                    ' Vorbereitung für kurze Wertung in den Vorrunden
                    If wr!WR_function = "Ob" Then
                        line = get_line("BW", "Observer", ppr)
                        line = fill_observer_verstoesse(line, re, ppr, RT_nr, sei_1, sei_2, wr!WR_ID)
                    Else
                       If ch_runde(Runde) = "VR" Then
                            line = get_line("BW", "Seite_k", ppr)  ' kurze Wertung
                        Else
                            line = get_line("BW", "Seite", ppr)
                        End If
                    End If
                    line = Replace(line, "x__st" & sei_1, re!Startnr)
                    line = Replace(line, "x__id" & sei_1, re!TP_ID)
                    line = Replace(line, "x__rh" & sei_1, rd)
                    If Not ppr Then
                        re.MoveNext
                        line = Replace(line, "x__st" & sei_2, re!Startnr)
                        line = Replace(line, "x__id" & sei_2, re!TP_ID)
                        line = Replace(line, "x__rh" & sei_2, rd)
                    End If
                
                Case "LH_"
                    line = get_line("LH", "Seite", ppr)
                    line = Replace(line, "x_st" & sei_1, re!Startnr)
                    line = Replace(line, "x_id" & sei_1, re!TP_ID)
                    line = Replace(line, "x_rh" & sei_1, rd)
                    If Not ppr Then
                        re.MoveNext
                        line = Replace(line, "x_st" & sei_2, re!Startnr)
                        line = Replace(line, "x_id" & sei_2, re!TP_ID)
                        line = Replace(line, "x_rh" & sei_2, rd)
                    End If
                
                Case "BS_"
                    If wr!WR_function = "Ob" Then
                        line = get_line("BS", "leer", ppr)
                    Else
                        If InStr(re!Startkl, "BS_BY_") > 0 Then
                            line = get_line("BS", "BS_BY", ppr)
                        Else
                            line = get_line("BS", "Seite", ppr)
                        End If
                    End If
                    line = Replace(line, "x__st" & sei_1, re!Startnr)
                    line = Replace(line, "x__id" & sei_1, re!TP_ID)
                    line = Replace(line, "x__rh" & sei_1, rd)
                    If Not ppr Then
                        re.MoveNext
                        line = Replace(line, "x__st" & sei_2, re!Startnr)
                        line = Replace(line, "x__id" & sei_2, re!TP_ID)
                        line = Replace(line, "x__rh" & sei_2, rd)
                    End If
                    
                Case "NBS_"
                    '----BS Nord------------------------------------------------------------------------------
                    Select Case left(re!Startkl, 6)
                        Case "BS_RR_"
                            line = get_line("BS", "NSeite", ppr)
                            'x_nPg
                            line = Replace(line, "x_st" & sei_1, re!Startnr)
                            line = Replace(line, "x_id" & sei_1, re!TP_ID)
                            line = Replace(line, "x_rh" & sei_1, rd)
                            line = Replace(line, "x_sh1", vorspan & vbCrLf & make_a_inp("sh1", 5))
                            line = Replace(line, "x_sd1", vorspan & vbCrLf & make_a_inp("sd1", 5))
                            line = Replace(line, "x_th1", vorspan & vbCrLf & make_a_inp("th1", 5))
                            line = Replace(line, "x_ck1", "")                   ' jetzt noch kein Check   a_check)
                            If Not ppr Then
                                a_check = ""
                                re.MoveNext
                                line = Replace(line, "x_st" & sei_2, re!Startnr)
                                line = Replace(line, "x_id" & sei_2, re!TP_ID)
                                line = Replace(line, "x_rh" & sei_2, rd)
                                line = Replace(line, "x_sh2", vorspan & vbCrLf & make_a_inp("sh2", 5))
                                line = Replace(line, "x_sd2", vorspan & vbCrLf & make_a_inp("sd2", 5))
                                line = Replace(line, "x_th2", vorspan & vbCrLf & make_a_inp("th2", 5))
                                line = Replace(line, "x_ck2", "")                  ' jetzt kein Check in Akro  a_check)
                            End If
                            line = Replace(line, "x_ber", 1)       ' für Berechnung in HTML-Seite
                    
                        Case "BS_F_R"
                            line = get_line("RR", "Form", ppr)
                            line = Replace(line, "x_st1", re!Startnr)
                            line = Replace(line, "x_id1", re!TP_ID)
                            line = Replace(line, "x_rh1", rd)
                            line = Replace(line, "x_sh1", vorspan & vbCrLf & make_a_inp("sh1", 10))
                            line = Replace(line, "x_th1", vorspan & vbCrLf & make_a_inp("th1", 10))
                            line = Replace(line, "x_ak1", make_akro(re, Runde, 1, a_check, 8, a_paare))
                            line = Replace(line, "x_ch1", vorspan & vbCrLf & make_a_inp("ch1", 10))
                            line = Replace(line, "x_ck1", "")                   ' jetzt kein Check in Akro  a_check)
                            line = Replace(line, "x_ber", "1")                 ' für Berechnung in HTML-Seite
                            
                        Case "BS_BW_"
                            line = get_line("BS", "Seite", ppr)
                            line = Replace(line, "x_st" & sei_1, re!Startnr)
                            line = Replace(line, "x_id" & sei_1, re!TP_ID)
                            line = Replace(line, "x_rh" & sei_1, rd)
                            If Not ppr Then
                                re.MoveNext
                                line = Replace(line, "x_st" & sei_2, re!Startnr)
                                line = Replace(line, "x_id" & sei_2, re!TP_ID)
                                line = Replace(line, "x_rh" & sei_2, rd)
                            End If
                            
                        Case Else
                            MsgBox "Fehler bei der BS-Selektion"
                            
                    End Select
                    '----/BS Nord-------------------------------------------------------------------------------------
                Case "BWBS_"
                    '----BS Baden-Württemberg-------------------------------------------------------------------------------------
                    Select Case left(re!Startkl, 6)
                        Case "BS_RR_E1", "BS_RR_J2"
                            line = get_line("BS", "BW_RR_lang", ppr)
                        
                        Case "BS_RR_BB", "BS_RR_J1", "BS_RR_S1", "BS_RR_S2"
                                line = get_line("BS", "BW_RR_kurz", ppr)
                            
                        Case "BS_BW_BW", "BS_BW_SH", "BS_F_BW_FO", "BS_F_RR_EF", "BS_F_RR_JF"
                            line = get_line("BS", "BW_BW_form", ppr)
                        
                        Case Else
                            MsgBox "Fehler bei der BS-Selektion"
                           
                    End Select
                    line = Replace(line, "x__st" & sei_1, re!Startnr)
                    line = Replace(line, "x__id" & sei_1, re!TP_ID)
                    line = Replace(line, "x__rh" & sei_1, rd)
                    If Not ppr Then
                        re.MoveNext
                        line = Replace(line, "x__st" & sei_2, re!Startnr)
                        line = Replace(line, "x__id" & sei_2, re!TP_ID)
                        line = Replace(line, "x__rh" & sei_2, rd)
                    End If
                    line = Replace(line, "x__title", Forms![A-Programmübersicht]!Turnierbez)
                    line = Replace(line, "x__vbs", "/cgi-bin/page.vbs")
                    line = Replace(line, "x__nPg", dNext & ".html")
                    line = Replace(line, "x__rnr", "Runde " & rd & " von " & rmax)
                    line = Replace(line, "x__wr", wr!Ausdr1)
                    line = Replace(line, "x__wid", wr!WR_ID)
                    line = Replace(line, "x__rd", rd_klasse)
                    line = Replace(line, "x__html", dNext & ".html")
                    line = Replace(line, "x__rt", RT_nr)
                    '----/BS Baden-Württemberg-------------------------------------------------------------------------------------
                Case "SLBS_"
                    '----BS Saarland-----------------------------------------------------------------------------------
                    line = get_line("BS", "SL_RR_kurz", ppr)
                        
                    line = Replace(line, "x__st" & sei_1, re!Startnr)
                    line = Replace(line, "x__id" & sei_1, re!TP_ID)
                    line = Replace(line, "x__rh" & sei_1, rd)
                    If Not ppr Then
                        re.MoveNext
                        line = Replace(line, "x__st" & sei_2, re!Startnr)
                        line = Replace(line, "x__id" & sei_2, re!TP_ID)
                        line = Replace(line, "x__rh" & sei_2, rd)
                    End If
                    line = Replace(line, "x__title", Forms![A-Programmübersicht]!Turnierbez)
                    line = Replace(line, "x__vbs", "/cgi-bin/page.vbs")
                    line = Replace(line, "x__nPg", dNext & ".html")
                    line = Replace(line, "x__rnr", "Runde " & rd & " von " & rmax)
                    line = Replace(line, "x__wr", wr!Ausdr1)
                    line = Replace(line, "x__wid", wr!WR_ID)
                    line = Replace(line, "x__rd", rd_klasse)
                    line = Replace(line, "x__html", dNext & ".html")
                    line = Replace(line, "x__rt", RT_nr)
                    '----/BS Saarland----------------------------------------------------------------------------------
                Case Else
                    MsgBox "Fehler bei der Selektion der Startklasse"
            End Select
            line = Replace(line, "x_title", Forms![A-Programmübersicht]!Turnierbez)
            line = Replace(line, "x_rnr", "Runde " & rd & " von " & rmax)
            line = Replace(line, "x_nPg", dNext & ".html")
            line = Replace(line, "x_wr", wr!Ausdr1)
            line = Replace(line, "x_wid", wr!WR_ID)
            line = Replace(line, "x_rd", rd_klasse)
            line = Replace(line, "x_html", dNext & ".html")
            line = Replace(line, "x_vbs", "/cgi-bin/page.vbs")
            line = Replace(line, "x_rt", RT_nr)
            out.writeline (line)
            out.Close
            re.MoveNext
            If re.EOF And fil = False Then
                pr.Filter = "nochmal = true"
                Set re = pr.OpenRecordset
                fil = True
                If Not re.EOF Then
                    re.MoveLast
                    re.MoveFirst
                End If
            End If
        Loop
        'HTML-Seite zu 1000
        Set out = file_handle(ht_pfad & dNext & ".html")
        line = "<!DOCTYPE html PUBLIC ""-//W3C//DTD HTML 4.01//EN"" >" & vbCrLf & "<html>" & "<head>" & _
               "<meta http-equiv=""refresh"" content=""0; URL=../" & _
               tr_nr & "R" & wr!WR_Lizenznr & "_K" & RT_nr & "_1000.html" & """>" & _
               "  <title></title>" & vbCrLf & "</head><body>" & vbCrLf & "</body></html>"
        out.writeline (line)
        'wird für platzieren benötigt
        write_index ht_pfad & tr_nr & "_index.html"
        
        ' HTML-Seite zur weiterleitung f1 ist warten  f2 an den Anfang
        Set out = file_handle(ht_pfad & tr_nr & "R" & wr!WR_Lizenznr & "_K" & RT_nr & "_1000.html")
        Set ht = open_re("All", "Warten")
        line = Replace(ht!F2, "x_title", Forms![A-Programmübersicht]!Turnierbez) ' kein warten
        line = Replace(line, "x_rd", rd_klasse)
        line = Replace(line, "x_wr", wr!Ausdr1)
        line = Replace(line, "x_html", tr_nr & "R" & wr!WR_Lizenznr & "_K" & RT_nr & "_1000.html")
        line = Replace(line, "x_pic", IpAddrs & "/" & tr_nr & "_K" & RT_nr & ".gif")
        out.writeline (line)
        out.Close
        wr.MoveNext
    Loop
    db.Execute ("UPDATE rundentab SET [HTML] = -1 WHERE RT_ID =" & RT_nr & ";")
    Start_Seite tr_nr
    
    pr.Filter = "Anwesend_Status=1"
    make_beobachter pr, rd_klasse, tr_nr, RT_nr
    
    
    'Schreibe all_cgi
    Set out = file_handle(vb_pfad & "all_cgi.vbs")
    Set ht = open_re("All", "Script")
    line = Replace(ht!F2, "x_pfad", getBaseDir)
    out.writeline (line)
    out.Close
    
    'Schreibe vbs zur Seite
    Set out = file_handle(vb_pfad & "page.vbs")
    Set ht = open_re("All", "Script")
    line = Replace(ht!f1, "x_url", IpAddrs)
    line = Replace(line, "x_pfad", vb_pfad)
    out.writeline (line)
    out.Close
    
    Set wr = Nothing
    Set ht = Nothing

End Sub

Function make_beobachter(pr, stkl_rde, tr_nr, RT_ID)
    Dim re, ht As Recordset
    Dim out
    Dim opt As String
    Dim line As String
    
    Set re = pr.OpenRecordset
    Set out = file_handle(getBaseDir & "Apache2\htdocs\" & tr_nr & "B99.html")
    Set ht = open_re("All", "Beobachter")
    line = ht!f1
    re.MoveFirst
    opt = Space(25) & "<option value=""0"">&nbsp;&nbsp;</option>" & vbCrLf
    Do Until re.EOF
        opt = opt & Space(25) & "<option value=""" & re!TP_ID & """>"
        opt = opt & Format(re!Startnr, "000") & "&nbsp;&nbsp;" & re!Da_Vorname & " " & re!Da_NAchname & " - " & re!He_Vorname & " " & re!He_Nachname & "</option>" & vbCrLf
        re.MoveNext
    Loop
    line = Replace(line, "x__option", opt)
    line = Replace(line, "x__title", Forms![A-Programmübersicht]!Turnierausw)
    line = Replace(line, "x__startklasse", stkl_rde)
    line = Replace(line, "x__npage", tr_nr & "B99.html")
    line = Replace(line, "x__rt_id", RT_ID)
    'Schreibe
    out.writeline (line)
    out.Close
End Function

Private Function write_index(fName)
    Dim out
    Dim line As String
    Set out = file_handle(fName)
    line = "<!DOCTYPE html PUBLIC ""-//W3C//DTD HTML 4.01//EN"" >" & vbCrLf & "<html>" & "<head>" & _
           "<meta http-equiv=""refresh"" content=""0; URL=index.html" & """>" & _
           "  <title></title>" & vbCrLf & "</head><body>" & vbCrLf & "</body></html>"
    out.writeline (line)

End Function

Private Function make_a_inpRRE(fName, max, lName, Optional st_kl) ' ak1x  , salto , 9
    Dim ak As String
    Dim t As Integer
    If InStr(1, fName, "ak") = 0 Then
        ak = tr & vbCrLf & Space(24) & "<td class=""akro"" colspan = ""19"">" & lName & "</td>" & vbCrLf
    Else
        If IsMissing(st_kl) Then
            ak = tr & vbCrLf & Space(24) & "<td class=""akro"" colspan = ""10"">" & lName & "</td>" & vbCrLf & make_akrofehler(fName)
        Else
            If left(st_kl, 3) = "F_R" Then
                ak = tr & vbCrLf & Space(24) & "<td class=""akro"" colspan = ""10"">" & lName & "</td>" & vbCrLf & make_akrofehler_Form(fName)
            Else
                ak = tr & vbCrLf & Space(24) & "<td style=""FONT-FAMILY: Arial; height:18; font-size:14px; padding-top:5px;"" colspan = ""21"">" & lName & "</td></tr><tr><td class=""akro"">" & vbCrLf & make_akrofehler(fName)
            End If
        End If
    End If
    ak = ak & Space(24) & "<td class=""akro"" colspan = ""2"" ><p id=""t" & fName & """>0%</p></td>" & vbCrLf & trn & vbCrLf & tr & vbCrLf
    For t = 0 To 20
        ak = ak & make_a_buttonRRE(fName, t, max) & "</td>" & vbCrLf
    Next
    ' ***** HM14.05 *****
    ' wegen Akrobatiken in Formationen
    If left(fName, 2) = "ak" Then
        ak = ak & vorspan & "<input class=""akrowert"" id=""w" & fName & """   type=""hidden"" name=""w" & fName & """></td>"
    Else
        ak = ak & vorspan & "<input id=""w" & fName & """   type=""hidden"" name=""w" & fName & """></td>"
    End If
    make_a_inpRRE = ak & vbCrLf & trn
End Function

Private Function make_a_inp(Name, max)   ' ak1x  , salto , 9
    Dim ak As String
    Dim t As Integer
    For t = 0 To (max * 2)
        ak = ak & make_a_button(Name, t, max) & "</td>" & vbCrLf
    Next
    make_a_inp = ak & vorspan & "<input id=""w" & Name & """   type=""hidden"" name=""w" & Name & """></td>"
End Function

Private Function make_akrofehler(fld)
    Dim akfl As String
    Dim n As String
    n = Mid(fld, 3, 1)
    akfl = Space(24) & "<td><input id=""fl" & n & "_" & fld & "_1""  type=""button"" value=""U2""   name=""fl" & n & "_" & fld & "_1""  class=""akfeh""  / LANGUAGE=javascript onclick=""return fehl('fl" & n & "_" & fld & "_1','U2')""   /></td>" & vbCrLf & _
           Space(24) & "<td><input id=""fl" & n & "_" & fld & "_2""  type=""button"" value=""10""  name=""fl" & n & "_" & fld & "_2""  class=""akfeh""  / LANGUAGE=javascript onclick=""return fehl('fl" & n & "_" & fld & "_2','U10')""  /></td>" & vbCrLf & _
           Space(24) & "<td><input id=""fl" & n & "_" & fld & "_3""  type=""button"" value=""20""  name=""fl" & n & "_" & fld & "_3""  class=""akfeh""  / LANGUAGE=javascript onclick=""return fehl('fl" & n & "_" & fld & "_3','U20')""  /></td>" & vbCrLf & _
           Space(24) & "<td><input id=""wfl" & n & "_" & fld & """   type=""hidden"" name=""wfl" & n & "_" & fld & """ value=""0""></td>" & vbCrLf & _
           Space(24) & "<td>&nbsp;</td>" & vbCrLf & _
           Space(24) & "<td><input id=""fl" & n & "_" & fld & "_5""  type=""button"" value=""S""  class=""akfeh""  ></td>" & vbCrLf & _
           Space(24) & "<td><input id=""fl" & n & "_" & fld & "_7""  type=""button"" value=""20""  name=""fl" & n & "_" & fld & "_7""  class=""akfeh""  / LANGUAGE=javascript onclick=""return fehl('fl" & n & "_" & fld & "_7','S20')""  /></td>" & vbCrLf & _
           Space(24) & "<td><input id=""tfl" & n & "_" & fld & """   type=""hidden"" name=""tfl" & n & "_" & fld & """ value=""""></td>" & vbCrLf & _
           Space(24) & "<td>&nbsp;</td>" & vbCrLf & _
           Space(24) & "<td><input id=""fl" & n & "_" & fld & "_9""  type=""button"" value=""V5""   name""fl" & n & "_" & fld & "_9""   class=""akfeh""  / LANGUAGE=javascript onclick=""return fehl('fl" & n & "_" & fld & "_9','V5')""   /></td>"
'           Space(24) & "<td><input id=""fl" & N & "_" & fld & "_6""  type=""button"" value=""10""  name=""fl" & N & "_" & fld & "_6""  class=""akfeh""  / LANGUAGE=javascript onclick=""return fehl('fl" & N & "_" & fld & "_6','S10')""  /></td>" & vbCrLf & _
           '"tak11_fl2"
    make_akrofehler = akfl
End Function

Private Function make_akrofehler_Form(fld)
    Dim akfl As String
    Dim n As String
    n = Mid(fld, 3, 1)
    akfl = Space(24) & "<td><input id=""fl" & n & "_" & fld & "_1""  type=""button"" value=""U2""   name=""fl" & n & "_" & fld & "_1""  class=""akfeh""  / LANGUAGE=javascript onclick=""return fehl('fl" & n & "_" & fld & "_1','U2')""   /></td>" & vbCrLf & _
           Space(24) & "<td><input id=""fl" & n & "_" & fld & "_2""  type=""button"" value=""10""  name=""fl" & n & "_" & fld & "_2""  class=""akfeh""  / LANGUAGE=javascript onclick=""return fehl('fl" & n & "_" & fld & "_2','U10')""  /></td>" & vbCrLf & _
           Space(24) & "<td><input id=""fl" & n & "_" & fld & "_3""  type=""button"" value=""20""  name=""fl" & n & "_" & fld & "_3""  class=""akfeh""  / LANGUAGE=javascript onclick=""return fehl('fl" & n & "_" & fld & "_3','U20')""  /></td>" & vbCrLf & _
           Space(24) & "<td><input id=""wfl" & n & "_" & fld & """   type=""hidden"" name=""wfl" & n & "_" & fld & """ value=""0""></td>" & vbCrLf & _
           Space(24) & "<td><input id=""fl" & n & "_" & fld & "_5""  type=""button"" style=""visibility:hidden;"" value=""S2""   name=""fl" & n & "_" & fld & "_5""  class=""akfeh""  / LANGUAGE=javascript onclick=""return fehl('fl" & n & "_" & fld & "_5','S2')""   /></td>" & vbCrLf & _
           Space(24) & "<td><input id=""fl" & n & "_" & fld & "_6""  type=""button"" style=""visibility:hidden;"" value=""10""  name=""fl" & n & "_" & fld & "_6""  class=""akfeh""  / LANGUAGE=javascript onclick=""return fehl('fl" & n & "_" & fld & "_6','S10')""  /></td>" & vbCrLf & _
           Space(24) & "<td><input id=""fl" & n & "_" & fld & "_7""  type=""button"" value=""S20""  name=""fl" & n & "_" & fld & "_7""  class=""akfeh""  / LANGUAGE=javascript onclick=""return fehl('fl" & n & "_" & fld & "_7','S20')""  /></td>" & vbCrLf & _
           Space(24) & "<td><input id=""tfl" & n & "_" & fld & """   type=""hidden"" name=""tfl" & n & "_" & fld & """ value=""""></td>" & vbCrLf & _
           Space(24) & "<td><input id=""fl" & n & "_" & fld & "_9""  type=""button"" value=""V5""   name""fl" & n & "_" & fld & "_9""   class=""akfeh""  / LANGUAGE=javascript onclick=""return fehl('fl" & n & "_" & fld & "_9','V5')""   /></td>"
           '"wak11_fl2"
           '"tak11_fl2"
    make_akrofehler_Form = akfl
End Function

Private Function make_a_buttonRRE(ID, wert, max)
    make_a_buttonRRE = vorspan & "<input id=""" & ID & "_" & wert & """  type=""button"" value=""" & wert * 5 & """ name=""" & ID & "_" & wert & """   class=""rr_leer""  onclick=""return wr_onclick('" & ID & "','" & wert & "','" & Replace(max, ",", ".") & "')"" >"
End Function

Private Function make_a_button(ID, wert, max)
    make_a_button = vorspan & "<input id=""" & ID & "_" & wert & """  type=""button"" value=""" & IIf(wert Mod 2, "-", wert / 2) & """ name=""" & ID & "_" & wert & """   class=""wr_leer""  / LANGUAGE=javascript onclick=""return wr_onclick('" & ID & "','" & wert & "','" & max * 2 & "')"" />"
End Function


Private Function vorspan()
    vorspan = Space(24) & "<td width=""15"" height=""30"" align=""center"" style=""white-space:nowrap"">"
End Function

Private Function make_akroRRE(re, Runde, seite, a_check, max, anz_p)
    Dim db As Database
    Dim ak As Recordset
    Dim out As String
    Dim t As Integer
    Dim pkt As Single
    Dim rde As String
    Set db = CurrentDb
    rde = ch_runde(Runde)
    out = kl_abstand
    If Runde = "End_r_Fuß" Then ' Bei A/B gibt es keine Akro in der FT-Runde
        out = "<td height=""200""></td>"
    Else
        For t = 1 To max
            Set ak = db.OpenRecordset("SELECT * FROM Akrobatiken WHERE Akrobatik='" & re("akro" & t & "_" & rde) & "' AND " & re!Startkl & "<> """";")
            If ak.RecordCount > 0 Then
                pkt = Nz(ak(re!Startkl))
                If pkt <> 0 Then
                    If anz_p = 1 Then
                        out = out & make_a_inpRRE("ak" & seite & t, pkt, ak!langtext, re!Startkl) & vbCrLf
                    Else
                        out = out & make_a_inpRRE("ak" & seite & t, pkt, ak!langtext, re!Startkl) & vbCrLf
                    End If
                    a_check = a_check & Chr(34) & "wak" & seite & t & Chr(34) & ", "
                Else
                    out = out & tr & vbCrLf & Space(24) & "<td></td>" & vbCrLf & trn & vbCrLf
                    out = out & tr & vbCrLf & Space(24) & "<td width=""15"" height=""15"" ></td>" & vbCrLf & trn & vbCrLf
                End If
            End If
        Next
    End If
    If a_check = "" Then a_check = Chr(34) & Chr(34) & ", "
    make_akroRRE = out & kl_abstand
End Function

Private Function make_akroOBS(re, Runde, seite, a_check, max, anz_p)
    Dim db As Database
    Dim ak As Recordset
    Dim out As String
    Dim t As Integer
    Dim pkt As Single
    Dim rde As String
    Set db = CurrentDb
    rde = ch_runde(Runde)
    out = kl_abstand
    If Runde = "End_r_Fuß" Then ' Bei A/B gibt es keine Akro in der FT-Runde
        out = "<td height=""200""></td>"
    Else
        For t = 1 To max
            Set ak = db.OpenRecordset("SELECT * FROM Akrobatiken WHERE Akrobatik='" & re("akro" & t & "_" & rde) & "' AND " & re!Startkl & "<> """";")
            If ak.RecordCount > 0 Then
                pkt = Nz(ak(re!Startkl))
                If pkt <> 0 Then
                    out = out & tr & vbCrLf & Space(24) & "<td class=""akro"" style=""padding-top:26px;"" colspan = ""13"">" & ak!langtext & "</td>" & vbCrLf & trn & vbCrLf
                    out = out & tr & vbCrLf & make_akrofehler("ak" & seite & t) & vbCrLf & trn & vbCrLf
                End If
            End If
        Next
        out = out & "<tr style=""visibility: hidden;"">" & vbCrLf
        For t = 1 To 10
            out = out & "<td style=""width:25px; heigth:20px"" ><input type=""button""></td> "
        Next
        out = out & trn & vbCrLf
    End If
    If a_check = "" Then a_check = Chr(34) & Chr(34) & ", "
    make_akroOBS = out & kl_abstand
End Function

Private Function make_akro(re, Runde, seite, a_check, max, anz_p)
    Dim out As String
    Dim t As Integer
    Dim rde As String
    rde = ch_runde(Runde)
    out = ""
    If Runde = "Vor_r_Fuß" Or Runde = "End_r_Fuß" Then ' Bei A/B gibt es keine Akro in der FT-Runde
        out = "<td height=""200""></td>"
    Else
        For t = 1 To max
            If Nz(re("akro" & t & "_" & rde)) <> "" Then               '0 Then
                If anz_p = 1 Then
                    out = out & tr & vbCrLf & Space(24) & "<td class=""akro"" width=30%>" & re("akro" & t & "_" & rde) & "</td>" & vbCrLf
                    out = out & vorspan & "</td>" & vbCrLf
                    out = out & make_a_inp("ak" & seite & t, Nz(re("wert" & t & "_" & rde))) & vbCrLf
                    out = out & Space(24) & "<td height=""40""></td>" & vbCrLf   ' nach height  width=30% entf
                Else
                    out = out & tr & vbCrLf & Space(24) & "<td class=""akro"" colspan = ""23"">" & re("akro" & t & "_" & rde) & "</td>" & vbCrLf
                    out = out & trn & vbCrLf & tr & vbCrLf
                    out = out & vorspan & "</td>" & vbCrLf
                    out = out & make_a_inp("ak" & seite & t, Nz(re("wert" & t & "_" & rde))) & vbCrLf
                End If
                a_check = a_check & """wak" & seite & t & """, "
                out = out & trn & vbCrLf
            Else
                out = out & tr & vbCrLf & Space(24) & "<td></td>" & vbCrLf & trn & vbCrLf
                out = out & tr & vbCrLf & Space(24) & "<td width=""15"" height=""15"" ></td>" & vbCrLf & trn & vbCrLf
            End If
        Next
    End If
    make_akro = out
    If a_check <> "" Then a_check = ", " & Mid(a_check, 1, Len(a_check) - 2)
End Function

Function get_line(cl, ber, ppr)
    Dim db As Database
    Dim ht As Recordset
    Set db = CurrentDb
    
    Set ht = db.OpenRecordset("SELECT * FROM HTML_Block WHERE Seite='" & cl & "' AND Bereich='" & ber & "';", DB_OPEN_DYNASET)
    If ppr Then
        get_line = ht!f1
    Else
        get_line = ht!F2
    End If
    Set ht = Nothing
    Set db = Nothing
End Function

Function kl_abstand()
    kl_abstand = tr & vbCrLf & Space(24) & "<td height=""4""></td>" & vbCrLf & trn & vbCrLf
End Function

Function akro_ber(s_kl, rde)        ' für Berechnung in HTML-Seite

    Select Case s_kl
        Case "RR_A"
            Select Case ch_runde(rde)
                Case "VR"
                    akro_ber = "an * 4"
                Case "ZR"
                    akro_ber = "an * 4"
                Case "ER"
                    akro_ber = "6 * 8"
                Case Else
                    akro_ber = 1
            End Select
                
        Case "RR_B"
            Select Case ch_runde(rde)
                Case "VR"
                    akro_ber = "an * 4"
                Case "ZR"
                    akro_ber = "an * 4"
                Case "ER"
                    akro_ber = "f_Iif(an < 6, 5, 6) * 8"
                Case Else
                    akro_ber = 1
            End Select
            
        Case "RR_C"
            If ch_runde(rde) = "ER" Then
                akro_ber = "4 * 4"
            Else
                akro_ber = "an * 4"
            End If
            
        Case "RR_J"
            If ch_runde(rde) = "ER" Then
                akro_ber = "3 * 4"
            Else
                akro_ber = "an * 4"
            End If
            
        Case "RR_S"
            akro_ber = " 1"
            
        Case "F_RR_ST", "F_RR_GF", "F_RR_LF", "F_RR_J"
            akro_ber = "an"
                            
        Case "F_RR_M"
            akro_ber = "f_Iif(an < 7, 6, an) "
        
        Case Else
            MsgBox "Fehler bei der Akroteiler Zuteilung"
            
    End Select

End Function

Public Function gen_default(pfad)   'erstellt bei Start des WebServers eine ganz einfache Startseite, falls noch keine existiert
    Dim out
    Dim nFile As String
    Dim ordner
    
    nFile = pfad & "index.html"
    ordner = Split(nFile, "\")
    If Dir(nFile) <> "index.html" Then
        Set out = file_handle(nFile)
        out.writeline ("<!DOCTYPE html><html><head><title>" & Forms![A-Programmübersicht]!Turnierbez & "</title><meta http-equiv=""expires"" content=""0""></head><body><p style=""font-size:50pt;""  align=""center""><br><br>" & Forms![A-Programmübersicht]!Turnierbez & "<br>" & ordner(4) & "</p></body></html>")
        out.Close
    End If
End Function

Function file_handle(fName)
    Dim fs As Variant
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set file_handle = fs.CreateTextFile(fName, True)
End Function

Public Sub Start_Seite(tr_nr)         'Alle WR und ihre Einteilungen
    Dim out
    Dim line As String
    Dim db As Database
    Dim wr As Recordset
    Dim ht As Recordset
    Dim re As Recordset
    Dim ht_pfad As String
    Dim i As Integer
    Dim einsatz As Boolean
    Dim kw, loc, nwr As String
    Dim twr As String
'    Dim bColor As String
    
    Set db = CurrentDb
    Set wr = db.OpenRecordset("SELECT Wert_Richter.*, [WR_Nachname] & "" "" & [WR_Vorname] AS Ausdr1 FROM Wert_Richter WHERE (((Wert_Richter.Turniernr)=" & get_aktTNr & ")) ORDER BY WR_Kuerzel;")
    ht_pfad = getBaseDir & "Apache2\htdocs\"
    
    ' Hier wird die Login-Seite für die WR erstellt
    Set out = file_handle(ht_pfad & "index.html")
    wr.MoveFirst
    Set ht = open_re("All", "WR_Select")
    line = Replace(ht!f1, "x__title", Forms![A-Programmübersicht]!Turnierbez)
    
    i = 1
    kw = "'0'"
    loc = "'0'"
    nwr = "'0'"
    Do Until wr.EOF
        kw = kw & ", '" & IIf(Nz(wr!WR_kenn) = "", wr!WR_Lizenznr, wr!WR_kenn) & "' "
        loc = loc & ", '" & tr_nr & "S" & wr!WR_Lizenznr & "' "
        nwr = nwr & ", '" & wr!Ausdr1 & "' "
        twr = twr & tr & vbCrLf & Space(24) & "<td class=""wr_m"" width=""200"">" & wr!WR_Kuerzel & "</td>" & vbCrLf
        twr = twr & Space(24) & "<td class=""wr_l"" width=""600""><a href=""javascript: weiter('" & i & "')"">" & wr!Ausdr1 & "</a></td>" & vbCrLf
        twr = twr & trn & vbCrLf
        wr.MoveNext
        i = i + 1
    Loop
    line = Replace(line, "x__wr", left(twr, Len(twr) - 1))
    line = Replace(line, "x__kw", kw)
    line = Replace(line, "x__loc", loc)
    line = Replace(line, "x__nwr", nwr)
    out.writeline (line)
    out.Close
    
    ' Ab hier werden die WR-Einteilungen erstellt
    Set ht = open_re("All", "WR_Start")
    wr.MoveFirst
    Set re = db.OpenRecordset("SELECT Startklasse.Startklasse, RT_ID, Startzeit, Rundentext & "" "" & Startklasse_text AS Ausdr1, HTML FROM Tanz_Runden INNER JOIN (Startklasse INNER JOIN Rundentab ON Startklasse.Startklasse = Rundentab.Startklasse) ON Tanz_Runden.Runde = Rundentab.Runde WHERE (Rundentab.Rundenreihenfolge<999 AND Rundentab.Turniernr=" & get_aktTNr & " AND Tanz_Runden.InAuswertung=1) OR (Rundentab.Rundenreihenfolge<999 AND Rundentab.Turniernr=" & get_aktTNr & " AND Right([Rundentab].[Runde],3)='Fuß') OR (Rundentab.Runde=""Sieger"") ORDER BY Rundentab.Rundenreihenfolge;")
    Do Until wr.EOF
        Set out = file_handle(ht_pfad & tr_nr & "S" & wr!WR_Lizenznr & ".html")
        re.MoveFirst
        line = Replace(ht!f1, "x_title", Forms![A-Programmübersicht]!Turnierbez)
        line = Replace(line, "wr_k", wr!WR_Kuerzel)
        line = Replace(line, "wr_n", wr!Ausdr1)
        out.writeline (line)
        Do Until re.EOF
            line = ht!F2
            line = Replace(line, "rt_z", Format(re!Startzeit, "hh:mm"))
            ' NewJudgingSystem Observer bekommt Verwarungen.
            If left(re!Startklasse, 3) = "BW_" Or left(re!Startklasse, 3) = "RR_" Then
                einsatz = (Nz(DLookup("[wr_id]", "Startklasse_Wertungsrichter", "Startklasse='" & re!Startklasse & "' AND WR_ID=" & wr!WR_ID)) > 0)
            Else
                einsatz = (Nz(DLookup("[wr_id]", "Startklasse_Wertungsrichter", "Startklasse='" & re!Startklasse & "' AND WR_ID=" & wr!WR_ID & " AND WR_function<>'Ob'")) > 0)
            End If
            If wr!WR_AzuBi Or einsatz Then
                'ist in Einsatz
                If re!HTML = True Then
                    line = Replace(line, "rt_c", "rd_v")
                    line = Replace(line, "rt_1f", tr_nr & "R" & wr!WR_Lizenznr & "_K" & re!RT_ID & "_1.html")
                Else
                    'ist erstellt
                    line = Replace(line, "rt_c", "rd_n")
                    line = Replace(line, "rt_1f", tr_nr & "S" & wr!WR_Lizenznr & ".html")
                End If
            Else
                'kein Einsatz
                line = Replace(line, "rt_c", "rd_k")
                line = Replace(line, "<a href=""rt_1f"">", "")
            End If
            line = Replace(line, "rt_1f", "R" & wr!WR_Lizenznr & "_K" & re!RT_ID & "_1.html")
            line = Replace(line, "rt_stk", re!Ausdr1)
            
            re.MoveNext
            out.writeline (line)
        Loop
        out.writeline ("    </table>" & vbCrLf & "   </form>" & vbCrLf & "</body>" & vbCrLf & "</html>")
        out.Close
        wr.MoveNext
    Loop
End Sub

Public Function GetIpAddrTable()  'IP-Adresse aus Regestry
    Dim Buf(0 To 511) As Byte
    Dim BufSize As Long
    Dim rc As Long
    Dim j As Integer, s As String
    Dim i As Integer
   
    BufSize = UBound(Buf) + 1
    rc = GetIpAddrTable_API(Buf(0), BufSize, 1)
    If rc <> 0 Then err.Raise vbObjectError, , "GetIpAddrTable lieferte kein Ergebnis." & rc
    Dim NrOfEntries As Integer
    NrOfEntries = Buf(1) * 256 + Buf(0)
    Select Case NrOfEntries
        Case 0, 1
            GetIpAddrTable = "keine Netzwerkschnittstelle"
            If get_properties("NetzwerkCheck") = "ein" Then MsgBox "Es ist keine Netzwerkschnittstelle aktiv"
        Case 2
            For i = 0 To NrOfEntries - 1
                s = ""
                For j = 0 To 3
                    s = s & IIf(j > 0, ".", "") & Buf(4 + i * 24 + j)
                Next
                If s <> "127.0.0.1" Then
                    GetIpAddrTable = s
                    Exit Function
                End If
            Next
        Case Is > 2
            For i = 0 To NrOfEntries - 1
                s = ""
                For j = 0 To 3
                    s = s & IIf(j > 0, ".", "") & Buf(4 + i * 24 + j)
                Next
                If s <> "127.0.0.1" Then
                    GetIpAddrTable = s
                    Exit For
                End If
            Next
            If get_properties("NetzwerkCheck") = "ein" Then MsgBox "Es sind mehrere Netzwerkschnittstellen aktiv! Fehlfunktion möglich!" & vbCrLf & "Deaktivieren Sie alle bis auf eine Schnittstelle!", , "Turnierprogramm"
    End Select
End Function

Function no_plazieren(RT_ID, WR_ID, mehrfach, pg_id, s_kl)
    Dim db As Database
    Dim re As Recordset
    Dim ht_pfad As String
    Dim out
    
'    Set db = CurrentDb
'    Set re = db.OpenRecordset("Select * from Wert_richter where wr_id = " & WR_ID & ";")
'    ht_pfad = getBaseDir
'
'    Set out = file_handle(ht_pfad & "T" & Forms![A-Programmübersicht]!Turnier_Nummer & "R" & re!WR_Lizenznr & "_K" & RT_ID & "_2000.html")
'    Set re = db.OpenRecordset("SELECT * FROM HTML_Block WHERE Seite=""BW"" AND Bereich=""Platzierung"";", DB_OPEN_DYNASET)
'    html_page = re!f2
'    out.writeline re!f2
'    out.Close


End Function

Public Sub pg_platzieren(RT_ID, WR_ID, mehrfach, pg_id, s_kl)
    Dim db As Database  ' Hier werden die Platzierungsseiten erstellt
    Dim re As Recordset
    Dim ht_t As Recordset
    Dim vars
    Dim out
    Dim fld As Variant
    Dim t, i As Integer
    Dim ht_pfad As String
    Dim html_page As String
    Dim wr_liz As String
    a_check = ""
        
    Set db = CurrentDb
    Set re = db.OpenRecordset("Select * from Wert_richter where wr_id = " & WR_ID & ";")
    wr_liz = re!WR_Lizenznr
    ht_pfad = getBaseDir & "Apache2\htdocs\"
    
    Set out = file_handle(ht_pfad & "T" & Forms![A-Programmübersicht]!Turnier_Nummer & "R" & wr_liz & "_K" & RT_ID & "_2000.html")
    Set re = db.OpenRecordset("SELECT Paare_Rundenqualifikation.RT_ID, Paare_Rundenqualifikation.Anwesend_Status, Auswertung.PR_ID, Auswertung.WR_ID, Paare_Rundenqualifikation.Rundennummer, Auswertung.Cgi_Input, Paare.Startnr, Auswertung.Punkte, Auswertung.Platz, [WR_Nachname] & "" "" & [WR_Vorname] AS Ausdr1, Startklasse.Startklasse_text, IIf(NZ([Name_Team])="""",[Da_Nachname] & ""<br>"" & [He_Nachname],[Name_Team]) AS Ausdr2, Paare_Rundenqualifikation.TP_ID FROM Startklasse INNER JOIN (Wert_Richter INNER JOIN (Paare INNER JOIN (Paare_Rundenqualifikation INNER JOIN Auswertung ON Paare_Rundenqualifikation.PR_ID = Auswertung.PR_ID) ON Paare.TP_ID = Paare_Rundenqualifikation.TP_ID) ON Wert_Richter.WR_ID = Auswertung.WR_ID) ON Startklasse.Startklasse = Paare.Startkl WHERE ((Paare_Rundenqualifikation.RT_ID=" & RT_ID & ") AND (Anwesend_Status=1) AND (Auswertung.WR_ID=" & WR_ID & ")) ORDER BY Auswertung.WR_ID, Rundennummer, Startnr;")
    re.MoveFirst
    
    Select Case get_bs_erg(s_kl, 3)
        Case "BS_"
            Set ht_t = open_re("BS", "Platzierung")
            fld = Array("x_wsh")
        
        Case "NBS_"
            Select Case left(s_kl, 6)
                Case "BS_RR_"
                    Set ht_t = open_re("RR", "Platzierung")
                    fld = Array("x_wsd", "x_wth", "x_wak")
                Case "BS_F_R"
                    Set ht_t = open_re("RR", "FormPlatzierung")
                    fld = Array("x_wth", "x_wak")
                Case "BS_"
                    Set ht_t = open_re("BS", "Platzierung")
                    fld = Array("x_wsh")
                Case Else
                    MsgBox "Fehler in der Breitensportplatzierung"
            End Select
        
        Case "BW_"
            Set ht_t = open_re("BW", "Platzierung")
            fld = Array("x_wsd", "x_wtd", "x_wfg")
            
        Case "F_B"
            Set ht_t = open_re("BW", "FormPlatzierung")
            fld = Array("x_wth", "x_wch", "x_wfg")
            
        Case "LH_"
            Set ht_t = open_re("LH", "Platzierung")
            fld = Array("x_wsd", "x_wtd", "x_wfg")
        
        Case "RR_"
            Set ht_t = open_re("RR", "Platzierung")
            fld = Array("x_wsd", "x_wth", "x_wak")

        Case "F_R", "F_F"
            Set ht_t = open_re("RR", "FormPlatzierung")
            fld = Array("x_wth", "x_wak")
            
    End Select
    html_page = fill_platzieren(ht_t!f1, re, mehrfach, fld)
    html_page = Replace(html_page, "x_ck1", Mid(a_check, 3))
    re.MoveFirst
    html_page = Replace(html_page, "x_title", Forms![A-Programmübersicht]!Turnierbez)
    html_page = Replace(html_page, "x_wrname", re!Ausdr1)         ' Wertungsrichter Name
    html_page = Replace(html_page, "x_wrid", re!WR_ID)            ' Wertungsrichter Index
    html_page = Replace(html_page, "x_html", get_TerNr & "_index.html")         ' nächste Seite von weiter
    html_page = Replace(html_page, "x_nPg", get_TerNr & "_index.html")          ' nächste Seite für Senden
    html_page = Replace(html_page, "x_rtname", "Plazierung für " & re!Startklasse_text)
    html_page = Replace(html_page, "x_rtid", re!RT_ID * 1000)    ' ID für Dateinamen
    
    out.writeline html_page
    out.Close

End Sub

Public Function open_re(p1, p2)
    Dim db As Database
    Set db = CurrentDb
    Set open_re = db.OpenRecordset("SELECT * FROM HTML_Block WHERE Seite=""" & p1 & """ AND Bereich=""" & p2 & """;", DB_OPEN_DYNASET)

End Function

Private Function fill_platzieren(html_page As String, re As Recordset, mehrfach As Variant, fld As Variant)
    Dim i, t As Integer
    Dim vars
    For t = 1 To 7
        If re.EOF Then
            For i = 0 To UBound(fld)
                 html_page = Replace(html_page, fld(i) & t, "&nbsp;")
            Next
            html_page = Replace(html_page, "x_wak" & t, "&nbsp;")
            html_page = Replace(html_page, "x_stn" & t, "&nbsp;")
            html_page = Replace(html_page, "x_pr" & t, 0)
            html_page = Replace(html_page, "x_wsh" & t, "&nbsp;")
            html_page = Replace(html_page, "x_wch" & t, "&nbsp;")
            html_page = Replace(html_page, "x_wfh" & t, "&nbsp;")
            html_page = Replace(html_page, "x_wft" & t, "&nbsp;")
            html_page = Replace(html_page, "x_wpt" & t, "&nbsp;")
            html_page = Replace(html_page, "x_inp" & t, "&nbsp;")
            html_page = Replace(html_page, "x_nam" & t, "&nbsp;")
        Else
            Set vars = zerlege(Nz(re!Cgi_Input))
            For i = 0 To UBound(fld)
                If fld(i) = "x_wak" Then
                    html_page = Replace(html_page, "x_wak" & t, Replace(vars.Item("wak11") & "/" & vars.Item("wak12") & "/" & vars.Item("wak13") & "/" & vars.Item("wak14") & "/" & vars.Item("wak15") & "/" & vars.Item("wak16") & "/" & vars.Item("wak17") & "/" & vars.Item("wak18"), "//", ""))
                Else
                    html_page = Replace(html_page, fld(i) & t, vars.Item(Mid(fld(i), 3) & "1"))
                End If
            Next
            html_page = Replace(html_page, "x_stn" & t, re!Startnr)
            html_page = Replace(html_page, "x_pr" & t, re!TP_ID)
            html_page = Replace(html_page, "x_wsh" & t, vars.Item("wsh1"))
            html_page = Replace(html_page, "x_wch" & t, vars.Item("wch1"))
            html_page = Replace(html_page, "x_wfh" & t, vars.Item("wfl1"))
            html_page = Replace(html_page, "x_wft" & t, vars.Item("fl1"))
            html_page = Replace(html_page, "x_wpt" & t, re!Punkte)
            html_page = Replace(html_page, "x_inp" & t, make_option(re, mehrfach, t))
            html_page = Replace(html_page, "x_nam" & t, re!Ausdr2)
            re.MoveNext
        End If
    Next
    fill_platzieren = html_page

End Function

Private Function make_option(re, mehrfach, nr)  'Hier wird ein Optionsfeld für Platzierung erstellt
    Dim i As Integer
    Dim opt As String
    If Nz(mehrfach(re!Platz)) = 0 Then
        make_option = Space(18) & "<input class=""pl_se"" name=""wpt" & nr & """ value=""" & re!Platz & """ disabled>"
    Else
        opt = Space(18) & "<select name=""wpt" & nr & """ id=""wpt" & nr & """ class=""pl_se"" / LANGUAGE=javascript onchange=""return wr_onclick('wpt" & nr & "')"" />"
        opt = opt & "<option selected></option>"
        For i = 0 To mehrfach(re!Platz) - 1                 'Sooft wie gleicher Platz vorhanden ist
            opt = opt & "<option value=""" & re!Platz + i & """>" & re!Platz + i & "</option>"
        Next
        make_option = opt & "</select>"
        a_check = a_check & ", ""wpt" & nr & """"
    End If
End Function

Public Function Umlaute_Umwandeln(XML_String As Variant)
    
    Dim XML_STRING_Neu As String
    Dim i As Integer
    Dim validChars As String
    
    validChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZäüöÄÜÖß-_.,1234567890 á'"
    
    If IsNull(XML_String) Then
        Umlaute_Umwandeln = ""
        Exit Function
    Else
        
    End If
    
    XML_STRING_Neu = CStr(XML_String)
    
    '*** Sonstige Sonderzeichen ersetzen
    For i = 1 To Len(XML_String)
    If InStr(1, validChars, Mid$(XML_String, i, 1)) = 0 Then
        XML_STRING_Neu = Replace(XML_STRING_Neu, Mid$(XML_String, i, 1), " ", , , vbBinaryCompare)
    End If
    Next i
    
    
    XML_STRING_Neu = Replace(XML_STRING_Neu, "Ä", "&#196;", , , vbBinaryCompare)
    XML_STRING_Neu = Replace(XML_STRING_Neu, "Ö", "&#214;", , , vbBinaryCompare)
    XML_STRING_Neu = Replace(XML_STRING_Neu, "Ü", "&#220;", , , vbBinaryCompare)
    XML_STRING_Neu = Replace(XML_STRING_Neu, "ä", "&#228;", , , vbBinaryCompare)
    XML_STRING_Neu = Replace(XML_STRING_Neu, "ö", "&#246;", , , vbBinaryCompare)
    XML_STRING_Neu = Replace(XML_STRING_Neu, "ü", "&#252;", , , vbBinaryCompare)
    XML_STRING_Neu = Replace(XML_STRING_Neu, "ß", "&#223;", , , vbBinaryCompare)
    XML_STRING_Neu = Replace(XML_STRING_Neu, "á", "&#225;", , , vbBinaryCompare)
    XML_STRING_Neu = Replace(XML_STRING_Neu, "é", "&#233;", , , vbBinaryCompare)
    XML_STRING_Neu = Replace(XML_STRING_Neu, "'", "&#39;", , , vbBinaryCompare)
    
    Umlaute_Umwandeln = XML_STRING_Neu
End Function

' ***** HM 14.05 *****
' KO-Runde hinzugefügt
Public Function ch_runde(rd)  'Hier wird ausgewählt welche Akrobatiken verwendet werden
    If InStr(1, rd, "_  Fuß") Then
        ch_runde = ""         ' Fußtechnikrunde keine Addition der Fußtechnik
    ElseIf InStr(1, rd, "Vor_") Then
        ch_runde = "VR"
    ElseIf InStr(1, rd, "Hoff_") Then
        ch_runde = "VR"
    ElseIf InStr(1, rd, "Zw_") Then
        ch_runde = "ZR"
    ElseIf InStr(1, rd, "KO_r") Then
        ch_runde = "ER"
    ElseIf rd = "Stich_r" Then
        ch_runde = "ZR"
    ElseIf rd = "Semi" Then
        ch_runde = "ER"
    ElseIf InStr(1, rd, "End_") Then
        ch_runde = "ER"
    ElseIf rd = "Stich_r_1pl" Then
        ch_runde = "ER"
    End If
End Function

Public Function get_bs_erg(st_kl, rueckgabe)
    If left(st_kl, 3) = "BS_" Then
        get_bs_erg = UCase(Nz(DLookup("BS_Erg", "Turnier", "Turniernum=" & get_aktTNr))) & left(st_kl, 3)
    Else
        get_bs_erg = left(st_kl, rueckgabe)
    End If
End Function

Public Function get_aktTNr()
    get_aktTNr = Forms![A-Programmübersicht]!Akt_Turnier
End Function

Public Function get_TerNr()
    get_TerNr = "T" & Forms![A-Programmübersicht]!Turnier_Nummer
End Function

Public Function make_wr_zeitplan()
    Dim db As Database
    Dim re As Recordset
    Dim wr_zeitplan As String

    Set db = CurrentDb
    Set re = db.OpenRecordset("SELECT RT_ID, Startzeit, Startklasse_text, Rundentext FROM Tanz_Runden INNER JOIN (Rundentab LEFT JOIN Startklasse ON Rundentab.Startklasse = Startklasse.Startklasse) ON Tanz_Runden.Runde = Rundentab.Runde WHERE (Rundentab.Rundenreihenfolge<1000) ORDER BY Rundentab.Rundenreihenfolge;")
    If re.RecordCount = 0 Then Exit Function
        Open getBaseDir & "webserver\views\Zeitplan.html" For Output As #1
        Print #1, "<!DOCTYPE html><head><title>beamer</title><meta http-equiv=""expires"" content=""0""><link rel=""stylesheet"" href=""EWS3.css"">"
        Print #1, "</head><body style=""height: 98%; font-family: Verdana;"" id=""beamer_seite"">"
        Print #1, "<table cellpadding=""0"" frame=""void"" class=""tb1""><tr height=""20%""><td><table width=""100%"">"
        Print #1, "<tr><td class=""kopf"" width=""300px""><img src=""Logo.jpg"" width=""290"" height=""180"" alt=""DRBV""></td>"
        Print #1, "<td class=""kopf"" width=""auto"" id=""beamer_kopf"">" & Forms![A-Programmübersicht]!Turnierausw & "</td ></tr></table></td></tr>"
        Print #1, "<tr height=""80%""><td><table style=""width: 100%; float: left; "" id=""beamer_inhalt""><tr><td>&nbsp;</td></tr>"
        
        Print #1, "<tr height=""100%""><td><table style=""width: 100%; float: left; "">"
        Print #1, "<thead><tr class=""runden"" role=""row""><th style=""width: 250px; padding-left:100px; "" colspan=""1"" rowspan=""1"" class=""sorting"">Beginn</th><th style=""width: auto;"" colspan=""1"" rowspan=""1"" class=""sorting"">Runde</th></tr></thead>"
        Print #1, "<tbody style=""font-size: 2.5vw;"">"
        
        re.MoveFirst
        Do Until re.EOF
            Print #1, "<tr class=""odd"" ><td style=""padding-left:100px;"" >" & Format(re!Startzeit, "hh:mm") & "</td><td>" & Umlaute_Umwandeln(re!Rundentext & " " & Nz(re!Startklasse_text)) & "</td></tr>"
            re.MoveNext
        Loop
        Print #1, "</tbody></table></td></tr></table></td></tr></table></body></html>"
        Close #1
End Function


