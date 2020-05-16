Option Compare Database
Option Explicit

Public Sub make_a_startlist(RT_ID As Integer)
    'Const RT_ID = 11
    
    If get_mod_on Then Exit Sub
    gen_Ordner getBaseDir & "Apache2\htdocs\moderator"

    Dim db As Database
    Dim ht, re As Recordset
    Dim st_kl, rde As String
    Dim out As Variant
    Dim ht_pfad As String
    Dim tr_nr As String
    Dim line As String
    
    Set db = CurrentDb
    Set re = db.OpenRecordset("SELECT Rundentab.*, Startklasse.Startklasse_text, Tanz_Runden.Rundentext FROM Tanz_Runden INNER JOIN (Rundentab INNER JOIN Startklasse ON Rundentab.Startklasse = Startklasse.Startklasse) ON Tanz_Runden.Runde = Rundentab.Runde WHERE (((Rundentab.Turniernr)=" & get_aktTNr & ") AND ((Rundentab.RT_ID)=" & RT_ID & "));")
    st_kl = re!Startklasse_text
    rde = re!Rundentext
    tr_nr = "T" & Forms![A-Programmübersicht]!Turnier_Nummer
    ht_pfad = getBaseDir & "Apache2\htdocs\"
    
    Set ht = open_re("Mod", "Start")
    Set out = file_handle(ht_pfad & "moderator\" & tr_nr & "_Sta_K" & RT_ID & ".html")
    
    line = Replace(ht!f1, "x_title", "Startliste  " & st_kl)
    line = Replace(line, "rt__txt", "Startliste  " & st_kl & " " & rde)
    
    Set re = db.OpenRecordset("SELECT Paare.Startnr, [Da_Vorname] & "" "" & [Da_Nachname] AS DName, [He_Vorname] & "" "" & [He_Nachname] AS HName, Paare.Name_Team, Paare.Verein_Name, IIf(Nz([Name_Team])="""",False,True) AS isTeam FROM Paare_Rundenqualifikation INNER JOIN Paare ON Paare_Rundenqualifikation.TP_ID = Paare.TP_ID WHERE (((Paare_Rundenqualifikation.RT_ID)=" & RT_ID & ")) ORDER BY Paare.Startnr;")
    re.MoveFirst
    Do Until re.EOF
         line = line & vbCrLf & "    <tr height=""40"">" & vbCrLf & _
                Space(8) & "<td width=6% align=""center"">" & re!Startnr & "</td>" & vbCrLf & _
                Space(8) & "<td width=47%>" & IIf(re!isTeam, re!Name_Team, re!dname & " - " & re!hName) & "</td>" & vbCrLf & _
                Space(8) & "<td width=47%>" & re!Verein_Name & "</td>" & vbCrLf & _
                "    </tr>"
        re.MoveNext
    Loop
    out.writeline (line)
    out.writeline ("    </table>" & vbCrLf & "   </form>" & vbCrLf & "</body>" & vbCrLf & "</html>")
    out.Close
    
    Set re = db.OpenRecordset("SELECT Rundentab.* FROM Rundentab WHERE (((Rundentab.Turniernr)=" & get_aktTNr & ") AND ((Rundentab.RT_ID)=" & RT_ID & "));")
    re.Edit
    re!RT_Stat = 1
    re.Update
    
    re.Close
    ht.Close
    db.Close
    make_a_schedule         'Zeiplan neu schreiben
    make_a_off              'Offiziellenliste schreiben
    make_a_wrlist           'WR-Einteilung schreiben
    make_a_Vorstellungslist 'Vorstellungslist nach Vereinen

End Sub

Public Sub make_a_Vorstellungslist()
' Moderator Vorstellung aller Paare nach Verein
    If get_mod_on Then Exit Sub
    gen_Ordner getBaseDir & "Apache2\htdocs\moderator"

    Dim db As Database
    Dim ht, re As Recordset
    Dim sText As String
    Dim out As Variant
    Dim ht_pfad As String
    Dim tr_nr As String
    Dim line As String
    Dim vText As String
    
    Set db = CurrentDb
    tr_nr = "T" & Forms![A-Programmübersicht]!Turnier_Nummer
    ht_pfad = getBaseDir & "Apache2\htdocs\"
    
    Set ht = open_re("Mod", "Start")
    
    Set out = file_handle(ht_pfad & "moderator\" & tr_nr & "_Vorstellung.html")
    
    Set re = db.OpenRecordset("SELECT Verein_Name, Paare.Startnr, [Da_Vorname] & "" "" & [Da_Nachname] AS DName, [He_Vorname] & "" "" & [He_Nachname] AS HName, Name_Team, IIf(Nz([Name_Team])="""",False,True) AS isTeam, Startklasse.Startklasse_text, Paare.Verein_nr FROM Startklasse INNER JOIN Paare ON Startklasse.Startklasse = Paare.Startkl WHERE ((Paare.Turniernr=" & get_aktTNr & ") AND (Paare.Anwesent_Status=1)) ORDER BY Paare.Verein_name, Startklasse.Startklasse_text, Paare.Startnr;")
    
    line = Replace(ht!f1, "x_title", "Vorstellung aller Tanzpaare")
    line = Replace(line, "rt__txt", re!Verein_Name)
    vText = Trim(re!Verein_Name)
    
    re.MoveFirst
    Do Until re.EOF
        line = line & vbCrLf & "    <tr height=""40"">" & vbCrLf & _
                Space(8) & "<td width=6% align=""center"">" & re!Startnr & "</td>" & vbCrLf & _
                Space(8) & "<td width=47%>" & IIf(re!isTeam, re!Name_Team, re!dname & " - " & re!hName) & "</td>" & vbCrLf & _
                Space(8) & "<td width=47%>" & re!Startklasse_text & "</td>" & vbCrLf & _
                "    </tr>"
        re.MoveNext
        If Not re.EOF Then
            If vText <> Trim(re!Verein_Name) Then
                line = line & vbCrLf & "    <tr>" & vbCrLf & _
                       "        <td colspan=""3"" class=""wr_l"" >" & re!Verein_Name & "</td>" & vbCrLf & _
                       "    </tr>"
                vText = Trim(re!Verein_Name)
            End If
        End If
    Loop
    out.writeline (line)
    out.writeline ("    </table>" & vbCrLf & "   </form>" & vbCrLf & "</body>" & vbCrLf & "</html>")
    out.Close
    
    re.Close
    ht.Close
    db.Close
End Sub

Public Sub make_a_round(pr As Recordset, st_klasse As String, Runde As String, runde_id)
' Moderator Rundeneinteilung
    If get_mod_on Then Exit Sub
    
    Dim db As Database
    Dim ht, re As Recordset
    Dim out As Variant
    Dim ht_pfad As String
    Dim tr_nr As String
    Dim line As String
    Dim Sta, cou, org As String
    Dim max_Runde As Integer
    Dim Runden_anz As Integer
    Dim fil As Boolean
    
    make_a_schedule
    make_a_off
    make_a_wrlist
    
    Set re = pr
    re.MoveFirst
    If IsNull(re!Rundennummer) Then
        MsgBox "Scheinbar wurde die Runde noch nicht ausgelost."
    Else
        re.MoveLast
    
        tr_nr = "T" & Forms![A-Programmübersicht]!Turnier_Nummer
        ht_pfad = getBaseDir & "Apache2\htdocs\"
        
        Set db = CurrentDb
        Set ht = open_re("Mod", "Runde")
        
        Set out = file_handle(ht_pfad & "moderator\" & tr_nr & "_Mod_K" & runde_id & ".html")
    
'        Call erste_Folie(oPPTPres.Slides(1), st_Klasse & " " & Runde)
        line = Replace(ht!f1, "x_title", "Rundeneinteilung" & " " & st_klasse)
        line = Replace(line, "rt__txt", st_klasse & " " & Runde)
'        out.Writeline (line)
        
        max_Runde = Int(re.RecordCount / re!Anz_Paare) + re.RecordCount Mod re!Anz_Paare
        re.MoveFirst
        Do Until re.EOF
            If Nz(re!Rundennummer) <> Runden_anz Then
                line = Replace(line, "rt__st", Sta)
                line = Replace(line, "rt__cou", cou)
                line = Replace(line, "rt__org", org)
                out.writeline (line)
                line = ht!F2
                Runden_anz = Runden_anz + 1
                line = Replace(line, "rt__nr", Runden_anz)
                line = Replace(line, "x__rhrd", "rhrd" & Runden_anz)
                Sta = ""
                cou = ""
                org = ""
            End If
            Sta = Sta & IIf(Len(Sta) = 0, "", "<br>") & re!Startnr
            If IsNull(re!Name_Team) Then    'Dame u Herr
                cou = cou & IIf(Len(cou) = 0, "", "<br>") & get_Dame(re) & " - " & get_Herr(re)
            Else                            'Team
                cou = cou & IIf(Len(cou) = 0, "", "<br>") & re!Name_Team
            End If
            If re!Anwesend_Status = 2 Then
                cou = cou & "<br><b>noch nicht Anwesend!</b>"
            End If
            org = org & IIf(Len(org) = 0, "", "<br>") & re!Verein_Name
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
        line = Replace(line, "rt__st", Sta)
        line = Replace(line, "rt__cou", cou)
        line = Replace(line, "rt__org", org)
        out.writeline (line)
        out.writeline ("    </table>" & vbCrLf & "   </form>" & vbCrLf & "</body>" & vbCrLf & "</html>")
        out.Close
    End If
End Sub

Public Sub make_a_siegerehrung(RT_ID As Integer)
    
    If get_mod_on Then Exit Sub
'Const RT_ID = 4
    Dim db As Database
    Dim ht, re, wr As Recordset
    Dim st_kl As String
    Dim out As Variant
    Dim ht_pfad As String
    Dim tr_nr As String
    Dim line As String
    Dim t As Integer
    Dim sie_id As String
    
    Set db = CurrentDb
    Set re = db.OpenRecordset("SELECT v.*, r.*,p.* FROM (View_Majoritaet v INNER JOIN View_Runden r ON v.RT_ID = r.RT_ID) INNER JOIN View_Paare p ON v.TP_ID = p.TP_ID WHERE (v.RT_ID=" & RT_ID & ") ORDER BY v.Platz DESC;")
    ' Wegen neuer Wertung im RR
    If re.RecordCount = 0 Then Exit Sub
    sie_id = get_siegerID(re![r.Startklasse])
    st_kl = re![r.Startklasse_text]
    tr_nr = "T" & Forms![A-Programmübersicht]!Turnier_Nummer
    ht_pfad = getBaseDir & "Apache2\htdocs\"
    
    Set ht = open_re("All", "Sieger")
    Set out = file_handle(ht_pfad & "moderator\" & tr_nr & "_Mod" & sie_id & ".html")
    Set wr = db.OpenRecordset("SELECT Startklasse, wr.* FROM Wert_Richter AS wr INNER JOIN Startklasse_Wertungsrichter ON wr.WR_ID = Startklasse_Wertungsrichter.WR_ID WHERE (Startklasse='" & re![r.Startklasse] & "' AND Startklasse_Wertungsrichter.WR_function<>'Ob') ORDER BY wr.WR_Kuerzel;")
    
    line = Replace(ht!f1, "x_title", "Siegerehrung  " & re![r.Startklasse_text])
    line = Replace(line, "rt__txt", "Siegerehrung  " & re![r.Startklasse_text] & "  Moderator")
    line = get_sieger(re, "MajoritaetKurz", line)
    out.writeline (line)
    out.writeline ("    </table>" & vbCrLf & "   </form>" & vbCrLf & "</body>" & vbCrLf & "</html>")
    re.MoveFirst
    
    If left(re![r.Startklasse], 3) <> "RR_" Then
        wr.MoveFirst
        t = 1
        Do Until wr.EOF
            re.MoveFirst
            Set out = file_handle(ht_pfad & "" & tr_nr & "R" & wr!WR_Lizenznr & sie_id & "_1.html")
            line = Replace(ht!F2, "x_title", "Siegerehrung  " & re![r.Startklasse_text])
            line = Replace(line, "rt__txt", "Siegerehrung  " & re![r.Startklasse_text] & "  " & wr!WR_Vorname & " " & wr!WR_Nachname)
            line = Replace(line, "rt__loc", tr_nr)
            line = get_sieger(re, "WR" & t, line)
            
            out.writeline (line)
            out.writeline ("    </table>" & vbCrLf & "   </form>" & vbCrLf & "</body>" & vbCrLf & "</html>")
            t = t + 1
            wr.MoveNext
        Loop
        out.Close
    End If
    make_a_schedule
    Start_Seite tr_nr
End Sub

Function get_siegerID(st_kl)
    Dim rt As Recordset
    Set rt = DBEngine(0)(0).OpenRecordset("SELECT * FROM rundentab WHERE Startklasse ='" & st_kl & "' and Runde ='Sieger';")
    If rt.RecordCount = 0 Then
        'MsgBox "Es wurde zu dieser Startklasse keine Siegerehrung gefunden!"
    Else
        get_siegerID = "_K" & rt!RT_ID
        rt.Edit
        rt!HTML = True
'        rt!RT_Stat = 1
        rt.Update
    End If
End Function

Function get_sieger(st, Feld, line)
    Dim t As Integer
    line = line & vbCrLf & "    <tr class=""sie_t"" >" & vbCrLf & _
           Space(8) & "<td width=8%>Platz</td>" & vbCrLf & _
           Space(8) & "<td width=8%>Startnr</td>" & vbCrLf & _
           Space(8) & "<td width=55% align=""Left"">Name und Verein</td>" & vbCrLf & _
           Space(8) & "<td width=29%>Eigene Wertung</td>" & vbCrLf & _
           "    </tr>"
           
    st.MoveFirst
    t = 1
    Do Until st.EOF
        line = line & vbCrLf & "    <tr id=""rhrd" & t & """ class=""sie_a"" >" & vbCrLf & _
               Space(8) & "<td width=8%><font size=""5"">" & st![v.Platz] & "</font></td>" & vbCrLf & _
               Space(8) & "<td width=8%>" & st!Startnr & "</td>" & vbCrLf & _
               Space(8) & "<td width=55% align=""Left"">" & st![p.Name] & "<br>" & st!Verein_Name & _
               IIf(st!DQ_ID > 0, "<br><span style=""background-color: #FF1010; color:#FFFFFF"">&nbsp;Disqualifiziert &nbsp;</span>", "") & _
               IIf(st!pa_grund <> "Alles OK", "<br><span style=""background-color: #FF1010; color:#FFFFFF"">&nbsp;Regelversto&szlig; &nbsp;</span>", "") & _
               "</td>" & vbCrLf & Space(8) & "<td width=29%><font size=""6""><a href=""javascript:cha_co('" & t & "','" & st(Feld) & "')"">" & IIf(left(st!Startkl, 3) = "RR_", "<img src=""../ball.red.png"" border=""0""  alt=""done"">", st(Feld)) & "</font></td>" & vbCrLf & _
               "    </tr>"
        st.MoveNext
        t = t + 1
    Loop
    get_sieger = line
End Function

Public Sub make_a_schedule()
' Moderator Zeitplan
    If get_mod_on Then Exit Sub
    
    Dim db As Database
    Dim re As Recordset
    Dim ht As Recordset
    Dim out As Variant
    Dim ht_pfad As String
    Dim line As String
    Dim tr_nr As String

    ht_pfad = getBaseDir & "Apache2\htdocs\"
    tr_nr = "T" & Forms![A-Programmübersicht]!Turnier_Nummer
    
    Set db = CurrentDb
    Set ht = open_re("Mod", "Zeit")
    Set re = db.OpenRecordset("SELECT Startklasse.Startklasse, RT_ID, Rundentab.Startzeit, rt_stat, Rundentext & "" "" & Startklasse_text AS Ausdr1, HTML, Turniernr, Right([Rundentab].[Runde],3)='Fuß' AS Ausdr2, InPunkteeingabe FROM Tanz_Runden RIGHT JOIN (Startklasse RIGHT JOIN Rundentab ON Startklasse.Startklasse = Rundentab.Startklasse) ON Tanz_Runden.Runde = Rundentab.Runde WHERE (Turniernr=" & get_aktTNr & " AND Rundentab.Rundenreihenfolge<999) ORDER BY Rundentab.Rundenreihenfolge;")
    
    Set out = file_handle(ht_pfad & "moderator\index.html")
    
    re.MoveFirst
    line = Replace(ht!f1, "x_title", "Zeitplan")
    line = Replace(line, "wr_k", "&nbsp")
    line = Replace(line, "tr__txt", Forms![A-Programmübersicht]!Turnierbez)
    out.writeline (line)
    line = "    <tr height=""50"">" & vbCrLf & _
           "        <td class=""rd_m"" width=10%>&nbsp;</td>" & vbCrLf & _
           "        <td class=""rd_n"" width=50%>Offizielle</td>" & vbCrLf & _
           "        <td class=""rd_n"" width=20%><a href=""" & tr_nr & "_offizielle.html"" Target=""_blank""><img src=""../ball.red.png"" border=""0""  alt=""done""></td>" & vbCrLf & _
           "        <td class=""rd_n"" width=20%><a href=""" & tr_nr & "_WRlist.html"" Target=""_blank""><img src=""../ball.red.png"" border=""0""  alt=""done""></td>" & vbCrLf & _
           "    </tr>"
    out.writeline (line)
    Do Until re.EOF
        If (re!HTML = False And Nz(re!RT_Stat) = False) Then
            If InStr(1, re!Ausdr1, "Vorstellung der Tanzpaare") > 0 Then
                'Vorstellung der Tanzpaare
                line = "    <tr height=""50"">" & vbCrLf & _
                       "        <td class=""rt_m"" width=10%>" & Format(re!Startzeit, "hh:mm") & " </td>" & vbCrLf & _
                       "        <td class=""rd_v"" width=50%>" & re!Ausdr1 & "</td>" & vbCrLf & _
                       "        <td class=""rd_v"" width=20%><a href=""" & tr_nr & "_Vorstellung.html"" Target=""_blank""><img src=""../ball.red.png"" border=""0"" alt=""done"">" & "</td>" & vbCrLf & _
                       "        <td class=""rd_v"" width=20%>&nbsp;</td>" & vbCrLf & _
                       "    </tr>"
            Else
                line = "    <tr height=""50"">" & vbCrLf & _
                       "        <td class=""rt_z"" width=10%>" & Format(re!Startzeit, "hh:mm") & "</td>" & vbCrLf & _
                       "        <td class=""rd_n"" width=50%>" & re!Ausdr1 & "</td>" & vbCrLf & _
                       "        <td class=""rd_n"" width=20%>&nbsp;</td>" & vbCrLf & _
                       "        <td class=""rd_n"" width=20%>&nbsp;</td>" & vbCrLf & _
                       "    </tr>"
            End If
        Else
            'ist erstellt
            line = "    <tr height=""50"">" & vbCrLf & _
                   "        <td class=""rt_m"" width=10%>" & Format(re!Startzeit, "hh:mm") & " </td>" & vbCrLf & _
                   "        <td class=""rd_v"" width=50%>" & re!Ausdr1 & "</td>" & vbCrLf & _
                   "        <td class=""rd_v"" width=20%>" & IIf(Nz(re!RT_Stat), "<a href=" & tr_nr & "_Sta_K" & re!RT_ID & ".html Target=""_blank""><img src=""../ball.red.png"" border=""0"" alt=""done"">", "&nbsp;") & "</td>" & vbCrLf & _
                   "        <td class=""rd_v"" width=20%>" & IIf(re!HTML, "<a href=" & tr_nr & "_Mod_K" & re!RT_ID & ".html Target=""_blank""><img src=""../ball.red.png"" border=""0"" alt=""done"">", "&nbsp;") & " </td>" & vbCrLf & _
                   "    </tr>"
        End If
        out.writeline (line)
        re.MoveNext
    Loop
    out.writeline ("    </table>" & vbCrLf & "   </form>" & vbCrLf & "</body>" & vbCrLf & "</html>")
    out.Close

End Sub

Private Sub make_a_off()
' Moderator Offiziellenliste
    If get_mod_on Then Exit Sub
    
    Dim db As Database
    Dim ht, re As Recordset
    Dim out As Variant
    Dim ht_pfad As String
    Dim tr_nr As String
    Dim line As String
    
    Set db = CurrentDb
    Set ht = open_re("Mod", "Off")
    tr_nr = "T" & Forms![A-Programmübersicht]!Turnier_Nummer
    
    ht_pfad = getBaseDir & "Apache2\htdocs\"
    Set re = db.OpenRecordset("SELECT TL_Vorname & "" "" & TL_Nachname AS TLName, TLF_Name FROM Turnierleiter_Funktion INNER JOIN Turnierleitung ON Turnierleiter_Funktion.TLF_ID = Turnierleitung.Art ORDER BY TLF_Reihenfolge;")
        
    Set out = file_handle(ht_pfad & "moderator\" & tr_nr & "_offizielle.html")
    If re.RecordCount = 0 Then
        MsgBox "Es ist keine Turnierleitung eingetragen!"
        Exit Sub
    End If
    re.MoveFirst
    line = Replace(ht!f1, "x_title", "Offizielle")
    line = line & "    <tr>" & vbCrLf & _
           "        <td colspan=""2"" width=""720"" class=""wr_l"" >Offizielle</td>" & vbCrLf & _
           "    </tr>"
    out.writeline (line)
    Do Until re.EOF
        line = Replace(ht!F2, "rt__of", re!TLF_Name)
        line = Replace(line, "rt__wr", re!TLName)
        out.writeline (line)
        re.MoveNext
    Loop

    Set re = db.OpenRecordset("SELECT [WR_Vorname] & "" "" & [WR_Nachname] AS WRName, WR_Azubi FROM Wert_Richter ORDER BY WR_Kuerzel;")
    If re.RecordCount = 0 Then
        MsgBox "Es sind keine Wertungsrichter eingetragen!"
        Exit Sub
    End If
    re.MoveFirst
    Do Until re.EOF
        line = Replace(ht!F2, "rt__of", IIf(re!WR_AzuBi = True, "Probe oder<br>SchattenWR", "Wertungsrichter"))
        line = Replace(line, "rt__wr", re!wrname)
        out.writeline (line)
        re.MoveNext
    Loop

    out.writeline ("    </table>" & vbCrLf & "   </form>" & vbCrLf & "</body>" & vbCrLf & "</html>")
    out.Close

End Sub

Private Sub make_a_wrlist()
' Moderator Wertungsrichtereinteilung
    Dim db As Database
    Dim ht, re, wr As Recordset
    Dim out As Variant
    Dim ht_pfad As String
    Dim tr_nr As String
    Dim line As String
    
    Set db = CurrentDb
    Set ht = open_re("Mod", "Off")
    tr_nr = "T" & Forms![A-Programmübersicht]!Turnier_Nummer
    
    ht_pfad = getBaseDir & "Apache2\htdocs\moderator\"
    Set out = file_handle(ht_pfad & "" & tr_nr & "_WRlist.html")
    
    Set re = db.OpenRecordset("SELECT Startklasse_text, Startklasse.Startklasse FROM Startklasse INNER JOIN Startklasse_Turnier ON Startklasse.Startklasse = Startklasse_Turnier.Startklasse WHERE (Turniernr = " & get_aktTNr & ") ORDER BY Reihenfolge;")
    Set wr = db.OpenRecordset("SELECT Wert_Richter.* FROM Wert_Richter WHERE (Turniernr=" & get_aktTNr & " AND WR_Azubi = False) ORDER BY WR_Kuerzel;")
    re.MoveFirst
    wr.MoveLast
    line = Replace(ht!f1, "x_title", "Wertungsrichtereinteilung")
    line = line & vbCrLf & "    <tr>" & vbCrLf & _
           "        <td colspan=""" & wr.RecordCount + 1 & """ class=""wr_l"" >Wertungsrichtereinteilung</td>" & vbCrLf & _
           "    </tr>"
    
    out.writeline (line)
    line = tr & vbCrLf
    line = line & "        <td class=""rd_m"" width=""200"">" & "&nbsp" & "</td>" & vbCrLf
    wr.MoveFirst
    Do Until wr.EOF
        line = line & "        <td class=""wr_name"">" & trans_wr(wr!WR_Nachname) & "</td>" & vbCrLf

        wr.MoveNext
    Loop
    out.writeline (line & trn)
    
    re.MoveFirst
    Do Until re.EOF
        wr.MoveFirst
        line = tr & vbCrLf & "        <td class=""wr_stkl"">" & _
               re!Startklasse_text & "</td>" & vbCrLf
        
                Do Until wr.EOF
                    line = line & "        <td class=""wr_cros"">" & _
                           IIf(IsNull(DLookup("[WR_ID]", "Startklasse_Wertungsrichter", "Startklasse='" & re!Startklasse & "' AND WR_ID = " & wr!WR_ID)), "&nbsp", "X") & _
                           "</td>" & vbCrLf
            
                    wr.MoveNext
                Loop
    
        line = line & trn
        out.writeline (line)
        re.MoveNext
    Loop
    
    out.writeline ("    </table>" & vbCrLf & "   </form>" & vbCrLf & "</body>" & vbCrLf & "</html>")
    out.Close

End Sub

Private Function trans_wr(WR_Name)
    Dim i As Integer
    Dim neu As String
    For i = 1 To Len(WR_Name)
        neu = neu & Mid(WR_Name, i, 1) & "<br>"
    Next
    trans_wr = neu
End Function

Public Function get_mod_on() As Boolean
    Dim db As Database
    Dim re As Recordset
    Set db = CurrentDb
    Set re = db.OpenRecordset("SELECT Kopie_an FROM Kopien WHERE Kopie_an='HTML-Moderator'")
    If re.RecordCount = 0 Then get_mod_on = True
    
End Function

