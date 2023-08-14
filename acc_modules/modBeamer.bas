Option Compare Database
Option Explicit
    
Public Function get_Filename(fenst)
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = fenst
    OpenFile.lpstrFile = String(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.flags = 0  'OFN_ALLOWMULTISELECT Or OFN_EXPLORER
    get_Filename = GetOpenFileName(OpenFile)
    ' Bei Multiselect sind Teile mit 0 getrennt
    OpenFile.lpstrFileTitle = Mid(OpenFile.lpstrFileTitle, 1, InStr(1, OpenFile.lpstrFileTitle, Chr(0)) - 1)
    OpenFile.lpstrFile = Mid(OpenFile.lpstrFile, 1, InStr(1, OpenFile.lpstrFile, Chr(0)) - 1)
    get_Filename = OpenFile.lpstrFile
    
End Function

Sub Beamer_generieren(Optional ausw)
    
    If get_properties("EWS") <> "EWS1" Then Exit Sub
    
    Dim out As Object
    Dim RT_ID As Integer
    Dim line As String
    Dim auswahl As Integer
    Dim ht_pfad As String
 
    If IsMissing(ausw) Then
        If Screen.ActiveControl.Name = "btn_ausw_1" Then
            auswahl = 6
            RT_ID = Forms!Majoritaet_ausrechnen!Startklasse
        Else
            auswahl = Nz(Forms!Wertung_einlesen!HTML_Select)
            RT_ID = Forms!Wertung_einlesen!Tanzrunde
        End If
    Else
        auswahl = ausw
    End If
    
    Select Case auswahl
        Case 1
            line = make_beamer_zeitplan(RT_ID)
        Case 2, 0
            line = make_beamer_runde(RT_ID)
        Case 3
            line = make_beamer_platzierung(RT_ID)
        Case 4
            Dim db As Database
            Dim re As Recordset
            Set db = CurrentDb
            Set re = db.OpenRecordset("SELECT * From Rundentab ORDER BY Rundenreihenfolge;")
            re.MoveFirst
            line = make_beamer_zeitplan(re!RT_ID)
            get_url_to_string_check ("http://" & GetIpAddrTable() & "/hand?msg=beamer_zeitplan&text=" & re!RT_ID)
            Set re = Nothing
            Set db = Nothing

        Case 5 'Tanzrundenergebnis z.B. KO-Runden
            line = make_beamer_runde_ergebnis(RT_ID)
        Case 6
            line = make_beamer_platzierung(RT_ID, auswahl)
            get_url_to_string_check ("http://" & GetIpAddrTable() & "/hand?msg=beamer_ranking&text=")
        Case 10
            line = make_beamer_werbung
    End Select
    If get_properties("EWS") = "EWS1" Then
        ht_pfad = getBaseDir & "Apache2\htdocs\beamer\"
        line = Replace(line, "x__zoom", "")                  ' "style=""padding:200px""")
        Set out = file_handle(ht_pfad & "index.html")
        out.writeline (line)
        out.Close
    End If

End Sub

Public Function make_beamer_werbung()
    Dim db As Database
    Dim re As Recordset
    Dim line As String
    Dim bilder As String
    Dim t As Integer
    Dim HTML_Turnier As String
    Set db = CurrentDb
    Set re = db.OpenRecordset("SELECT * FROM beamer_werbung;")
    HTML_Turnier = Forms![A-Programmübersicht]!Turnierbez
    
    line = get_line("Beamer", "Werbung", 1)  'holt HTML-Seite aus HTML-Block
    re.MoveFirst
    line = Replace(line, "x__Turnier", HTML_Turnier)
    line = Replace(line, "x__Zeit", re!werb_dauer * 1000)
    line = Replace(line, "x__Bild1", Nz(re!werb_Datei1))
    For t = 2 To 10
        If Nz(re("werb_Datei" & t)) <> "" Then
            bilder = bilder & re("werb_Datei" & t) & """, """
        End If
    Next
    line = Replace(line, "x__height", re!werb_height)
    line = Replace(line, "x__width", re!werb_width)
    line = Replace(line, "x__kopf", IIf(re!werb_kopf, "", "class=""in_vis"""))
    line = Replace(line, "x__Bildarray", left(bilder, Len(bilder) - 4))
    make_beamer_werbung = line

End Function

Function make_beamer_runde(RT_ID As Integer)
'On Error Resume Next
    Dim db As Database
    Dim re As Recordset
    Dim beginn As Boolean
    Dim HTML_Turnier As String
    Dim HTML_StNr As String
    Dim HTML_paar As String
    Dim Anz_WR As Integer
    Dim HTML_Runde As String
    Dim line As String
    Dim i As Integer
    Dim max_runden As Integer
    
    Set db = CurrentDb
    Set re = db.OpenRecordset("SELECT AnzahlWR FROM Rundentab INNER JOIN Startklasse_Turnier ON (Rundentab.Turniernr = Startklasse_Turnier.Turniernr) AND (Rundentab.Startklasse = Startklasse_Turnier.Startklasse) WHERE (Rundentab.RT_ID=" & RT_ID & ");")
    Anz_WR = re!AnzahlWR
    Set re = db.OpenRecordset("SELECT RT.Turniernr, S.Startklasse_text, TR.R_NAME_ABLAUF, RT.Rundenreihenfolge, RT.RT_ID, P.TP_ID, RT.Anz_Paare, PRQ.Auslosung, PRQ.Rundennummer, P.Startnr, PRQ.Anwesend_Status, PRQ.nochmal, P.Da_Vorname, P.Da_Nachname, P.He_Vorname, P.He_Nachname, P.Verein_Name, P.Name_Team, RT.Anz_Paare FROM (((Paare_Rundenqualifikation AS PRQ INNER JOIN Rundentab AS RT ON PRQ.RT_ID = RT.RT_ID) INNER JOIN Paare AS P ON PRQ.TP_ID = P.TP_ID) INNER JOIN Startklasse AS S ON RT.Startklasse = S.Startklasse) INNER JOIN Tanz_Runden_fix AS TR ON RT.Runde = TR.Runde WHERE (((RT.RT_ID)=" & RT_ID & ") AND ((P.TP_ID) Not In (SELECT AW.Paar_ID FROM Abgegebene_Wertungen AS AW GROUP BY AW.Paar_ID, AW.RundenTab_ID HAVING (((AW.RundenTab_ID)=" & RT_ID & "));)) AND (PRQ.Anwesend_Status=1)) ORDER BY PRQ.Rundennummer, P.Startnr;", DB_OPEN_DYNASET)
    
    If re.RecordCount <> 0 Then
        re.MoveLast
        max_runden = re!Rundennummer
        HTML_Turnier = Umlaute_Umwandeln(Forms![A-Programmübersicht]!Turnierbez) & "<br>"
        HTML_Turnier = HTML_Turnier & Umlaute_Umwandeln(re!R_NAME_ABLAUF & " " & re!Startklasse_text)
        re.MoveFirst
        For i = 1 To re!Anz_Paare
            HTML_Runde = "Runde " & re!Rundennummer & " von " & max_runden
            HTML_StNr = HTML_StNr & Space(16) & "<td class=""stnr"">" & re!Startnr & "</td>" & vbCrLf
            If Nz(re!Name_Team) = "" Then
                HTML_paar = HTML_paar & Space(16) & "<td class=""tzer"">" & Umlaute_Umwandeln(re!He_Vorname) & " " & _
                                                                            Umlaute_Umwandeln(re!He_Nachname) & "<br>" & _
                                                                            Umlaute_Umwandeln(re!Da_Vorname) & " " & _
                                                                            Umlaute_Umwandeln(re!Da_NAchname) & "<br><p class=""tver"">" & _
                                                                            Umlaute_Umwandeln(re!Verein_Name) & "</p></td>" & vbCrLf
            Else
                HTML_paar = HTML_paar & Space(16) & "<td class=""tzer"">" & Umlaute_Umwandeln(re!Name_Team) & "<br><p class=""tver"">" & _
                                                                            Umlaute_Umwandeln(re!Verein_Name) & "</p></td>" & vbCrLf
            End If
            re.MoveNext
            If re.EOF Then Exit For
        Next
    Else
       HTML_Turnier = Umlaute_Umwandeln(Forms![A-Programmübersicht]!Turnierbez)
       HTML_StNr = HTML_StNr & Space(16) & "<td class=""stnr"">&nbsp;</td>" & vbCrLf
       HTML_paar = HTML_paar & Space(16) & "<td class=""tzer"">&nbsp;</td>" & vbCrLf
    End If
    line = get_line("Beamer", "Runde", 1)  'holt HTML-Seite aus HTML-Block
    line = Replace(line, "x__Turnier", HTML_Turnier)
    line = Replace(line, "x__Runde", HTML_Runde)
    line = Replace(line, "x__StNr", HTML_StNr)
    line = Replace(line, "x__Paar", HTML_paar)
    make_beamer_runde = line
End Function

Function make_beamer_runde_ergebnis(RT_ID As Integer)
'On Error Resume Next
    Dim db As Database
    Dim re As Recordset
    Dim beginn As Boolean
    Dim HTML_Turnier As String
    Dim HTML_StNr As String
    Dim HTML_paar As String
    Dim HTML_Runde As String
    Dim line, sql As String
    Dim i As Integer
    
'letzte_runde_sql = SELECT TOP 1 AW.rh FROM Abgegebene_Wertungen AS AW INNER JOIN Majoritaet AS M ON (M.TP_ID = AW.Paar_ID) AND (AW.RundenTab_ID = M.RT_ID) GROUP BY AW.rh, M.RT_ID, AW.RundenTab_ID HAVING (AW.RundenTab_ID=4) ORDER BY AW.rh DESC;
    Set db = CurrentDb
    sql = "SELECT RT.Turniernr, S.Startklasse_text, TR.R_NAME_ABLAUF, RT.Rundenreihenfolge, RT.RT_ID, RT.Anz_Paare, PRQ.Auslosung, PRQ.Rundennummer, P.Startnr, PRQ.Anwesend_Status, PRQ.nochmal, P.Da_Vorname, P.Da_Nachname, P.He_Vorname, P.He_Nachname, P.Verein_Name, P.Name_Team, RT.Anz_Paare, Majoritaet.Platz, Majoritaet.KO_Sieger"
    sql = sql & " FROM Majoritaet INNER JOIN (((Rundentab AS RT INNER JOIN Startklasse AS S ON RT.Startklasse = S.Startklasse) INNER JOIN Tanz_Runden_fix AS TR ON RT.Runde = TR.Runde) INNER JOIN (Paare AS P INNER JOIN Paare_Rundenqualifikation AS PRQ ON P.TP_ID = PRQ.TP_ID) ON RT.RT_ID = PRQ.RT_ID) ON (Majoritaet.TP_ID = PRQ.TP_ID) AND (Majoritaet.RT_ID = PRQ.RT_ID) WHERE (((RT.RT_ID)=" & RT_ID & ") AND ((PRQ.Rundennummer)=(SELECT TOP 1 AW.rh FROM Abgegebene_Wertungen AS AW INNER JOIN Majoritaet AS M ON (M.TP_ID = AW.Paar_ID) AND (AW.RundenTab_ID = M.RT_ID) GROUP BY AW.rh, M.RT_ID, AW.RundenTab_ID HAVING (AW.RundenTab_ID=" & RT_ID & ") ORDER BY AW.rh DESC)) AND ((PRQ.Anwesend_Status)=1)) ORDER BY PRQ.Rundennummer, P.Startnr;"
    Set re = db.OpenRecordset(sql, DB_OPEN_DYNASET)
    
    If re.RecordCount <> 0 Then
        re.MoveLast
        HTML_Turnier = Umlaute_Umwandeln(Forms![A-Programmübersicht]!Turnierbez) & "<br>"
        HTML_Turnier = HTML_Turnier & Umlaute_Umwandeln(re!R_NAME_ABLAUF & " " & re!Startklasse_text)
        re.MoveFirst
        For i = 1 To re!Anz_Paare
            HTML_Runde = "Runde " & re!Rundennummer
            HTML_StNr = HTML_StNr & Space(16) & "<td class=""stnr"">" & re!Startnr & IIf(re!Ko_Sieger, "<img src=""res/logo_winner.svg"" width=""90"" height=""80""alt=""DRBV"">", "") & "</td>" & vbCrLf
            If Nz(re!Name_Team) = "" Then
                HTML_paar = HTML_paar & Space(16) & "<td class=""tzer"">" & Umlaute_Umwandeln(re!He_Vorname) & " " & _
                                                                            Umlaute_Umwandeln(re!He_Nachname) & "<br>" & _
                                                                            Umlaute_Umwandeln(re!Da_Vorname) & " " & _
                                                                            Umlaute_Umwandeln(re!Da_NAchname) & "<br><p class=""tver"">" & _
                                                                            Umlaute_Umwandeln(re!Verein_Name) & "</p></td>" & vbCrLf
            Else
                HTML_paar = HTML_paar & Space(16) & "<td class=""tzer"">" & Umlaute_Umwandeln(re!Name_Team) & "<br><p class=""tver"">" & _
                                                                            Umlaute_Umwandeln(re!Verein_Name) & "</p></td>" & vbCrLf
            End If
            re.MoveNext
            If re.EOF Then Exit For
        Next
    Else
       HTML_Turnier = Umlaute_Umwandeln(Forms![A-Programmübersicht]!Turnierbez)
       HTML_StNr = HTML_StNr & Space(16) & "<td class=""stnr"">&nbsp;</td>" & vbCrLf
       HTML_paar = HTML_paar & Space(16) & "<td class=""tzer"">&nbsp;</td>" & vbCrLf
    End If
    line = get_line("Beamer", "Runde", 1)  'holt HTML-Seite aus HTML-Block
    line = Replace(line, "x__Turnier", HTML_Turnier)
    line = Replace(line, "x__Runde", HTML_Runde)
    line = Replace(line, "x__StNr", HTML_StNr)
    line = Replace(line, "x__Paar", HTML_paar)
    make_beamer_runde_ergebnis = line
End Function

Function make_beamer_zeitplan(RT_ID As Integer)
'On Error Resume Next
    Dim db As Database
    Dim re As Recordset
    Dim beginn As Boolean
    Dim HTML_Turnier As String
    Dim HTML_text As String
    Dim line As String
    
    Set db = CurrentDb
    Set re = db.OpenRecordset("SELECT RT.RT_ID, RT.Turniernr, RT.Rundenreihenfolge, RT.Startzeit, Startklasse_text, Rundentext FROM Tanz_Runden INNER JOIN (Rundentab AS RT LEFT JOIN Startklasse ON RT.Startklasse = Startklasse.Startklasse) ON Tanz_Runden.Runde = RT.Runde WHERE RT.Rundenreihenfolge < 999 ORDER BY RT.Rundenreihenfolge;", DB_OPEN_DYNASET)
    beginn = False
    re.MoveFirst
    Do Until re.EOF
        If re!RT_ID = RT_ID Then beginn = True

        If beginn And re!Rundentext <> "Letzte Startbuchabgabe" Then
            HTML_text = HTML_text & Space(24) & "<tr class=""odd"">" & vbCrLf
            HTML_text = HTML_text & Space(28) & "<td>" & Format(re!Startzeit, "hh:mm") & "</td>" & vbCrLf
            HTML_text = HTML_text & Space(28) & "<td>" & Umlaute_Umwandeln(re!Rundentext & "  " & re!Startklasse_text) & "</td>" & vbCrLf
            HTML_text = HTML_text & Space(24) & "</tr>" & vbCrLf
        End If
        re.MoveNext
    Loop
    line = get_line("Beamer", "Zeitplan", 1)  'holt HTML-Seite aus HTML-Block
    HTML_Turnier = Umlaute_Umwandeln(Forms![A-Programmübersicht]!Turnierbez)
    HTML_Turnier = HTML_Turnier & "<br>Zeitplan"
    line = Replace(line, "x__Turnier", HTML_Turnier)
    line = Replace(line, "x__Zeit", HTML_text)
    make_beamer_zeitplan = line
End Function

Function make_beamer_platzierung(RT_ID As Integer, Optional sieger)
'On Error Resume Next
    Dim getanztePaare_SQL, Tanzrunde_SQL, Startklasse_Txt, Runde_txt As String
    Dim Letzte_Runde_SQL As String
    Dim db As Database
    Dim getanztePaare, Tanzrunde, Letzte_Runde As Recordset
    Dim Counter_i As Integer
    Dim Paare_weiter_anzahl As Integer
    Dim Paare_weiter As Recordset
    Dim laufende_Tanzrunde As Integer
    Dim HTML_Turnier As String
    Dim HTML_text As String
    Dim HTML_paar As String
    Dim HTML_class As String
    Dim line As String
    Dim Platz As String
    
    Set db = CurrentDb()

    Tanzrunde_SQL = "SELECT Rundentab.RT_ID, Rundentab.Rundenreihenfolge, Startklasse.Startklasse_text, Tanz_Runden_fix.Rundentext, Tanz_Runden_fix.R_IS_ENDRUNDE FROM (Rundentab INNER JOIN Startklasse ON Rundentab.Startklasse = Startklasse.Startklasse) INNER JOIN Tanz_Runden_fix ON Rundentab.Runde = Tanz_Runden_fix.Runde WHERE (((Rundentab.RT_ID)=" & RT_ID & "));"
    Set Tanzrunde = db.OpenRecordset(Tanzrunde_SQL)
    
    Tanzrunde.MoveFirst
    Startklasse_Txt = Umlaute_Umwandeln(Tanzrunde!Startklasse_text)
    
    Set Paare_weiter = db.OpenRecordset("SELECT RT.RT_ID, RT.Rundenreihenfolge, S.Startklasse_text, RT.Paare FROM (Rundentab AS RT INNER JOIN Startklasse AS S ON RT.Startklasse = S.Startklasse) INNER JOIN Tanz_Runden_fix AS TR ON RT.Runde = TR.Runde WHERE (((S.Startklasse_text)='" & Tanzrunde!Startklasse_text & "')) ORDER BY RT.Rundenreihenfolge, TR.Rundenreihenfolge;")
    
    Paare_weiter.FindFirst "RT_ID = " & RT_ID
    Paare_weiter.MoveNext
    
    Paare_weiter_anzahl = 0
    
    If Tanzrunde!R_IS_ENDRUNDE = 1 Then
        Runde_txt = "Endrunde"
        Paare_weiter_anzahl = 7
    Else
        Runde_txt = Tanzrunde!Rundentext
        If Not IsNull(Paare_weiter!Paare) Then Paare_weiter_anzahl = Nz(Paare_weiter!Paare)
    End If
    
    Letzte_Runde_SQL = "SELECT AW.rh FROM Abgegebene_Wertungen AS AW INNER JOIN (Paare AS P INNER JOIN Majoritaet AS M ON P.TP_ID = M.TP_ID) ON (AW.Paar_ID = P.TP_ID) AND (AW.RundenTab_ID = M.RT_ID) GROUP BY M.RT_ID, AW.RundenTab_ID, AW.rh HAVING (((AW.RundenTab_ID)=" & RT_ID & ")) ORDER BY AW.rh DESC;"
    Set Letzte_Runde = db.OpenRecordset(Letzte_Runde_SQL)
    
    getanztePaare_SQL = "SELECT M.RT_ID, M.TP_ID, M.WR7, M.Platz, M.KO_Sieger, P.Startkl, P.Startnr, P.Da_Vorname, P.Da_Nachname, P.He_Vorname, P.He_Nachname, P.Verein_Name, P.Name_Team, Paare_Rundenqualifikation.Rundennummer, m.DQ_ID, m.PA_ID FROM (Paare AS P INNER JOIN Majoritaet AS M ON P.TP_ID = M.TP_ID) INNER JOIN Paare_Rundenqualifikation ON (M.RT_ID = Paare_Rundenqualifikation.RT_ID) AND (M.TP_ID = Paare_Rundenqualifikation.TP_ID) WHERE (M.RT_ID=" & RT_ID & " AND (Paare_Rundenqualifikation.PR_ID) In (SELECT [PR_ID] FROM [Auswertung] )) ORDER BY M.Platz;"
    Set getanztePaare = db.OpenRecordset(getanztePaare_SQL)
    
    If getanztePaare.RecordCount > 0 Then
        Counter_i = 1
        getanztePaare.MoveFirst
        Do Until getanztePaare.EOF
            If Paare_weiter_anzahl = Counter_i - 1 Then
                HTML_text = HTML_text & Space(28) & "<tr class=""trenn"">" & vbCrLf
                HTML_text = HTML_text & Space(32) & "<td>&nbsp;</td>" & vbCrLf
                HTML_text = HTML_text & Space(32) & "<td>&nbsp;</td>" & vbCrLf
                HTML_text = HTML_text & Space(32) & "<td>&nbsp;</td>" & vbCrLf
                HTML_text = HTML_text & Space(32) & "<td>&nbsp;</td>" & vbCrLf
                HTML_text = HTML_text & Space(28) & "</tr>" & vbCrLf
            End If
            If Nz(getanztePaare!Name_Team) = "" Then
                HTML_paar = Umlaute_Umwandeln(getanztePaare!Da_Vorname) & " "
                HTML_paar = HTML_paar & Umlaute_Umwandeln(getanztePaare!Da_NAchname)
                HTML_paar = HTML_paar & " - " & Umlaute_Umwandeln(getanztePaare!He_Vorname) & " "
                HTML_paar = HTML_paar & Umlaute_Umwandeln(getanztePaare!He_Nachname)
            Else
                HTML_paar = Umlaute_Umwandeln(getanztePaare!Name_Team)
            End If
            ' zuletzt getanztes Paar anzeigen
            If Paare_weiter_anzahl >= Counter_i Then
                If Letzte_Runde!rh = getanztePaare!Rundennummer And IsMissing(sieger) Then
                    Platz = "*" & getanztePaare!Platz
                    If Tanzrunde!Rundentext = "KO-Runde" And getanztePaare!Ko_Sieger = False Then
                        HTML_class = "<tr class=""weiter"" style=""background-color:#ffc0c0;"">"
                    Else
                        HTML_class = "<tr class=""weiter"" style=""background-color:#c0ffc0;"">"
                    End If
                Else
                    Platz = getanztePaare!Platz
                    HTML_class = "<tr class=""weiter"">"
                End If
            Else
                If Letzte_Runde!rh = getanztePaare!Rundennummer And IsMissing(sieger) Then
                    Platz = "*" & getanztePaare!Platz
                    If Tanzrunde!Rundentext = "KO-Runde" And getanztePaare!Ko_Sieger = False Then
                        HTML_class = "<tr class=""raus"" style=""background-color:#ffc0c0;"">"
                    Else
                        HTML_class = "<tr class=""raus"" style=""background-color:#c0ffc0;"">"
                    End If
                Else
                    Platz = getanztePaare!Platz
                    HTML_class = "<tr class=""raus"">"
                End If
            End If
            
            HTML_text = HTML_text & Space(28) & HTML_class & vbCrLf
            If getanztePaare!DQ_ID > 0 Or getanztePaare!PA_ID > 0 Then
                HTML_text = HTML_text & Space(32) & "<td>" & getanztePaare!Platz & "&nbsp;*</td>" & vbCrLf
            Else
                HTML_text = HTML_text & Space(32) & "<td>" & Platz & "</td>" & vbCrLf
            End If
            HTML_text = HTML_text & Space(32) & "<td>" & getanztePaare!Startnr & "</td>" & vbCrLf
            HTML_text = HTML_text & Space(32) & "<td class=""text_left"">" & HTML_paar & "</td>" & vbCrLf
'            If Left(Startklasse_Txt, 3) = "BW " Then
'                HTML_text = HTML_text & Space(32) & "<td>&nbsp;</td>" & vbCrLf
'            Else
                HTML_text = HTML_text & Space(32) & "<td>" & Format(getanztePaare!WR7, "##,##0.00") & "</td>" & vbCrLf
'            End If
            HTML_text = HTML_text & Space(28) & "</tr>" & vbCrLf
            
            getanztePaare.MoveNext
            Counter_i = Counter_i + 1
        Loop
    End If
    line = get_line("Beamer", "Platzierung", 1)  'holt HTML-Seite aus HTML-Block
    HTML_Turnier = Umlaute_Umwandeln(Forms![A-Programmübersicht]!Turnierbez) & "<br>"
    HTML_Turnier = HTML_Turnier & Umlaute_Umwandeln(Tanzrunde!Rundentext & " " & Tanzrunde!Startklasse_text)
    line = Replace(line, "x__Turnier", HTML_Turnier)
    line = Replace(line, "x__Platzierung", HTML_text)
    make_beamer_platzierung = line
End Function
