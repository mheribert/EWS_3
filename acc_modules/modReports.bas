Option Compare Database

    Public away_werte

Public Function rep_show_lines(beName, Trennlinien)
    For i = 12 To 19
        beName("Linie" & (i)).Visible = False
    Next
    For i = 0 To UBound(Trennlinien)
        beName("Linie" & Trennlinien(i)).Visible = True
    Next

End Function

Public Function rep_fill_fields(beName, fld, PR_ID, Cgi_Input, rde)
    Dim cgivar
    Dim ft_var, verstoss
    Dim i As Integer
    Dim t As Integer
    Dim str As String
    Dim Runde As String
    Dim akro As String
    Dim akro_wert  As Single
    Dim wert As Variant
    Dim v_ak As String
    If IsNull(Cgi_Input) Or Nz(beName(fld)) = "" Then
        str = ""
    Else
        Set cgivar = zerlege(Cgi_Input)
        i = eins_zwei(PR_ID, cgivar)
       
        If (left(rde, 3) = "MK_" And left(rde, 4) <> "MK_5") Or beName!Startklasse = "RR_S1" Or beName!Startklasse = "RR_S2" Then
            Select Case fld
                Case "Ber1"
                    str = cgivar.Item("wmk_th1")
                Case "Ber2"
                    str = cgivar.Item("wmk_dh1")
                Case "Ber3"
                    str = cgivar.Item("wmk_td1")
                Case "Ber4"
                    str = cgivar.Item("wmk_dd1")
            End Select
        Else
            Select Case left(beName(fld), 3)
                Case "wak"
                    Runde = ch_runde(beName!Unterformular1!Runde)
                    For t = 1 To 8
                        v_ak = Trim(cgivar.Item("tflak" & i & t))
                        If cgivar.Item("wak" & i & t) <> "" Then
                            akro = DLookup("Akro" & t & "_" & Runde, "Paare", "TP_ID=" & PR_ID)
                            wert = CSng(DLookup(beName.Startklasse, "Akrobatiken", "Akrobatik='" & akro & "' AND " & beName.Startklasse & "<> ''"))
                            akro_wert = 100 - CSng(cgivar.Item("wak" & i & t)) / wert * 100
                            str = str & Round(akro_wert, 0) & IIf(Len(v_ak) > 0, " / " & v_ak, "")
                            str = Replace(str, " /  / ", "")
                        End If
                        If cgivar.Item("wak" & i & t) = "" And Len(v_ak) > 0 Then
                            str = str & v_ak
                        End If
                        str = str & vbCrLf
                    Next
                    If Right(str, 3) = " / " Then
                        str = left(str, Len(str) - 3)
                    End If
                    If cgivar.exists("wak" & i) Then
                        str = cgivar.Item("wak" & i)
                    End If
                Case "Ob"
                    verstoss = fill_verstoss
                    For t = 0 To 7
                        If cgivar.Item("w" & verstoss(t, 0) & i) <> "" Then
                            str = str & verstoss(t, 0) & "  " & CSng(cgivar.Item("w" & verstoss(t, 0) & i)) & vbCrLf
                        End If
                    Next
                Case "f_w"
                    Set ft_var = ft_wertung(cgivar.Item("WR_ID"), cgivar.Item("rt_ID"), beName.Startklasse, PR_ID)
                    i = eins_zwei(PR_ID, cgivar)
                    If (ft_var.Item(Mid(Nz(beName(fld)), 3) & i) = "" And cgivar.Item(Mid(Nz(beName(fld)), 3) & i) = "") Then
                        str = "/"
                    ElseIf fld = "Ber8" Or fld = "Ber9" Then
                        str = ft_var.Item(Mid(Nz(beName(fld)), 3) & i) & " / " & cgivar.Item(Mid(Nz(beName(fld)), 3) & i)
                    Else
                        str = (10 - CSng(ft_var.Item(Mid(Nz(beName(fld)), 3) & i))) * 10 & " / " & (10 - CSng(cgivar.Item(Mid(Nz(beName(fld)), 3) & i))) * 10
                    End If
                Case "f_t"
                    Set ft_var = ft_wertung(cgivar.Item("WR_ID"), cgivar.Item("rt_ID"), beName.Startklasse, PR_ID)
                    i = eins_zwei(PR_ID, cgivar)
                    str = ft_var.Item(Mid(Nz(beName(fld)), 3) & i) & " / " & cgivar.Item(Mid(Nz(beName(fld)), 3) & i)
                Case "RR_"
                        str = cgivar.Item("wsh" & i) & " / " & cgivar.Item("wsd" & i)
                Case Else
                    If left(beName!Startklasse, 3) = "BW_" Then
                        str = Boogie_auswertung(beName, cgivar, fld, i, rde)
                    ElseIf left(beName!Startklasse, 3) = "RR_" Or left(beName!Startklasse, 3) = "F_R" Then
                        If fld = "Ber8" Or fld = "Ber9" Or cgivar.Item(beName(fld) & i) = "" Then
                            str = cgivar.Item(beName(fld) & i)
                            If cgivar.Item("wfl" & i & "a20") = "20" And fld = "Ber8" Then
                                str = str & " / A20"
                            End If
                        Else
                            str = (10 - CSng(cgivar.Item(beName(fld) & i))) * 10
                        End If
                    Else
                        str = cgivar.Item(beName(fld) & i)
                    End If
            End Select
        End If
    End If
    If str = "0" And InStr(1, beName(fld), "wfl") > 0 Then str = ""
    If str = " /" Then str = ""
    rep_fill_fields = Trim(str)
End Function

Function Boogie_auswertung(beName, cgivar, fld, aIndex, rde)
    Dim kl_punkte
    Dim erg As String
    If away_werte = 1 Then
        kl_punkte = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
    Else
        kl_punkte = Punkteverteilung("BW_NG", ch_runde(rde), rde)
    End If
    Select Case fld
        Case "Ber1"     'ttd
            erg = cgivar.Item("wng_ttd" & aIndex) * (kl_punkte(0))
        Case "Ber2"     'tth
            erg = cgivar.Item("wng_tth" & aIndex) * (kl_punkte(1))
        Case "Ber3"     ' Basic Dancing,Lead
            erg = cgivar.Item("wng_bda" & aIndex) * (kl_punkte(2)) & "/" & cgivar.Item("wng_dap" & aIndex) * (kl_punkte(3))
        Case "Ber4"     ' Basic Dancing,Lead Bonus
            erg = cgivar.Item("wng_bdb" & aIndex) * (kl_punkte(4))
        Case "Ber5"     ' Tanzfiguren
            erg = cgivar.Item("wng_fta" & aIndex) * (kl_punkte(5)) & "/" & cgivar.Item("wng_fts" & aIndex) * (kl_punkte(6))
        Case "Ber6"     ' Tanzfiguren Bonus
            erg = cgivar.Item("wng_ftb" & aIndex) * (kl_punkte(7))
        Case "Ber10" 'interpretation
            erg = cgivar.Item("wng_inf" & aIndex) * (kl_punkte(8)) & "/" & cgivar.Item("wng_ins" & aIndex) * (kl_punkte(9))
        Case "Ber7" ' interpretation Bonus
            erg = cgivar.Item("wng_inb" & aIndex) * (kl_punkte(10))
        Case Else
            erg = ""
    End Select
    Boogie_auswertung = erg
End Function

Function Boogie_auswertung_alt(beName, cgivar, fld, aIndex, rde)
    Dim kl_punkte
    Dim erg As String
    kl_punkte = Punkteverteilung(beName!Startklasse, ch_runde(rde), rde)
    If ch_runde(rde) = "VR" Then
        Select Case fld
            Case "Ber1"
                erg = cgivar.Item(beName(fld) & aIndex) * (kl_punkte(0) + kl_punkte(1)) / 10
            Case "Ber2"
                erg = ""
            Case "Ber5"
                erg = cgivar.Item(beName(fld) & aIndex) * (kl_punkte(4) + kl_punkte(5)) / 10
            Case "Ber6"
                erg = ""
            Case "Ber10"
                erg = cgivar.Item(beName(fld) & aIndex) * kl_punkte(6) / 10
            Case Else
                erg = cgivar.Item(beName(fld) & aIndex) * kl_punkte(Right(fld, 1) - 1) / 10
        End Select
    Else
        If fld = "Ber10" Then
            erg = cgivar.Item(beName(fld) & aIndex) * kl_punkte(6) / 10
        Else
            erg = cgivar.Item(beName(fld) & aIndex) * kl_punkte(Right(fld, 1) - 1) / 10
        End If
    End If
    Boogie_auswertung_alt = erg
End Function

Function ft_wertung(wr, rtid, rd, PR_ID)
    Dim db As Database
    Dim re As Recordset
    Set db = CurrentDb()
    '  Fußtechnik suchen
    Set re = db.OpenRecordset("SELECT * FROM rundentab WHERE rt_id =" & rtid & ";")
    Set re = db.OpenRecordset("SELECT * from RundenTab WHERE Startklasse = '" & rd & "' AND Runde = '" & left(re!Runde, 3) & "_r_Fuß';", DB_OPEN_DYNASET)
    Set re = db.OpenRecordset("SELECT * FROM Paare_Rundenqualifikation INNER JOIN Auswertung ON Paare_Rundenqualifikation.PR_ID = Auswertung.PR_ID WHERE (Paare_Rundenqualifikation.TP_ID=" & PR_ID & " AND WR_ID=" & wr & " AND RT_ID=" & re!RT_ID & ");", DB_OPEN_DYNASET)
    
    If re.EOF Then
        'MsgBox "Es existiert keine Fußtechnik"
        Set ft_wertung = zerlege("rh1=A")
    Else
        Set ft_wertung = zerlege(re!Cgi_Input)
    End If

End Function

Public Sub read_raw()
    Dim dbs As Database
    Dim re As Recordset
    Dim ziel As Recordset
    Dim fName As String
    Dim pr As Integer
    Dim strZeile
    Dim vars
    Set dbs = CurrentDb
    Set re = dbs.OpenRecordset("Rundentab")
    Set ziel = dbs.OpenRecordset("show")
    
    dbs.Execute "Delete FROM show;"
    
    If re.RecordCount > 0 Then re.MoveFirst
    Do Until re.EOF
        If get_properties("EWS") = "EWS3" Then
            fName = getBaseDir & get_TerNr() & "_" & "RT" & re!RT_ID & "_raw.txt"
        Else
            fName = getBaseDir & get_TerNr() & "_" & "RT" & re!RT_ID & ".txt"
        End If
        If Dir(fName) <> "" Then
            Open fName For Input As #1
            Do While Not EOF(1)
                Line Input #1, strZeile
                vars = Split(strZeile, ";")
                pr = DLookup("PR_ID", "paare_rundenqualifikation", "TP_ID=" & vars(0) & " AND RT_ID=" & re!RT_ID)
'                If Val(vars(1)) = Me!WR_ID Then
                    Set ziel = dbs.OpenRecordset("SELECT * FROM show WHERE pr_id=" & pr & " AND wr_id =" & vars(1) & " AND Left([Cgi_Input],3) = 'Obs';")
                    If ziel.EOF() Then
                        ziel.AddNew
                        ziel!PR_ID = pr
                        ziel!WR_ID = vars(1)
                    Else
                        ziel.Edit
                    End If
                    ziel!Cgi_Input = vars(2)
                    ziel.Update
'                End If
            Loop
            Close #1
        End If
        re.MoveNext
    Loop

'    dbs.Execute "INSERT INTO show (PR_ID, WR_ID, SH_Wert, cgi_input) SELECT PR_ID, WR_ID, Punkte, cgi_input FROM auswertung WHERE WR_ID=" & Me!WR_ID & ";"
    

End Sub
