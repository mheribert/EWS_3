Option Compare Database

    Public sel_fld As String

Sub Tabellen_leeren()
    Dim db As Database
    Dim sql As String
    Dim i As Integer
    Dim tbls As Variant
    
    tbls = Array("TLP_BW_PAARE", "TLP_RR_PAARE", "TLP_FORMATIONEN", "Abgegebene_Wertungen", _
            "TLP_OFFIZIELLE", "TLP_TERMINE", "MSys__Akrobatiken WHERE [Nr#] <>'All'", "Show")

    Set db = CurrentDb
    For i = 0 To UBound(tbls)
        sql = "DELETE FROM " & tbls(i) & ";"
        db.Execute sql
    Next i
End Sub

'------------------------------------------------------------
' Turnier_aktuell_check
'
'------------------------------------------------------------
Sub Turnier_aktuell_check_VB()
On Error GoTo Turnier_aktuell_check_Err

    If (Eval("[Forms]![A-Programmübersicht]![Akt_Turnier] Is Null")) Then
        Beep
        MsgBox "Markieren Sie bitte auf der Programmübersicht das aktuelle Turnier!", vbOKOnly, ""
        DoCmd.OpenForm "A-Programmübersicht", acNormal, "", "", , acNormal
    End If

Turnier_aktuell_check_Exit:
    Exit Sub

Turnier_aktuell_check_Err:
    MsgBox Error$
    Resume Turnier_aktuell_check_Exit

End Sub

Public Function setzte_buttons(frm, btns, Land)
    Set db = CurrentDb
    Set re = db.OpenRecordset("SELECT * FROM Dokumente WHERE btn Like 'btn_" & btns & "_*';")
    If re.RecordCount > 0 Then re.MoveFirst
    Do Until re.EOF
        If re(Land & "_Caption") = ". . ." Then
            Forms(frm)(re!btn).Visible = False
        Else
            Forms(frm)(re!btn).Visible = True
            If Not re!Bef Then Forms(frm)(re!btn).Caption = re(Land & "_Caption")
        End If
        re.MoveNext
    Loop
End Function

' Ermittelt, ob zu einem Tanzpaar bereits Wertungen eingegeben wurden oder nicht
Public Function hasWertungen(TP_ID As Integer) As Boolean
    Dim dbs As Database
    Set dbs = CurrentDb
    
    Dim stmt As String
    
    stmt = "Select count(*) as anz from Auswertung a, Paare_Rundenqualifikation pr where pr.pr_id=a.pr_id and pr.tp_id=" & TP_ID
    stmt = stmt & " and a.Punkte is not null"
    
    Dim rst As Recordset
    Set rst = dbs.OpenRecordset(stmt)
    hasWertungen = (rst!anz > 0)
    rst.Close
End Function

Public Function getRT_ID(Turniernr As Integer, Startkl As String, Runde As String) As Integer
    Dim dbs As Database
    Set dbs = CurrentDb
    
    Dim rst As Recordset
    Set rst = dbs.OpenRecordset("Select * from View_Runden where Turniernr=" & Turniernr & " and Startklasse='" & Startkl & "' and Runde='" & Runde & "'")
    getRT_ID = rst!RT_ID
    rst.Close
End Function

Public Function get_rde(stkl, rde) As Recordset
    Dim db As Database
    
    Set db = CurrentDb
    Set get_rde = db.OpenRecordset("SELECT * from view_Runden WHERE Startklasse = '" & stkl & "' AND Runde " & IIf(InStr(1, rde, "*") > 0, "LIKE '", "= '") & rde & "' AND turniernr =" & get_aktTNr & ";", DB_OPEN_DYNASET)
     
End Function

Public Sub showDocument(url As String)
    Dim completeURL As String
    completeURL = getBaseDir() & url
    If Len(Dir(completeURL)) = 0 Then
        MsgBox "Das Dokument wurde nicht gefunden."
    Else
        FollowHyperlink completeURL
    End If
End Sub

Public Function Get_WR(wr, Startklasse)
    Dim dbs As Database
    Dim rst As Recordset
    
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("SELECT WR_function as wert FROM Wert_Richter INNER JOIN Startklasse_Wertungsrichter ON Wert_Richter.WR_ID = Startklasse_Wertungsrichter.WR_ID WHERE (((Startklasse_Wertungsrichter.Startklasse)=""" & Startklasse & """) AND ((Wert_Richter.WR_Kuerzel)=""" & wr & """)AND ((Wert_Richter.Turniernr)= " & [Forms]![A-Programmübersicht]![Akt_Turnier] & "));")
    If Not rst.EOF Then
        Get_WR = rst!wert
    End If
End Function

Public Function Get_Paare(Runde, Startklasse)
    Dim dbs As Database
    Dim rst As Recordset
    
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("SELECT Rundentab.Paare From Rundentab WHERE ((Rundentab.Startklasse = """ & Startklasse & """) And (Rundentab.runde Like """ & Runde & "*"") And (Rundentab.Rundenreihenfolge < 999)) ORDER BY Rundentab.Paare DESC;")
    

    If Not rst.EOF Then
        rst.MoveFirst
        Get_Paare = rst!Paare
    End If
    Set rst = Nothing
    Set dbs = Nothing

End Function

Public Sub make_new_TDaten(T_Nr)
    Dim wsdb As Workspace
    Dim db As Database
    Dim dbnew As Database
    Dim re As Recordset
    Dim cm As Recordset
    Dim strsql As String
    Dim nTurnier As String
    Dim i As Integer
    
    nTurnier = getBaseDir & "T" & T_Nr & "_TDaten.mdb"
    Set wsdb = DBEngine.Workspaces(0)
    Set dbnew = wsdb.CreateDatabase(nTurnier, dbLangGeneral, dbVersion40)
    DoEvents
    Set db = CurrentDb
    For i = 0 To db.TableDefs.Count - 1
        If left(db.TableDefs(i).Name, 6) = "MSys__" Then
            Application.SetHiddenAttribute acTable, db.TableDefs(i).Name, False
            DoCmd.CopyObject nTurnier, Mid(db.TableDefs(i).Name, 7), acTable, db.TableDefs(i).Name
            Application.SetHiddenAttribute acTable, db.TableDefs(i).Name, True
        End If
    Next i
    dbnew.TableDefs("Analyse").Attributes = dbHiddenObject
    DoCmd.CopyObject nTurnier, "Anwesend_Status", acTable, "Anwesend_Status"
    DoCmd.CopyObject nTurnier, "Disqualifikationsgrund", acTable, "Disqualifikationsgrund"
    DoCmd.CopyObject nTurnier, "Punktabzug", acTable, "Punktabzug"
    DoCmd.CopyObject nTurnier, "Startbuch_Status", acTable, "Startbuch_Status"
    DoCmd.CopyObject nTurnier, "Turnierleiter_Funktion", acTable, "Turnierleiter_Funktion"
    DoCmd.CopyObject nTurnier, "Tanz_Runden", acQuery, "Tanz_Runden"
    DoCmd.CopyObject nTurnier, "Properties", acTable, "Properties"
    DoCmd.CopyObject nTurnier, "View_Rundenablauf", acQuery, "View_Rundenablauf"
    
    dbnew.Execute "ALTER TABLE Paare_Rundenqualifikation " _
        & "ADD CONSTRAINT PaareRelationship " _
        & "FOREIGN KEY (TP_ID) " _
        & "REFERENCES Paare (TP_ID);"
        
    dbnew.Execute "ALTER TABLE Paare_Rundenqualifikation " _
        & "ADD CONSTRAINT RundenTabRelationship " _
        & "FOREIGN KEY (RT_ID) " _
        & "REFERENCES Rundentab (RT_ID);"
        
    dbnew.Execute "ALTER TABLE Startklasse_Wertungsrichter " _
        & "ADD CONSTRAINT WRRelationship " _
        & "FOREIGN KEY (WR_ID) " _
        & "REFERENCES Wert_Richter (WR_ID);"

    db.Close
End Sub

Public Function Pfeil_up_down(KeyCode As Integer, Shift As Integer)
    On Error GoTo Fehlerout
    If KeyCode = 40 And Shift = 0 Then
        DoCmd.GoToRecord , , acNext
        KeyCode = 0
    End If
    If KeyCode = 38 And Shift = 0 Then
        DoCmd.GoToRecord , , acPrevious
        KeyCode = 0
    End If
Fehlerout:
    If err = 2105 Then Resume Next
End Function

Public Sub start_config_webserver()
    Dim strZeile As String
    Dim neuPfad As String
    Dim nodePfad As String
    Dim retVal
    
    Call Bilderspeichern
    neuPfad = getBaseDir & "Apache2"
    If get_properties("EWS") = "EWS1" And Len(Dir(neuPfad & "\conf\httpd.conf.original")) > 0 Then
        Open neuPfad & "\conf\httpd.conf.original" For Input As #1          ' original
        Open neuPfad & "\conf\httpd.conf" For Output As #2                  ' in conf mit akt. pfad
        Do While Not EOF(1)
            Line Input #1, strZeile
            If left(strZeile, 1) <> "#" Then
                Print #2, Replace(strZeile, "C:/Apache2", Replace(neuPfad, "\", "/"))
            End If
        Loop
        Close #1
        Close #2
        retVal = Shell(neuPfad & "\bin\apache.exe", vbMinimizedNoFocus)
        If retVal = 0 Then MsgBox "Der WebServer konnte nicht gestartet werden!"
    End If
    If get_properties("EWS") = "EWS3" Then
        make_wr_zeitplan
        write_config_json
        nodePfad = getBaseDir & "webserver"
        retVal = Shell(nodePfad & "\node.exe " & nodePfad & "\server.js", vbMinimizedNoFocus)
    End If
    
End Sub

Public Sub write_config_json()
    If get_properties("Externer_Pfad") = 0 Then
        neuPfad = getBaseDir & "webserver"
        gen_Ordner neuPfad
        Open neuPfad & "\config.json" For Output As #2
        Print #2, "{""db"": """ & get_TerNr & "_TDaten.mdb"", ""pfad"": """ & Replace(getBaseDir, "\", "\\") & """, ""port"": 80}"
        Close #2
        Forms![A-Programmübersicht]!Befehl26.Visible = True
    Else
        Forms![A-Programmübersicht]!Befehl26.Visible = False
    End If
End Sub

Public Function get_properties(PROP_KEY)
    On Error Resume Next
    Dim db As Database
    Dim re As Recordset
    get_properties = ""
    Set db = CurrentDb
    Set re = db.OpenRecordset("SELECT PROP_VALUE FROM Properties WHERE Prop_Key ='" & PROP_KEY & "';")
    get_properties = Nz(re!PROP_VALUE)
End Function

Public Function get_mk()
    On Error Resume Next
    Dim db As Database
    Dim re As Recordset
    Set db = CurrentDb
    Set re = db.OpenRecordset("SELECT MehrkampfStationen FROM Turnier WHERE Turniernum=1;")
    get_mk = Nz(re!MehrkampfStationen)
End Function

Function db_Ver()
    db_Ver = get_properties("DB_VERSION") & "-" & get_properties("DB_SUBVERSION")
End Function

Public Sub Print_Givaway(RundenTab_ID, Runde)
    Dim re As Recordset
    Dim fil As String
    Set re = DBEngine(0)(0).OpenRecordset("SELECT TP_ID FROM Majoritaet WHERE  RT_ID=" & RundenTab_ID & " And RT_ID Is Not Null AND Runde_Report=1;")
'*****AB***** V13.05 - falls es sich um eine Endrunde handelt andere Abfrage ohne Runde_Report
'*****HM 14.07 ***** - auf geteilte Endrunden erweitert
    If Runde = "Endrunde" Or Runde = "Endrunde Akrobatik" Or Runde = "Schnelle Endrunde" Or Runde = "Endrunde 2" Or Runde = "MK_Tanz" Then
        Set re = DBEngine(0)(0).OpenRecordset("SELECT TP_ID FROM Majoritaet WHERE  RT_ID=" & RundenTab_ID & " And RT_ID Is Not Null;")
    End If
    If re.RecordCount = 0 Then
        MsgBox "Es gibt für diese Runde keine platzierten Paare"
    Else
        re.MoveFirst
        Do Until re.EOF
            fil = fil & IIf(Len(fil) = 0, "TP_ID=", " OR TP_ID=") & re!TP_ID
            re.MoveNext
        Loop
        stDocName = "Giveaway"
        DoCmd.OpenReport stDocName, acPreview, , fil
    End If
End Sub

