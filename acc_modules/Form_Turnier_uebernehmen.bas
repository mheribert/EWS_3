Option Compare Database

Private Sub btnAbbrechen_Click()
    DoCmd.Close acForm, "Turnier_uebernehmen"
End Sub

Private Sub btnOK_Click()
    Dim dbs As Database
    Set dbs = CurrentDb
    Dim rst, ziel As Recordset
    Dim Land As String
    
    If Me!T_Name = "" Or Nz(Me!T_Nr) = "" Then
        MsgBox "Bitte Turniername und Turniernummer eingeben!"
    Else
        If Len(Me!T_Nr) < 6 Or Not IsNumeric(Me!T_Nr) Then
            MsgBox "Die Turniernummer muss mindestens 7 Zahlen lang sein."
            Exit Sub
        End If
        If Len(Dir(getBaseDir & "T" & Me!T_Nr & "_TDaten.mdb")) > 0 Then
            MsgBox "Turnier existiert bereits!", , "Turnierprogramm"
            Exit Sub
        End If
        make_new_TDaten Me!T_Nr
        bind_exttbl Me!T_Nr
        Set rst = dbs.OpenRecordset("Turnier")
    
        rst.AddNew
        
        rst!Turnier_Name = Me!T_Name
        rst!Turnier_Nummer = Me!T_Nr
        rst!T_Datum = Me!T_Datum
        rst!Veranst_Ort = Me!T_Ort
        rst!Veranst_Clubnr = 0
        If (IsNumeric(Me!T_VereinNr)) Then
            rst!Veranst_Clubnr = Me!T_VereinNr
        End If
        
        If (IsDate(Me!T_Anfang)) Then
            rst!Anfang = Me!T_Anfang
        End If
        If (IsDate(Me!T_Ende)) Then
            rst!Ende = Me!T_Ende
        End If
        
        If (Me!T_Veranstalter <> "") Then
            rst!Veranst_Name = Me!T_Veranstalter
        End If
        Land = get_properties("LAENDER_VERSION")
        rst!BS_Erg = Land
        rst.Update
        dbs.Execute "DELETE * FROM Startklasse WHERE Land<>'" & Land & "';"
        dbs.Execute "DELETE * FROM Tanz_Runden_fix WHERE Land<>'" & Land & "';"
        
        Set rst = dbs.OpenRecordset("SELECT * FROM TLP_Offizielle WHERE ((([WVorname] & "" "" & [WName])='" & Me!Turnierleiter & "') AND ((TLP_OFFIZIELLE.Lizenz)='TL'));")
        If rst.RecordCount > 0 Then
            Set ziel = dbs.OpenRecordset("Turnierleitung")
            ziel.AddNew
            ziel!TL_Vorname = rst!WVorname
            ziel!TL_Nachname = rst!WName
            ziel!Lizenznr = rst!Lizenzn
            ziel!Vereinsnr = rst!Club
            ziel!Turniernr = 1
            ziel!Art = "TL"
            ziel.Update
        End If
        rst.Close
        btnAbbrechen_Click
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
    If get_properties("LAENDER_VERSION") = "D" Then
        Me.ListeTurnierdaten.Height = 5900
        Me!ListeTurnierdaten.RowSource = "SELECT TLP_TERMINE.Terminnummer AS Turniernr, TLP_TERMINE.Datum, TLP_TERMINE.Bezeichnung, [PLZ] & "" "" & [Ort] AS Name, TLP_TERMINE.PLZ, TLP_TERMINE.Ort, TLP_TERMINE.Mitgliedsnr, TLP_TERMINE.Raum, TLP_TERMINE.Straße, TLP_TERMINE.Beginn, TLP_TERMINE.Ende, TLP_TERMINE.Clubname_kurz, Left([Terminnummer],1) AS Ausdr1, TLP_TERMINE.Turnierleiter FROM TLP_TERMINE WHERE (((TLP_TERMINE.Datum)>=Now()-1) AND ((Left([Terminnummer],1))=1)) ORDER BY TLP_TERMINE.Datum, [PLZ] & "" "" & [Ort], TLP_TERMINE.Bezeichnung;"
    Else
        Me.ListeTurnierdaten.Height = 2660
        Me!ListeTurnierdaten.RowSource = "SELECT TLP_TERMINE.Terminnummer AS Turniernr, TLP_TERMINE.Datum, TLP_TERMINE.Bezeichnung, [PLZ] & "" "" & [Ort] AS Name, TLP_TERMINE.PLZ, TLP_TERMINE.Ort, TLP_TERMINE.Mitgliedsnr, TLP_TERMINE.Raum, TLP_TERMINE.Straße, TLP_TERMINE.Beginn, TLP_TERMINE.Ende, TLP_TERMINE.Clubname_kurz, Left([Terminnummer],1) AS Ausdr1, TLP_TERMINE.Turnierleiter FROM TLP_TERMINE WHERE (((TLP_TERMINE.Datum)>=Now()-1) AND ((Left([Terminnummer],1))=2)) ORDER BY TLP_TERMINE.Datum, [PLZ] & "" "" & [Ort], TLP_TERMINE.Bezeichnung;"
    End If
End Sub

Private Sub ListeTurnierdaten_Click()
    Me!T_Name = ListeTurnierdaten.Column(2)
    Me!T_Nr = ListeTurnierdaten.Column(0)
    Me!T_Datum = ListeTurnierdaten.Column(1)
    Me!T_Ort = ListeTurnierdaten.Column(3)
    Me!T_VereinNr = 0
    If (IsNumeric(ListeTurnierdaten.Column(6))) Then
        Me!T_VereinNr = ListeTurnierdaten.Column(6)
    End If
    
    If (IsDate(ListeTurnierdaten.Column(9))) Then
        Me!T_Anfang = ListeTurnierdaten.Column(9)
    End If
    If (IsDate(ListeTurnierdaten.Column(10))) Then
        Me!T_Ende = ListeTurnierdaten.Column(10)
    End If
    Me!T_Veranstalter = ListeTurnierdaten.Column(11)
    Me!Turnierleiter = ListeTurnierdaten.Column(13)
    
End Sub

Private Sub ListeTurnierdaten_DblClick(Cancel As Integer)
    ListeTurnierdaten_Click
    btnOK_Click
End Sub
