Option Compare Database
Option Explicit

    Public gReportStartklasse As String
    Dim stDocName As String
    Private Const VK_SHIFT = &H10

Private Sub Befehl107_Click()   'Export des Turnierberichts
    Dim TBerichtName As String
    If turnier_selected Then Exit Sub

    [Form_A-Programmübersicht]!Report_Turniernum = [Form_A-Programmübersicht]!Akt_Turnier
    ' 20111118 HK Turnierbericht als RTF speichern
    TBerichtName = gen_Ordner(getBaseDir & "_Versand\") & get_TerNr & "_Turnierbericht.rtf"
    DoCmd.OutputTo acOutputReport, "Turnierbericht", "RichTextFormat(*.rtf)", TBerichtName, False, ""
    MsgBox ("Den abgespeicherten Turnierbericht " & TBerichtName & ", per eMail an die, in der TSO vermerkte Position, versenden!")
End Sub

Private Sub Befehl12_Click()
    If turnier_selected Then Exit Sub

    stDocName = "Turnier aufnehmen"
    DoCmd.OpenForm stDocName
End Sub
 
Private Sub Befehl122_Click()
    Dim sFilepath As String
    If turnier_selected Then Exit Sub
    
    sFilepath = gen_Ordner(getBaseDir & "_Versand\") & get_TerNr & "_Rangliste" & ".xls"
        
    If Len(sFilepath) Then
        DoCmd.OutputTo acQuery, "Ergebnisliste_Text", "MicrosoftExcel(*.xls)", sFilepath, False, ""
    End If
End Sub

Private Sub Befehl125_Click()
    If turnier_selected Then Exit Sub

    stDocName = "Einstellungen"
    DoCmd.OpenForm stDocName
End Sub

Private Sub Befehl13_Click()
    If turnier_selected Then Exit Sub

    stDocName = "Wertungsrichter_aufnehmen"
    DoCmd.OpenForm stDocName
End Sub

Private Sub Befehl14_Click()
    If turnier_selected Then Exit Sub
    
    DoCmd.OpenForm "Tanzpaare_aufnehmen"
End Sub

Private Sub Befehl23_Click()
    If turnier_selected Then Exit Sub

    stDocName = "ablaufplanung"
    DoCmd.OpenForm stDocName
End Sub

Private Sub Befehl26_Click()
    start_config_webserver
End Sub

Private Sub Befehl27_Click()
    If turnier_selected Then Exit Sub

    stDocName = "Wertung_aufnehmen"
    DoCmd.OpenForm stDocName
End Sub

Private Sub Befehl33_Click()
    If turnier_selected Then Exit Sub

    stDocName = "Rundenauslosung"
    DoCmd.OpenForm stDocName
End Sub

Private Sub Befehl36_Click()
    If turnier_selected Then Exit Sub

    stDocName = "Ausdrucke"
    DoCmd.OpenForm stDocName
End Sub

Private Sub Befehl37_Click()
    If turnier_selected Then Exit Sub

    stDocName = "Paare_in erste Runde nehmen"
    DoCmd.OpenForm stDocName
End Sub

Private Sub Befehl46_Click()
    If turnier_selected Then Exit Sub

    stDocName = "Monitor_Runden"
    DoCmd.OpenForm stDocName
End Sub

Private Sub Befehl51_Click()
    If turnier_selected Then Exit Sub

    stDocName = "Majoritaet_ausrechnen"
    DoCmd.OpenForm stDocName
End Sub

Private Sub Befehl75_Click()
    If turnier_selected Then Exit Sub

    stDocName = "Aktive_uebernehmen"
    DoCmd.OpenForm stDocName
End Sub

Private Sub Befehl76_Click()
    If turnier_selected Then Exit Sub
    
    DoCmd.OpenForm "TL_BS_aufnehmen"
End Sub

Private Sub Befehl94_Click()
    If turnier_selected Then Exit Sub
    
    DoCmd.OpenForm "Paare_ohne_Startbuch"
End Sub

Private Sub btn_Dokumentation_40_Click()
    If turnier_selected Then Exit Sub

    stDocName = "Mehrkampf_Wertung_aufnehmen"
    DoCmd.OpenForm stDocName
End Sub

Private Sub btn_Dokumentation_41_Click()    'Stellprobe
    If turnier_selected Then Exit Sub
    DoCmd.OpenForm "Stellprobe"
End Sub

Private Sub btn_Dokumentation_42_Click()
    If turnier_selected Then Exit Sub
    Gen_Mail
End Sub

Private Sub btn_Dokumentation_44_Click()
    
    If (IsNull(Akt_Turnier)) Then Exit Sub
    
    versand_ausschreibung get_TerNr

End Sub

Private Sub btnErgebnisliste_Click()    'Ergebnisliste
    If (IsNull(Forms![A-Programmübersicht]![Akt_Turnier]) Or (Forms![A-Programmübersicht]![Akt_Turnier] = 0)) Then
       MsgBox ("Bitte Turnier auswählen!")
       Exit Sub
    End If
    
    Dim sFilepath As String
    
    '************* HM ** Datei wird nun in _Versand gespeichert
    sFilepath = gen_Ordner(getBaseDir & "_Versand\") & get_TerNr & "_Ergebnisliste.txt"
    If Len(sFilepath) Then
        Call writeErgebnisliste(sFilepath)
    End If
End Sub

Private Sub Form_Activate()
    
    Me.Refresh
    
    If (IsNull(Akt_Turnier)) Then
        Exit Sub
    End If
    
    Call Turnierausw_AfterUpdate
End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim db As Database
    Dim re As Recordset
    Dim T_Name As String
    Dim t_spei As String
    Dim t_sel  As String
    Dim t_Pfad As String
    Dim s_row  As String

    Set db = CurrentDb
    
    setzte_buttons "A-Programmübersicht", "Dokumentation", get_properties("LAENDER_VERSION")
    Akt_Turnier = 0
    If get_properties("Externer_Pfad") Then
        t_Pfad = get_Filename(Me.hwnd)
        t_Pfad = left(t_Pfad, InStrRev(t_Pfad, "\"))
    Else
        t_Pfad = getBaseDir()
    End If
    T_Name = Dir(t_Pfad & "T*_TDaten.mdb")
    Do Until Len(T_Name) = 0
        t_spei = T_Name
        Set db = DBEngine.Workspaces(0).OpenDatabase(t_Pfad & T_Name)
        Set re = db.OpenRecordset("Turnier", DB_OPEN_DYNASET)
        re.MoveFirst
        s_row = s_row & re!Turniernum & ";""" & re!Turnier_Name & """;""" & re!T_Datum & """;" & Nz(re!Turnier_Nummer) & ";""" & re!Veranst_Name & """;" & re!Getrennte_Auslosung & ";""" & re!Veranst_Ort & """;"""";""" & re!BS_Erg & """;"
        t_sel = re!Turnier_Name
        re.Close
        db.Close
        T_Name = Dir(t_Pfad & "T*_TDaten*.mdb")
        Do Until T_Name = t_spei
            T_Name = Dir
        Loop
        T_Name = Dir
    Loop
    If s_row <> "" Then Me!Turnierausw.RowSource = Mid(s_row, 1, Len(s_row) - 1)
    If (Turnierausw.ListCount = 1) Then
        Me!Turnierausw = t_sel
        Call Turnierausw_AfterUpdate
    End If
    
End Sub

Private Sub Neues_Turnier_Click()
    Dim stDocName As String
    stDocName = "Turnier_uebernehmen"
    DoCmd.OpenForm stDocName, , , , , acDialog
    Call Form_Open(0)
End Sub

Private Sub Turnierausw_AfterUpdate()
    Dim lae As String
    If Me!Turnierausw.Column(0) > 0 Then
        bind_exttbl Me!Turnierausw.Column(3)
        Akt_Turnier = Turnierausw.Column(0)
        Turnierauswahl = Turnierausw.Column(0)
        Turnierbez = Turnierausw.Column(1)
        Tur_Datum = Turnierausw.Column(2)
        Turnier_Nummer = Turnierausw.Column(3)
        Turnierveranstalter = Turnierausw.Column(4)
        Getrennte_Auslosung = Turnierausw.Column(5)
        Land = Turnierausw.Column(8)
        'akt_Turnier
        setzte_logo Turnierausw.Column(1)
        write_config_json getBaseDir & "webserver"
    End If
    lae = Nz(Forms![A-Programmübersicht]!Turnierausw.Column(8))
    setzte_buttons "A-Programmübersicht", "Dokumentation", IIf(lae = "", get_properties("LAENDER_VERSION"), lae)
End Sub

Sub setzte_logo(turnier)
    Dim db As Database
    Dim ht As Recordset
    Dim Buffer() As Byte
    Dim Dateigroesse As Long
    Dim BilddateiID As Long
    Dim dbPfad As String
        
    dbPfad = getBaseDir() & "\webserver\views\"
    If Len(Dir(dbPfad & turnier & ".jpg")) > 0 Then
        FileCopy dbPfad & turnier & ".jpg", dbPfad & "logo.jpg"
    Else
        dbPfad = getBaseDir()
        Set db = CurrentDb
        Set ht = db.OpenRecordset("Select * FROM HTML_Block WHERE Seite = 'Logo' and Bereich = 'Bild';")
        ht.MoveFirst
        Dateigroesse = Nz(LenB(ht!F3), 0)
        BilddateiID = FreeFile
        
        ReDim Buffer(Dateigroesse)
        Open dbPfad & Trim(ht!F1) For Binary Access Write As BilddateiID
        Buffer = ht!F3.GetChunk(0, Dateigroesse)
        Put BilddateiID, , Buffer
        Close BilddateiID
    End If
End Sub

Private Sub Befehl93_Click()
    stDocName = "Startliste_Runden"
    DoCmd.OpenReport stDocName, acPreview
End Sub

Private Sub Wertung_einlesen_Click()
    If turnier_selected Then Exit Sub
    
    stDocName = "Wertung_einlesen"
    DoCmd.OpenForm stDocName
End Sub

Function turnier_selected()
    If (IsNull(Forms![A-Programmübersicht]![Akt_Turnier]) Or (Forms![A-Programmübersicht]![Akt_Turnier] = 0)) Then
       MsgBox ("Bitte Turnier auswählen!")
       turnier_selected = True
    End If
End Function

Function Doc_btn(nr)
    Dim doc As String
    Dim lae As String
    lae = IIf(IsNull(Forms![A-Programmübersicht]!Turnierausw.Column(8)), get_properties("LAENDER_VERSION"), Forms![A-Programmübersicht]!Turnierausw.Column(8))
    If Me("btn_Dokumentation_" & nr).Caption = ". . ." Then
        MsgBox "Hier ist kein Dokument hinterlegt."
    Else
        doc = DLookup(lae & "_Dokumentation", "Dokumente", "btn = 'btn_Dokumentation_" & nr & "'")
        If InStr(doc, ".pdf") > 0 Or InStr(doc, "\") > 0 Then
            Call showDocument(doc)
        Else
            DoCmd.OpenReport doc, acViewPreview
        End If
    End If
End Function
