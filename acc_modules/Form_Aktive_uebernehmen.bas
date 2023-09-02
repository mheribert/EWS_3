Option Compare Database
Option Explicit

Private Sub AuswahlBW_AfterUpdate()
    gewvnr = Me!AuswahlBW.Column(0)
    BDame_auswählen.Requery
    bw_startliste.Requery

End Sub

Private Sub AuswahlFO_AfterUpdate()
    gewvnr = Forms!Aktive_uebernehmen!AuswahlFO.Column(0)
    Formation_liste.Requery
    Formation_auswahl.Requery

End Sub

Private Sub AuswahlRR_AfterUpdate()
    If ([Klassen] <> Null) Then
        gewkl = Mid([Klassen], InStr(1, [Klassen], "_") + 1)
    End If

End Sub

Private Sub AuswahlRR_Click()
    gewvnr = [AuswahlRR]
    If ([AuswahlRR].ListIndex = -1) Then
        gewvnr = -9999
        [AuswahlRR] = -9999
    End If
    [Klassen].Requery
    Paare_in_Startliste.Requery

End Sub

Public Sub Befehl114_Click()
    If Nz(gewvnr) < 0 Then
        MsgBox "Bitte wählen Sie einen Verein aus!"
        Exit Sub
    End If
    If Nz(BSTkarteD) = "" Then
        MsgBox "Bitte wählen Sie ein Paar aus!"
        Exit Sub
    End If
    Dim dbs As Database
    Set dbs = CurrentDb ' Bezug auf aktuelle Datenbank zurückgeben.
    
    Dim rstpaare As Recordset
    Set rstpaare = dbs.OpenRecordset("select * from Paare where turniernr = " & get_aktTNr & " and Startkl = '" & BWStartkl & "' Order By Startnr;")
    Dim s_nr As Double
    s_nr = 0
    If Not rstpaare.EOF() Then
       rstpaare.MoveLast
       s_nr = rstpaare!Startnr
    End If
    With rstpaare
            .AddNew
            !Turniernr = Forms![A-Programmübersicht]![Akt_Turnier]
            !Startkl = BWStartkl
            s_nr = s_nr + 1
            !Startnr = s_nr
            !Da_Vorname = BVName_Dame
            !Da_NAchname = BNName_Dame
            !Da_Alterskontrolle = BAlter_Dame
            !He_Vorname = BVName_Herr
            !He_Nachname = BNName_Herr
            !He_Alterskontrolle = BAlter_Herr
            !Verein_nr = gewvnr
            !Verein_Name = AuswahlBW.Column(1)
            !Boogie_Startkarte_D = BSTkarteD
            !Boogie_Startkarte_H = BSTkarteH
            !Anwesent_Status = 1
            !Platz = 0
            !Punkte = 0
            .Update
            End With
 
    bw_startliste.Requery
End Sub

Private Sub Befehl12_Click()
DoCmd.Close
End Sub

Private Sub Befehl59_Click()
If Forms![A-Programmübersicht]![Akt_Turnier] = 0 Then
   MsgBox ("Bitte Turnier auswählen")
   End
End If
Dim stLinkCriteria As String
    DoCmd.OpenForm "RR_Paare_aufnehmen", , , stLinkCriteria
End Sub

Private Sub Befehl77_Click()
If Forms![A-Programmübersicht]![Akt_Turnier] = 0 Then
   MsgBox ("Bitte Turnier auswählen")
   End
End If
Dim stLinkCriteria As String
    DoCmd.OpenForm "BW_Paare_aufnehmen", , , stLinkCriteria
End Sub

Private Sub Befehl78_Click()
If Forms![A-Programmübersicht]![Akt_Turnier] = 0 Then
   MsgBox ("Bitte Turnier auswählen")
   End
End If
Dim stLinkCriteria As String
    DoCmd.OpenForm "TL_BS_aufnehmen", , , stLinkCriteria
End Sub

Private Sub Befehl79_Click()
If Forms![A-Programmübersicht]![Akt_Turnier] = 0 Then
   MsgBox ("Bitte Turnier auswählen")
   End
End If
Dim stLinkCriteria As String
    DoCmd.OpenForm "Formationen_aufnehmen", , , stLinkCriteria
End Sub

Public Sub Befehl34_Click()
    Dim dbs As Database
    Dim rstpaare As Recordset
    Dim sqlstmt As String
    
    Set dbs = CurrentDb
    
    sqlstmt = "select count(*) as vorhanden from Paare where turniernr=" & get_aktTNr & " and Startbuch=" & FBuch & ";"
    Set rstpaare = dbs.OpenRecordset(sqlstmt)
    If Not rstpaare.EOF() Then
       rstpaare.MoveLast
    End If
    
    If (rstpaare!vorhanden > 0) Then
        MsgBox "Die Formation wurde bereits diesem Turnier hinzugefügt!"
        Exit Sub
    End If

    ' Bezug auf aktuelle Datenbank zurückgeben.
    Set dbs = CurrentDb
    If IsNull(FStartklasse) Or IsNull(FBuch) Then
       MsgBox ("Keine Formation ausgewählt!")
       End
    End If
  
    Set rstpaare = dbs.OpenRecordset("select * from Paare where turniernr = " & get_aktTNr & " and Startkl = '" & [FStartklasse] & "' ORDER BY Startnr DESC;")
    Dim s_nr As Double
    s_nr = 0
    If Not rstpaare.EOF() Then
       s_nr = rstpaare!Startnr
    End If
    With rstpaare
            .AddNew
            !Turniernr = Forms![A-Programmübersicht]![Akt_Turnier]
            !Startkl = [FStartklasse]
            s_nr = s_nr + 1
            !Startnr = s_nr
            !Name_Team = formationsname
            !Verein_nr = gewvnr
            !Verein_Name = AuswahlFO.Column(1)
            !Startbuch = FBuch
            !Anwesent_Status = 1
            !Platz = 0
            !Punkte = 0
            .Update
            End With
    Me!Formation_auswahl.Requery

End Sub

Private Sub btn_aktive_1_Click()    ' paare xls-datei laden
    On Error GoTo Err_Befehl80_Click
    Dim Akt_Turnier As Integer
    Dim i As Integer
    Dim dbs As Database
    Dim rstimport, rstpaare As Recordset
    Dim importiert As Integer
    
    Akt_Turnier = [Form_A-Programmübersicht]!Akt_Turnier
    
    If Akt_Turnier = 0 Then
       MsgBox ("Bitte Turnier auswählen")
       End
    End If
    ' Bezug auf aktuelle Datenbank zurückgeben.
    Set dbs = CurrentDb
    
    Call bindExcel(getBaseDir, "PAARE_IMPORT_EXCEL", "Paare_import.xlsx")

    Set rstimport = dbs.OpenRecordset("select * from PAARE_IMPORT_EXCEL")
    If rstimport.EOF() Then
       MsgBox ("Keine Datensätze gefunden!")
       Exit Sub
    End If
    Set rstpaare = dbs.OpenRecordset("select * from Paare where turniernr = " & Akt_Turnier & ";")
       
    importiert = 0
    Do While Not rstimport.EOF()
        If (Not IsNull(rstimport!Startkl) And rstimport!Startkl <> "") Then
            
            With rstpaare
                .AddNew
                !Turniernr = Akt_Turnier
                !Startkl = rstimport!Startkl
                !Startnr = rstimport!Startnr
                If Nz(rstimport!Da_Vorname) <> "" Then !Da_Vorname = left(rstimport!Da_Vorname, 50)
                If Nz(rstimport!Da_NAchname) <> "" Then !Da_NAchname = left(rstimport!Da_NAchname, 50)
                If Nz(rstimport!He_Vorname) <> "" Then !He_Vorname = left(rstimport!He_Vorname, 50)
                If Nz(rstimport!He_Nachname) <> "" Then !He_Nachname = left(rstimport!He_Nachname, 50)
                !Verein_nr = Nz(rstimport!Verein_nr)
                !Verein_Name = left(rstimport!Verein_Name, 50)
                If Nz(rstimport!Name_Team) <> "" Then !Name_Team = Nz(left(rstimport!Name_Team, 50))
                !Startbuch = rstimport!Startbuch
                !Boogie_Startkarte_H = rstimport!Boogie_Startkarte_H
                !Boogie_Startkarte_D = rstimport!Boogie_Startkarte_D
                !Anwesent_Status = 1
                !Platz = 0
                !Punkte = 0
                For i = 1 To 8
                    rstpaare("Akro" & i & "_VR") = rstimport("Akro" & i & "_VR")
                    rstpaare("Wert" & i & "_VR") = rstimport("Wert" & i & "_VR")
                    rstpaare("Akro" & i & "_ZR") = rstimport("Akro" & i & "_VR")
                    rstpaare("Wert" & i & "_ZR") = rstimport("Wert" & i & "_VR")
                    rstpaare("Akro" & i & "_ER") = rstimport("Akro" & i & "_VR")
                    rstpaare("Wert" & i & "_ER") = rstimport("Wert" & i & "_VR")
                Next
                .Update
                importiert = importiert + 1
            End With
        End If
        rstimport.MoveNext
    Loop
    Set rstpaare = dbs.OpenRecordset("SELECT DISTINCT Paare.Startkl FROM Paare WHERE Turniernr =" & Akt_Turnier & ";")
    write_startklassen rstpaare

    MsgBox (importiert & " Paare/Formationen importiert")
Exit_Befehl80_Click:
    Exit Sub

Err_Befehl80_Click:
    MsgBox err.Description
    Resume Exit_Befehl80_Click
    
End Sub

Private Sub btn_aktive_2_Click()    'gemeldete paare vom Server laden
    Dim Akt_Turnier As Integer
    Dim i, cnt As Integer
    Dim dbs As Database
    Dim rstpaare As Recordset
    Dim importiert As Integer
    Dim fName As String
    Dim retl As Long
    
    Akt_Turnier = get_aktTNr
    If Akt_Turnier = 0 Then
       MsgBox ("Bitte Turnier auswählen")
       End
    End If
    Set dbs = CurrentDb
    
    'fName = "T" & Forms![A-Programmübersicht]!Turnier_Nummer & "_TPaare.txt"
    '*****AB***** V14.02 neuer Dateiname für den Paarimport vom Server ab Version 14.02
    cnt = updateTLP(True, False)
    If cnt > 0 Then
        dbs.Execute ("DELETE FROM Akrobatiken;")
        dbs.Execute ("INSERT INTO Akrobatiken SELECT * FROM MSys__Akrobatiken;")
        fName = "T" & Forms![A-Programmübersicht]!Turnier_Nummer & "_Anmeldung.txt"
        retl = get_url_to_file("http://www.drbv.de/cms/images/Download/TurnierProgramm/startlisten/" & fName, getBaseDir() & "Turnierleiterpaket\" & fName)
        
        If retl = 0 Then
            retl = update_drbv_tables("Paare", fName, getBaseDir() & "Turnierleiterpaket\")
            Set rstpaare = dbs.OpenRecordset("SELECT DISTINCT Paare.Startkl FROM Paare WHERE Turniernr =" & Akt_Turnier & ";")
            write_startklassen rstpaare
            Set rstpaare = dbs.OpenRecordset("SELECT Count(0) AS Anz FROM Paare;")
            MsgBox "Es wurden " & rstpaare!anz & " Paare und " & vbCrLf & vbCrLf & (cnt + retl) & " von 7 Dateien importiert.", , "Turnierprogramm"
        Else
            MsgBox " Es wurde keine Datei für dieses Turnier gefunden."
        End If
    Else
        MsgBox "Es wurden keine Daten aktualisiert"
    End If
End Sub

Private Sub btn_aktive_3_Click()
    Dim conf As Integer
    
    If MsgBox("Sie überschreiben alle Akrobatiken" & vbCrLf & "Sicher aktualisieren?", vbYesNo) = vbYes Then
        Dim db As Database
        Set db = CurrentDb
        
        db.Execute ("DELETE FROM Akrobatiken;")
        db.Execute ("INSERT INTO Akrobatiken SELECT * FROM MSys__Akrobatiken;")
    End If

End Sub

Private Sub btn_aktive_4_Click()
    Dim db As Database
    Dim re As Recordset
    Dim fName As String
    Dim retl As Long
    Set db = CurrentDb
     
    If MsgBox("Es werden die gemeldeten Akrobatiken aktualisiert!" & vbCrLf & "Sicher aktualisieren?", vbYesNo) = vbYes Then
        fName = "T" & Forms![A-Programmübersicht]!Turnier_Nummer & "_Anmeldung_2.txt"
        retl = get_url_to_file("http://www.drbv.de/cms/images/Download/TurnierProgramm/startlisten/" & fName, getBaseDir() & "Turnierleiterpaket\" & fName)
        If retl = 0 Then
            retl = update_drbv_tables("Paare", fName, getBaseDir() & "Turnierleiterpaket\")
            Set re = db.OpenRecordset("SELECT Count(0) AS Anz FROM Paare;")
            MsgBox "Es wurden " & re!anz & " Paare aktualisiert.", , "Turnierprogramm"
    
        Else
            MsgBox " Es wurde keine Datei für dieses Turnier gefunden."
        End If
    End If

End Sub

Public Sub btnAddPaar_Click()
    Dim rstpaare As Recordset
    Dim dbs As Database
    Dim sk As String
    Dim sqlstmt As String
    ' Bezug auf aktuelle Datenbank zurückgeben.
    Set dbs = CurrentDb
    If IsNull(gewkl) Then
       MsgBox ("keine Klasse ausgewählt")
       End
    End If
    
    sk = IIf(left([Klassen], 3) = "BS_", "", "RR_") & [Klassen]
    
    sqlstmt = "select count(*) as vorhanden from Paare where turniernr=" & Turnier_Nummer & " and Startbuch=" & STBuchnum & ";"
    Set rstpaare = dbs.OpenRecordset(sqlstmt)
    If Not rstpaare.EOF() Then
       rstpaare.MoveLast
    End If
    
    If (rstpaare!vorhanden > 0) Then
        MsgBox "Das Tanzpaar wurde bereits diesem Turnier hinzugefügt!"
        Exit Sub
    End If
    
    Set rstpaare = dbs.OpenRecordset("select * from Paare where turniernr = " & Turnier_Nummer & " and Startkl = '" & sk & "' order by Startnr;")
    Dim s_nr As Double
    s_nr = 0
    If Not rstpaare.EOF() Then
       rstpaare.MoveLast
       s_nr = rstpaare!Startnr
    End If
    
    With rstpaare
        .AddNew
        !Turniernr = Forms![A-Programmübersicht]![Akt_Turnier]
        !Startkl = sk
        s_nr = s_nr + 1
        !Startnr = s_nr
        !Da_Vorname = VName_Dame
        !Da_NAchname = NName_Dame
        !Da_Alterskontrolle = Alter_Dame
        !He_Vorname = VName_Herr
        !He_Nachname = NName_Herr
        !He_Alterskontrolle = Alter_Herr
        !Verein_nr = gewvnr
        !Verein_Name = AuswahlRR.Column(1)
        !Startbuch = STBuchnum
        !Anwesent_Status = 1
        !Platz = 0
        !Punkte = 0
        .Update
    End With
    Paare_in_Startliste.Requery

End Sub

Private Sub btnDeletePaar_Click()
    Dim res As Integer
    Dim strsql As String
    If (IsNull(Me!STBuchnum)) Then
        Exit Sub
    End If
    
    'Sicherheitsabfrage
    res = MsgBox("Wollen Sie das Paar wirklich löschen?", vbYesNo)
    If (res = vbYes) Then
        Dim dbs As Database
        ' Bezug auf aktuelle Datenbank zurückgeben.
        Set dbs = CurrentDb
        strsql = "delete from paare where Startbuch=" & Me!STBuchnum & " and Turniernr=" & get_aktTNr
        
        dbs.Execute (strsql)
        Me!STBuchnum = ""
        Me!VName_Dame = ""
        Me!NName_Dame = ""
        Me!Alter_Dame = ""
        Me!VName_Herr = ""
        Me!NName_Herr = ""
        Me!Alter_Herr = ""

        Me![Paare in Startliste].Requery
    End If

End Sub

Private Sub btnFormationDelete_Click()
    Dim res As Integer
    Dim strsql As String
    If (IsNull(Me!FBuch)) Then
        Exit Sub
    End If
    
    
    'Sicherheitsabfrage
    res = MsgBox("Wollen Sie die Formation wirklich löschen?", vbYesNo)
    If (res = vbYes) Then
        Dim dbs As Database
        Set dbs = CurrentDb
        strsql = "delete from paare where Startbuch=" & Me!FBuch & " and Turniernr=" & get_aktTNr
        
        dbs.Execute (strsql)
        Me!formationsname = ""
        Me!Clubname_kurz = ""
        Me!FBuch = ""
        Me!FStartklasse = ""
        Me!Formation_auswahl.Requery
    End If

End Sub

Private Sub Form_Open(Cancel As Integer)
    setzte_buttons "Aktive_uebernehmen", "aktive", Forms![A-Programmübersicht]!Turnierausw.Column(8)
End Sub

Private Sub Klassen_Click()
    gewkl = [Klassen]
    [Dame_auswählen].Requery
End Sub

Private Sub SearchName_Change()
    Me!AuswahlRR.Requery
    Me!AuswahlBW.Requery
    Me!AuswahlFO.Requery
    gewvnr = [AuswahlRR]
    If ([AuswahlRR].ListIndex = -1) Then
        gewvnr = -9999
        [AuswahlRR] = -9999
    End If
    
    [Klassen].Requery
    [Dame_auswählen].Requery
    Paare_in_Startliste.Requery

End Sub

Private Sub Seite120_Click()
    Me.SearchName.Visible = False
End Sub

Private Sub Rock_n_Roll_Paare_Click()
    Me.SearchName.Visible = True
End Sub

Private Sub Formationen_Click()
    Me.SearchName.Visible = True
End Sub

Private Sub Boogie_Woogie_Paare_Click()
    Me.SearchName.Visible = True
End Sub

Private Sub Form_Current()
    Call RegisterStr82_Change
End Sub

Private Sub RegisterStr82_Change()
    If Me!RegisterStr82.Value = 0 Then
        Me.SearchName.Visible = False
    Else
        Me.SearchName.Visible = True
    End If
End Sub

Function write_startklassen(rstpaare)
    Dim dbs As Database
    Dim rstimport As Recordset
    Set dbs = CurrentDb
    If rstpaare.RecordCount > 0 Then rstpaare.MoveFirst
    Do Until rstpaare.EOF()
        Set rstimport = dbs.OpenRecordset("SELECT * FROM Startklasse_Turnier WHERE Startklasse ='" & rstpaare!Startkl & "';")
        If rstimport.RecordCount = 0 Then
            rstimport.AddNew
            rstimport!Startklasse = rstpaare!Startkl
            rstimport!Turniernr = get_aktTNr
            rstimport.Update
        End If
        rstpaare.MoveNext
    Loop
End Function
