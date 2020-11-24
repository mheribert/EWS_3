Option Compare Database
Option Explicit
Public akt_st As String

Private Sub Akro_anzeigen_Click()

    If Not IsNull(Me!TP_ID) Then DoCmd.OpenForm "Paare_Akrobatiken", , , "TP_ID = " & Me!TP_ID

End Sub

Private Sub Befehl12_Click()
On Error GoTo Err_Befehl12_Click


    DoCmd.Close

Exit_Befehl12_Click:
    Exit Sub

Err_Befehl12_Click:
    MsgBox err.Description
    Resume Exit_Befehl12_Click
    
End Sub

Private Sub Befehl65_Click()
    
    If Not IsNull(Me!TP_ID) Then DoCmd.OpenReport "Giveaway", acViewPreview, , "TP_ID = " & Me!TP_ID
    
End Sub

Private Sub btnAktualisieren_Click()
    Requery
End Sub

Private Sub Da_NAchname_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub Da_Vorname_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub FilterStartklasse_DblClick(Cancel As Integer)
    Me!FilterStartklasse = -1
    FilterStartklasse_Change
End Sub

Private Sub Form_Load()
    If get_properties("EWS") = "EWS3" Then
        Me!moderator_vorstellung.Visible = True
    Else
        Me!moderator_vorstellung.Visible = False
    End If
    Const xoff = 630
    Select Case Forms![A-Programmübersicht]!Turnierausw.Column(8)
        Case "SL"
            Me.DA_Vorname1.left = Me.Da_Vorname.left - xoff
            Me.Da_Nachname1.left = Me.Da_NAchname.left - xoff
            Me.He_Vorname1.left = Me.He_Vorname.left - xoff
            Me.He_Nachname1.left = Me.He_Nachname.left - xoff
            Me.Da_Vorname.left = Me.Da_Vorname.left - xoff
            Me.Da_NAchname.left = Me.Da_NAchname.left - xoff
            Me.He_Vorname.left = Me.He_Vorname.left - xoff
            Me.He_Nachname.left = Me.He_Nachname.left - xoff
            Me!Tanz.Visible = True
            Me!Tanz1.Visible = True
            Me.Da_Alterskontrolle.Visible = False
            Me.DA_Alterskontrolle1.Visible = False
            Me!Boogie_Startkarte_D.Visible = False
            Me!Boogie_Startkarte_D1.Visible = False
            Me.He_Alterskontrolle.Visible = False
            Me!Boogie_Startkarte_H.Visible = False
            Me!Boogie_Startkarte_H1.Visible = False
            Me!Startbuch.Visible = False
            Me!Startbuch1.Visible = False
            Me.Wertungen_ausdrucken.Visible = False
        Case "BW"
            Me.Wertungen_ausdrucken.Visible = False
        Case "BY"
            Me.Wertungen_ausdrucken.Visible = False
            Me.Akro_anzeigen.Visible = False
        Case Else
    End Select

End Sub

Private Sub He_Nachname_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub He_Vorname_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub Kombinationsfeld36_DblClick(Cancel As Integer)
    Me!FilterStartklasse = Me!Kombinationsfeld36
    FilterStartklasse_Change
End Sub

Private Sub Kombinationsfeld36_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub moderator_vorstellung_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim st As String
'    If (IsNull([FilterStartklasse]) Or [FilterStartklasse] = -1) Then
        st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=moderator_vorstellung&text=0")
'    Else
'        st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=moderator_vorstellung&text=" & Me!FilterStartklasse)
'    End If
End Sub

Private Sub Paar_Status_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub Startbuch_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub Startnr_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub Verein_Name_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub Verein_nr_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub FilterStartklasse_Change()
    'MsgBox "Startklasse = " & [FilterStartklasse]
    If (IsNull([FilterStartklasse]) Or [FilterStartklasse] = -1) Then
        Me.Filter = ""
        Startnummernvergabe.Enabled = False
        Me.FilterOn = False
    Else
        Me.Filter = "Startkl = '" & [FilterStartklasse] & "'"
        Startnummernvergabe.Enabled = True
        Me.FilterOn = True
    End If
    'Me.Refresh
    
End Sub

Private Sub Form_Activate()
    Call FilterStartklasse_Change
End Sub

Private Sub Form_SelectionChange()
akt_st = Anwesent_Status
End Sub

Private Sub Liste25_AfterUpdate()
Dim dbs As Database
Dim rstauswertung, rststartnr As Recordset
' Bezug auf aktuelle Datenbank zurückgeben.
Set dbs = CurrentDb
' Paare Rundenqualifikation zuordnen und FIlter auf turniernummer setzen
Set rstauswertung = dbs.OpenRecordset("select * from Paare_Rundenqualifikation where turniernr = " & Turniernr & ";")
If rstauswertung.EOF() Then
   End
End If
' Status = entschuldigt, Wenn das Paar in der Tabelle Rundenqualifikation vorhanden ist wird dieses nun darin gelöscht.
If Anwesent_Status = 0 Then
    rstauswertung.FindFirst ("Startnummer = " & Startnr & " And startklass = '" & Startkl & "' and tanzrund = 'Vor_r'")
    If rstauswertung.NoMatch Then
        End
    End If
    With rstauswertung
        .Delete
        MsgBox ("Das Paar mit der Startnummer " & Startnr & " aus der Startklasse " & Startkl & " wurde aus der Rundenqualifikation gelöscht")
    End With
    End
End If
If Anwesent_Status = 1 Then
    rstauswertung.FindFirst ("Startnummer = " & Startnr & " And startklass = '" & Startkl & "' and (tanzrund = 'Vor_r' or tanzrund = 'End_r')")
    If rstauswertung.NoMatch Then
        End
    End If
    With rstauswertung
        .Edit
        !Anwesend = 1
        .Update
        MsgBox ("Das Paar mit der Startnummer " & Startnr & " aus der Startklasse " & Startkl & " wurde in der Rundenqualifikation auf ANWESEND gesetzt")
    End With
    Set rstauswertung = dbs.OpenRecordset("select * from auswertung where turniernr = " & Turniernr & " and startkl = '" & Startkl & "';")
    If Not rstauswertung.EOF() Then
        ' Anfang unentschuldigte Paare nach dem eintreffen noch in die Auswertung der WR anfügen HK 02.06.04
        rstauswertung.FindFirst ("Startnr = " & Startnr & " And startkl = '" & Startkl & "' and (t_runde = 'Vor_r' or t_runde = 'end_r')")
        If rstauswertung.NoMatch Then
           rstauswertung.Sort = "wert_ken"
           rstauswertung.MoveFirst
           Dim WR_K As String, akt_r As String
           akt_r = rstauswertung!T_Runde
           Set rststartnr = dbs.OpenRecordset("auswertung")
           Do While Not rstauswertung.NoMatch
                WR_K = rstauswertung!Wert_Ken
                rststartnr.AddNew
                rststartnr!Wert_Ken = WR_K
                rststartnr!Startnr = Startnr
                rststartnr!T_Runde = akt_r
                rststartnr!Turniernr = Turniernr
                rststartnr!Startkl = Startkl
                rststartnr!Punkte = 0
                rststartnr!Platz = 0
                rststartnr!Reihenfolge = 9999
                rststartnr.Update
                MsgBox ("Das Paar " & Startnr & " aus der Startklasse " & Startkl & " wurde für den WR " & WR_K & ", an die bereits begonnene Eingabe der Wertungen, angefügt")
                rstauswertung.FindNext ("wert_ken <> '" & WR_K & "' and startkl = '" & Startkl & "' and t_runde = '" & akt_r & "'")
                Loop
        End If
        ' ende 02.06.04
    End If
    End
End If
If Anwesent_Status = 2 Then
    rstauswertung.FindFirst ("Startnummer = " & Startnr & " And startklass = '" & Startkl & "' and (tanzrund = 'Vor_r' or tanzrund = 'end_r')")
    If rstauswertung.NoMatch Then
        End
    End If
    With rstauswertung
        .Edit
        !Anwesend = 2
        .Update
        MsgBox ("Das Paar mit der Startnummer " & Startnr & " aus der Startklasse " & Startkl & " wurde in der Rundenqualifikation auf UNENTSCHULDIGT gesetzt")
    End With
    End
End If
End Sub

Private Sub Text29_Dirty(Cancel As Integer)
    akt_st = Anwesent_Status
End Sub

Private Sub Paar_Status_AfterUpdate()
    If (Not hasWertungen(TP_ID)) Then
        Dim dbs As Database
        Set dbs = CurrentDb
        Dim rst As Recordset
        Dim stmt As String
        stmt = "Select * from Paare_Rundenqualifikation pr where tp_id=" & TP_ID
        Set rst = dbs.OpenRecordset(stmt)
        Do While (Not rst.EOF)
            rst.Edit
            rst!Anwesend_Status = Anwesent_Status
            rst.Update
            rst.MoveNext
        Loop
        rst.Close
    End If
End Sub


Private Sub Startnummernvergabe_Click()
    
    Dim dbs As Database
    Dim rstpaare As Recordset
    ' Bezug auf aktuelle Datenbank zurückgeben.
    Set dbs = CurrentDb
    
    Dim firstNummer As Integer
    firstNummer = 1
    
    ' Bisherige erste Startnummer ermitteln
    Dim sqlString As String
    
    sqlString = "select * from Paare where turniernr = " & Turniernr & " and Startkl = '" & [FilterStartklasse] & "' order by startnr;"
    Set rstpaare = dbs.OpenRecordset(sqlString)
    Dim s_nr As Double
    If Not rstpaare.EOF() Then
      firstNummer = Nz(rstpaare!Startnr)
    End If
    rstpaare.Close
    
    ' Maximale alte Startnummer ermitteln
    sqlString = "select max(Startnr) as maxStartnr from Paare where turniernr = " & Turniernr & " and Startkl = '" & [FilterStartklasse] & "'"
    Set rstpaare = dbs.OpenRecordset(sqlString)
    Dim maxStartnr As Double
    If Not rstpaare.EOF() Then
      maxStartnr = Nz(rstpaare!maxStartnr)
    End If
    rstpaare.Close
    
    ' Anzahl der Paare in dieser Startklasse ermitteln
    sqlString = "select count(*) as Anzahl from Paare where turniernr = " & Turniernr & " and Startkl = '" & [FilterStartklasse] & "'"
    Set rstpaare = dbs.OpenRecordset(sqlString)
    Dim countPaare As Double
    If Not rstpaare.EOF() Then
      countPaare = rstpaare!Anzahl
    End If
    rstpaare.Close
    
    ' Startnummer über Dialog abfragen
    Dim benutzereingabe As String
    benutzereingabe = InputBox("Bitte geben Sie die erste Startnummer für die " & [FilterStartklasse].Column(1) & " ein:", "Startnummernvergabe", firstNummer)
    
    If (benutzereingabe = "") Then
        Exit Sub
    End If
    
    If (Not IsNumeric(benutzereingabe)) Then
        MsgBox "Bitte geben Sie eine Nummer ein!"
        Exit Sub
    End If
    
    ' Dummymäßig die Startnummern erstmal nach 10000 verlegen damit sich anschließend die Nummern nicht überschneiden
    firstNummer = 10000
 
    ' Startklasse jetzt mit der neuen Nummer durchnummerieren
        
    sqlString = "select * from Paare where turniernr = " & Turniernr & " and Startkl = '" & [FilterStartklasse] & "' order by startnr;"
    
    Set rstpaare = dbs.OpenRecordset(sqlString)
    
    While Not rstpaare.EOF()
        With rstpaare
          .Edit
          
          !Startnr = firstNummer
          firstNummer = firstNummer + 1
          .Update
        End With
        
        rstpaare.MoveNext
    Wend
    
    rstpaare.Close
    
    ' Jetzt die richtige Neuvergabe der Startnummern durchführen
    firstNummer = benutzereingabe
    Set rstpaare = dbs.OpenRecordset(sqlString)
    
    While Not rstpaare.EOF()
        With rstpaare
          .Edit
          
          !Startnr = firstNummer
          firstNummer = firstNummer + 1
          .Update
        End With
        
        rstpaare.MoveNext
    Wend
    
    rstpaare.Close
    
    Me.Refresh
    
End Sub

Private Sub Wertungen_ausdrucken_Click()
    If Not IsNull(Me!TP_ID) Then DoCmd.OpenReport "Giveaway", acViewPreview, , "TP_ID = " & Me!TP_ID
End Sub
