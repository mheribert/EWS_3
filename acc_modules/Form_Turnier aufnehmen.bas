Option Compare Database
Option Explicit

Private Sub Befehl8_Click()
On Error GoTo Err_Befehl8_Click

    DoCmd.Close

Exit_Befehl8_Click:
    Exit Sub

Err_Befehl8_Click:
    MsgBox err.Description
    Resume Exit_Befehl8_Click
    
End Sub

Private Sub Befehl9_Click()
On Error GoTo Err_Befehl9_Click


    DoCmd.Close

Exit_Befehl9_Click:
    Exit Sub

Err_Befehl9_Click:
    MsgBox err.Description
    Resume Exit_Befehl9_Click
    
End Sub

Private Sub btnTurnierbericht_Click()
On Error GoTo Err_btnTurnierbericht_Click

    [Form_A-Programmübersicht]![Report_Turniernum] = Turniernum
    Dim stDocName As String
    stDocName = "Turnierbericht"
    DoCmd.OpenReport stDocName, acPreview

Exit_btnTurnierbericht_Click:
    Exit Sub

Err_btnTurnierbericht_Click:
    MsgBox err.Description
    Resume Exit_btnTurnierbericht_Click
    
End Sub

Private Sub btnTurnieruebernahme_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Turnier_uebernehmen"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , acDialog
End Sub

Private Sub Form_Close()
    Dim re As Recordset
    Dim vars
    Dim i, anzWR As Integer
    Set re = Forms![Turnier aufnehmen]![Startklasse_Turnier Unterformular].Form.RecordsetClone
    If re.RecordCount <> 0 Then re.MoveFirst
    Do Until re.EOF
        anzWR = 0
        vars = Split(Nz(re!SelectWR), "+")
        For i = 0 To UBound(vars)
            anzWR = anzWR + vars(i)
        Next
        If anzWR <> re!AnzahlWR Then
            MsgBox "Bei " & re!Startklasse_text & " stimmt die Anzahl der Wertungsrichter nicht!" & vbCrLf & "Bitte neu eingeben!"
        End If
        re.MoveNext
    Loop
    
End Sub

Private Sub Form_Open(Cancel As Integer)
    If (Not IsNull([Form_A-Programmübersicht]![Akt_Turnier]) And [Form_A-Programmübersicht]![Akt_Turnier] <> 0 And [Form_A-Programmübersicht]![Akt_Turnier] <> "") Then
        Me.RecordsetClone.FindFirst "Turniernum=" & [Form_A-Programmübersicht]![Akt_Turnier]
        Me.Bookmark = Me.RecordsetClone.Bookmark
    End If
End Sub

Private Sub Form_Resize()
    If Me.InsideHeight > 7000 Then
        Me![Startklasse_Turnier Unterformular].Height = Me.InsideHeight - 6000
        Me![besondere_Vorkommnisse].Height = Me.InsideHeight - 6000
        Me.ScrollBars = 0
    Else
        Me.ScrollBars = 2
    End If
End Sub

Sub Kombinationsfeld35_AfterUpdate()
    ' Den mit dem Steuerelement übereinstimmenden Datensatz suchen.
    Me.RecordsetClone.FindFirst "Turniernum=" & Me![Kombinationsfeld35]
    Me.Bookmark = Me.RecordsetClone.Bookmark
End Sub

Private Sub TurnierAnlegen_Click()
On Error GoTo Err_TurnierAnlegen_Click

    Dim sqlstr As String
    sqlstr = "INSERT INTO TURNIER(TURNIER_NAME) VALUES ('<Name des neuen Turniers>')"
    Dim dbs As Database
    Set dbs = CurrentDb   ' Bezug auf aktuelle Datenbank zurückgeben
    dbs.Execute (sqlstr)
    
    Requery
    Kombinationsfeld35.Requery
    DoCmd.GoToRecord , , acLast
    
Exit_TurnierAnlegen_Click:
    Exit Sub

Err_TurnierAnlegen_Click:
    MsgBox err.Description
    Resume Exit_TurnierAnlegen_Click
    
End Sub

Private Sub TurnierLoeschen_Click()
On Error GoTo Err_TurnierLoeschen_Click

    ' Abbruch, falls kein Turnier ausgewählt wurde
    If (IsNull(Turnier_Nummer) Or Turnier_Nummer = "" Or Not IsNumeric(Turnier_Nummer)) Then
        MsgBox "Sie haben kein Turnier zum Löschen ausgewählt!"
        Exit Sub
    End If
    
    Dim Turniername As String
    Turniername = Turnier_Name
    
    Dim eingabe As String
    eingabe = InputBox("Bitten bestätigen Sie das Löschen des Turniers " & Chr(13) & "'" & Turniername & "'" & Chr(13) & "durch die Eingabe der Turniernummer:", "Turnier löschen")
    
    If (eingabe = "") Then
        Exit Sub
    End If
    
    If (Not IsNumeric(eingabe)) Then
        MsgBox "Die eingegebene Turniernummer ist ungültig!"
        Exit Sub
    End If
    
    ' Abbruch, falls die Turniernummer falsch ist
    If (Turnier_Nummer <> eingabe) Then
        MsgBox "Die eingegebene Turniernummer ist falsch"
        Exit Sub
    End If

    eingabe = MsgBox("Wollen Sie das ausgewählte Turnier wirklich löschen?", vbYesNo)
    
    If (eingabe = vbNo) Then
        Exit Sub
    End If
    
    Dim dbs As Database
    Set dbs = CurrentDb   ' Bezug auf aktuelle Datenbank zurückgeben
    
    Dim Turniernr As Integer
    Turniernr = Turniernum
    
    
    dbs.Execute ("DELETE FROM Turnierleitung WHERE Turniernr=" & Turniernr)
    dbs.Execute ("DELETE FROM Wert_richter WHERE Turniernr=" & Turniernr)
    dbs.Execute ("DELETE FROM Anzahl_Paare WHERE Turniernr=" & Turniernr)
    dbs.Execute ("DELETE FROM Rundentab WHERE Turniernr=" & Turniernr)
    dbs.Execute ("DELETE FROM Paare WHERE Turniernr=" & Turniernr)
    dbs.Execute ("DELETE FROM Turnier WHERE Turniernum=" & Turniernr)
    
    MsgBox "Das Turnier '" & Turniername & "' wurde gelöscht!"
    
    Requery
    Kombinationsfeld35.Requery
    
    If (Kombinationsfeld35.ListCount > 0) Then
        DoCmd.GoToRecord , , acFirst
    End If
Exit_TurnierLoeschen_Click:
    Exit Sub

Err_TurnierLoeschen_Click:
    MsgBox err.Description
    Resume Exit_TurnierLoeschen_Click
    
End Sub
