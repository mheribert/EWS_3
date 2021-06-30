Option Compare Database

    Dim dbs As Database

Private Sub Auswahl_AfterUpdate()
    Dim gewvnr
    gewvnr = Forms!RR_Paare_aufnehmen!auswahl.Column(0)
    Me.Refresh
End Sub

Private Sub Befehl0_Click()
    DoCmd.Close
End Sub

Private Sub Befehl231_Click()
    If Me!Offizielle.Form.RecordsetClone.RecordCount > 0 Then
        DoCmd.OpenReport "Wertungsrichter_Login", acPreview
    End If
End Sub

Public Sub btnAddOffiziellen_Click()
    Dim rstoff, rsCheck As Recordset
    Dim Count, i As Integer
    
    ' Bezug auf aktuelle Datenbank zurückgeben.
    Set dbs = CurrentDb
    
    ' Prüfen, ob der WR schon in der DB vorhanden ist (nur, wenn mit Lizenznr. eingegeben
    If (Not IsNull(Lizenznr) And Lizenznr <> "") Then
        sqlstr = "select count(*) as anzahl from Wert_Richter where turniernr = " & Turnier_Nummer & " and WR_Lizenznr='" & Lizenznr & "';"
        Set rsCheck = dbs.OpenRecordset(sqlstr)
        rsCheck.MoveFirst
        Count = rsCheck!Anzahl
        rsCheck.Close
        
        If (Count > 0) Then
            MsgBox "Dieser Wertungsrichter wurde dem Turnier schon hinzugefügt!"
            Exit Sub
        End If
    Else
        Set rsCheck = dbs.OpenRecordset("Select MAX(WR_Lizenznr) AS maxLiz FROM wert_richter;")
        i = 9000
        If Nz(rsCheck!maxLiz) >= i Then i = rsCheck!maxLiz + 1
        Me!Lizenznr = i
        Me!Club = 0
    End If
    
    If Not IsNull(VName) And Not IsNull(NName) Then
        Set rstoff = dbs.OpenRecordset("select * from Wert_richter where turniernr = " & Turnier_Nummer & " order by wr_kuerzel;")
        Dim ZW_WR As String
        ZW_WR = "@"
        If Not rstoff.EOF() Then
           rstoff.MoveLast
           ZW_WR = rstoff!WR_Kuerzel
        End If
        With rstoff
        .AddNew
        !Turniernr = Forms![A-Programmübersicht]![Akt_Turnier]
        !WR_Lizenznr = Lizenznr
        !WR_Vorname = VName
        !WR_Nachname = NName
        !Vereinsnr = Club
        !WR_Kuerzel = Chr(Asc(ZW_WR) + 1)
        .Update
        End With
        Me!Lizenznr = ""
        Me!VName = ""
        Me!NName = ""
        Me!Club = ""
        
    End If
    Offizielle.Requery
End Sub

Private Sub btnTurnierbericht_Click()
On Error GoTo Err_Befehl51_Click

    Dim stDocName As String

    stDocName = "Wertungsrichter_Einteilung"
    DoCmd.OpenReport stDocName, acPreview

Exit_Befehl51_Click:
    Exit Sub

Err_Befehl51_Click:
    MsgBox err.Description
    Resume Exit_Befehl51_Click
End Sub

Private Sub FilterName_Change()
    off_auswählen.Requery
End Sub

Private Sub FilterNameEingabe_Change()
    FilterName = FilterNameEingabe.text
    off_auswählen.Requery
End Sub

' ***** HM 14.05 *****
' es werden alle Einträge aus Startklasse_Wertungsrichter entfernt wenn ein WR gelöscht wird
Public Sub Form_Close()
    Dim db As Database
    Dim sqlstr As String
    Set db = CurrentDb
    sqlstr = "DELETE * FROM Startklasse_Wertungsrichter WHERE WR_ID NOT IN (SELECT WR_ID FROM Wert_Richter);"
    db.Execute sqlstr
    sqlstr = "DELETE * FROM Startklasse_Wertungsrichter WHERE Startklasse NOT IN (SELECT Startklasse FROM Startklasse_Turnier);"
    db.Execute sqlstr
End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim re As Recordset
    Dim lo As Integer

    Set dbs = CurrentDb
    Set re = dbs.OpenRecordset("SELECT Wert_Richter.WR_Kuerzel, Wert_Richter.WR_ID, Wert_Richter.WR_func, [WR_Nachname] & "" "" & [WR_Vorname] AS Ausdr1 FROM Wert_Richter WHERE (Wert_Richter.Turniernr=" & [Forms]![A-Programmübersicht]![Akt_Turnier] & " AND WR_Azubi = false) ORDER BY Wert_Richter.WR_Kuerzel;")

    If Not re.EOF Then re.MoveFirst
    lo = 1
    Do Until (re.EOF Or lo = 18)
        Me!UForm_wr_liste.Form.Controls("Text" & Format(lo, "0#")).ControlTipText = re!Ausdr1
        Me!UForm_wr_liste.Form.Controls("Name" & Format(lo, "0#")).Caption = re!Ausdr1
        Me!UForm_wr_liste.Form.Controls("Text" & Format(lo, "0#")).Visible = True
        Me!UForm_wr_liste.Form.Controls("Text" & Format(lo, "0#")).ControlSource = "=Sum(iif([WR_" & re!WR_Kuerzel & "]<>"" "",1,0))"
        Me!UForm_wr_liste.Form.Controls("CTRL" & Format(lo, "0#")).ControlTipText = re!Ausdr1
        Me!UForm_wr_liste.Form.Controls("CTRL" & Format(lo, "0#")).Visible = True
        Me!UForm_wr_liste.Form.Controls("CTRL" & Format(lo, "0#")).ControlSource = "WR_" & re!WR_Kuerzel
    
        lo = lo + 1
        re.MoveNext
    Loop
End Sub

Private Sub Form_Resize()
    If Me.WindowHeight > 6600 Then
        Me.RegisterStr65.Height = Me.WindowHeight - 2100
        Me.Offizielle.Height = Me.WindowHeight - 5700
        Me.UForm_wr_liste.Height = Me.WindowHeight - 2800
        Me.Detailbereich.Height = Me.WindowHeight - 200
    End If
End Sub

Private Sub Login_generieren_Click()
    Dim retl As Integer
    Dim wr As Recordset
    If Nz(Me!Offizielle.Form!WR_kenn) <> "" Then
        retl = MsgBox("Es gibt bereits ein Login sollen alle überschieben werden?", vbYesNo)
        If retl = vbNo Then Exit Sub
    End If
    Set wr = Me!Offizielle.Form.RecordsetClone
    For retl = 1 To 23
        rnd
    Next
    wr.MoveFirst
    Do Until wr.EOF
        retl = Int((9999 * rnd) + 1)
        wr.Edit
        wr!WR_kenn = Format(retl, "0000")
        wr.Update
        wr.MoveNext
    Loop
    DoCmd.Requery
End Sub

Private Sub RegisterStr65_Change()
    Dim lo As Integer
    If Me!Offizielle.Form.RecordsetClone.RecordCount > 0 Then
        If Me!UForm_wr_liste.Form.RecordsetClone.RecordCount > 0 Then
            If Me!RegisterStr65.Value = 1 Then
                Me!UForm_wr_liste.Form.CTRL01.SetFocus
                For lo = 2 To 16
                    Me!UForm_wr_liste.Form.Controls("Text" & Format(lo, "0#")).Visible = False
                    Me!UForm_wr_liste.Form.Controls("CTRL" & Format(lo, "0#")).Visible = False
                    Me!UForm_wr_liste.Form.Controls("Name" & Format(lo, "0#")).Caption = ""
                Next lo
            
                Call Form_Open(1)
            End If
        Else
            If Me!RegisterStr65.Value = 1 Then
                MsgBox "Es wurden noch keine Startklassen definiert!", vbOKOnly
            End If
        End If
    Else
        If Me!RegisterStr65.Value = 1 Then
            MsgBox "Es wurden noch keine Wertungsrichter eingegeben!", vbOKOnly
        End If
    End If
    Me.Requery

End Sub

Function Einteil()
    Dim sqlcmd As String
    Dim sel As String
    
    Set dbs = CurrentDb
    
    sel = Screen.ActiveControl.Name
    If Screen.ActiveControl = "X" Then
        Screen.ActiveControl = ""
        sqlcmd = "delete from Startklasse_wertungsrichter skwr where (skwr.wr_id=" & Me.Controls("W" & left(sel, 1)).ControlTipText & " and skwr.startklasse=""" & Me.Controls("Klasse" & Format(Mid(sel, 2, 2), "#0")).ControlTipText & """);"
    Else
        Screen.ActiveControl = "X"
        sqlcmd = "insert into Startklasse_wertungsrichter( WR_ID, startklasse)"
        sqlcmd = sqlcmd & " values(" & Me.Controls("W" & left(sel, 1)).ControlTipText & ", """ & Me.Controls("Klasse" & Format(Mid(sel, 2, 2), "#0")).ControlTipText & """);"
    End If
    dbs.Execute (sqlcmd)
End Function

Private Sub Wertungsrichterdeckblatt_Click()
    If Me!druck Then
        DoCmd.OpenReport "Deckblatt_quer", acViewPreview
    Else
        DoCmd.OpenReport "Deckblatt", acViewPreview
    End If
End Sub

Private Sub EMail_Click()
On Error GoTo EMail_noSend
    Dim wr, re As Recordset
    Dim MailAn As String
    Dim body As String
    Set dbs = CurrentDb
    Set re = Forms!Wertungsrichter_aufnehmen!Offizielle.Form.RecordsetClone
    re.MoveFirst
    Do Until re.EOF
        Set wr = dbs.OpenRecordset("SELECT * FROM TLP_OFFIZIELLE WHERE Lizenzn = '" & re!WR_Lizenznr & "';")
        If wr.RecordCount > 0 Then
            If wr![e-mail] <> "" Then
                MailAn = MailAn & wr![e-mail] & "; "
            End If
        End If
        re.MoveNext
    Loop
    MailAn = left(MailAn, Len(MailAn) - 2)
    body = "Liebe Wertungsrichter," & vbCrLf & vbCrLf & "am " & DLookup("T_Datum", "Turnier", "Turniernum =1") & " findet der " & _
           DLookup("Turnier_Name", "Turnier", "Turniernum =1") & " statt."
    DoCmd.SendObject , , , , , MailAn, Forms![A-Programmübersicht]!Turnierbez, body, True
    Exit Sub
    
EMail_noSend:
    If err.Number <> 2501 Then MsgBox "Error: " & err
    
End Sub

