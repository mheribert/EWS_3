Option Compare Database

    Dim dbs As Database

'Private Sub Auswahl_AfterUpdate()
'    Dim gewvnr
'    gewvnr = Forms!RR_Paare_aufnehmen!auswahl.Column(0)
'    Me.Refresh
'End Sub
'
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
    Set dbs = CurrentDb
    
    ' Pr�fen, ob der WR schon in der DB vorhanden ist (nur, wenn mit Lizenznr. eingegeben
    If (Not IsNull(Lizenznr) And Lizenznr <> "") Then
        sqlstr = "select count(*) as anzahl from Wert_Richter where turniernr = " & Turnier_Nummer & " and WR_Lizenznr='" & Lizenznr & "';"
        Set rsCheck = dbs.OpenRecordset(sqlstr)
        rsCheck.MoveFirst
        Count = rsCheck!Anzahl
        rsCheck.Close
        
        If (Count > 0) Then
            MsgBox "Dieser Wertungsrichter wurde dem Turnier schon hinzugef�gt!"
            Exit Sub
        End If
    Else
        Set rsCheck = dbs.OpenRecordset("Select MAX(WR_Lizenznr) AS maxLiz FROM wert_richter;")
        i = 9000
        If Nz(rsCheck!maxLiz) >= i Then i = rsCheck!maxLiz + 1
        Me!Lizenznr = i
        Me!Club = 0
    End If
    
    ' Pr�fen ob WR-Anzahl max erreicht.
    If DCount("WR_Kuerzel", "Wert_Richter") > get_properties("max_WR") - 1 Then
        MsgBox "Die maximale Anzahl der Wertungsrichter ist erreicht!"
        Exit Sub
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
        !Turniernr = Forms![A-Programm�bersicht]![Akt_Turnier]
        !WR_Lizenznr = Lizenznr
        !WR_Vorname = VName
        !WR_Nachname = NName
        !Vereinsnr = IIf(Club = "", 0, Club)
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
    off_ausw�hlen.Requery
End Sub

Private Sub FilterNameEingabe_Change()
    FilterName = FilterNameEingabe.text
    off_ausw�hlen.Requery
End Sub

' ***** HM 14.05 *****
' es werden alle Eintr�ge aus Startklasse_Wertungsrichter entfernt wenn ein WR gel�scht wird
Public Sub Form_Close()
    Dim db As Database
    Dim re As Recordset
    Dim sqlstr As String
    Dim stkl As String
    Set db = CurrentDb
    sqlstr = "DELETE * FROM Startklasse_Wertungsrichter WHERE WR_ID NOT IN (SELECT WR_ID FROM Wert_Richter);"
    db.Execute sqlstr
    sqlstr = "DELETE * FROM Startklasse_Wertungsrichter WHERE Startklasse NOT IN (SELECT Startklasse FROM Startklasse_Turnier);"
    db.Execute sqlstr
    Set re = db.OpenRecordset("SELECT Count(Wert_Richter.WR_ID) AS AnzWR, Startklasse FROM Wert_Richter INNER JOIN Startklasse_Wertungsrichter ON Wert_Richter.WR_ID = Startklasse_Wertungsrichter.WR_ID WHERE (WR_Azubi=False AND WR_Function <> 'Ob' AND WR_Function<>'MA' AND WR_Function<>'MB') GROUP BY Startklasse;")
    If re.RecordCount > 0 Then re.MoveFirst
    Do Until re.EOF
        If re!anzWR > 8 Then
            stkl = stkl & re!Startklasse & vbCrLf
        End If
        re.MoveNext
    Loop
    If Len(stkl) > 0 Then
        MsgBox "Die Anzahl der aktiven Wertungsrichter bei " & vbCrLf & stkl & "�berschreitet die maximale Anzahl des TLP."
    End If
    ' Sperre Observer Probewertungsrichter
    Set re = db.OpenRecordset("SELECT Startklasse_Wertungsrichter.WR_ID FROM Wert_Richter INNER JOIN Startklasse_Wertungsrichter ON Wert_Richter.WR_ID = Startklasse_Wertungsrichter.WR_ID WHERE WR_function='Ob' AND WR_Azubi=True;")
    If re.RecordCount > 0 Then
        re.MoveFirst
        Do Until re.EOF
            sqlstr = "DELETE * FROM Startklasse_Wertungsrichter WHERE WR_ID = " & re!WR_ID
            db.Execute sqlstr
            re.MoveNext
        Loop
        MsgBox "Ein Probe/Schattenwertungsrichter darf kein Observer sein!" & vbCrLf & "Diese Einteilung wurde gel�scht!"
    End If
    ' Check ob zuviele MK-WR
    Set re = dbs.OpenRecordset("SELECT Startklasse FROM Startklasse_Wertungsrichter WHERE WR_function='MA' GROUP BY Startklasse HAVING Count(WR_function)>1;")
    If Not re.EOF Then _
        MsgBox "Achtung!" & vbCrLf & "zu viele MA-Wertungsrichter " & re!Startklasse
    Set re = dbs.OpenRecordset("SELECT Startklasse FROM Startklasse_Wertungsrichter WHERE WR_function='MB' GROUP BY Startklasse HAVING Count(WR_function)>3;")
    If Not re.EOF Then _
        MsgBox "Achtung!" & vbCrLf & "zu viele MB-Wertungsrichter bei " & re!Startklasse
    ' Check ob alle MK_WR eingetragen
    If get_mk() <> "" Then
        Set re = db.OpenRecordset("SELECT Startklasse_Turnier.Startklasse, (SELECT Count(WR_ID) from Startklasse_Wertungsrichter where WR_function LIKE 'M*' and Startklasse_Wertungsrichter.Startklasse = Startklasse_Turnier.Startklasse ) AS Mx, (SELECT Count(TP_ID) from Paare where Paare.Startkl = Startklasse_Turnier.Startklasse and Paare.Anwesent_Status = 1 ) AS Pa FROM Startklasse_Turnier;")
        st_kl = "RR_J RR_S RR_S1 RR_S2"
        If re.RecordCount > 0 Then re.MoveFirst
        Do Until re.EOF()
            If InStr(st_kl, re!Startklasse) > 0 Then
                If re!Mx = 0 And re!pa > 0 Then
                    MsgBox "Bei " & re!Startklasse & " gibt es keine MK-WR!"
                End If
            End If
            re.MoveNext
        Loop
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim re As Recordset
    Dim lo As Integer
    Dim maxWR As Integer

    Set dbs = CurrentDb
    Set re = dbs.OpenRecordset("SELECT Wert_Richter.WR_Kuerzel, Wert_Richter.WR_ID, Wert_Richter.WR_func, WR_Azubi, [WR_Nachname] & "" "" & [WR_Vorname] AS Ausdr1 FROM Wert_Richter WHERE (Wert_Richter.Turniernr=" & get_aktTNr & ") ORDER BY Wert_Richter.WR_Kuerzel;")

    If Not re.EOF Then re.MoveFirst
    lo = 1
    maxWR = get_properties("max_WR")
    Do Until (re.EOF Or (lo = maxWR + 1))
        If re!WR_AzuBi = True Then
            Me!UForm_wr_liste.Form.Controls("Name" & Format(lo, "0#")).BackStyle = 1
        Else
            Me!UForm_wr_liste.Form.Controls("Name" & Format(lo, "0#")).BackStyle = 0
        End If
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
        Me.RegisterStr65.Height = Me.WindowHeight - 1600
        Me.Offizielle.Height = Me.WindowHeight - 5250
        Me.UForm_wr_liste.Height = Me.WindowHeight - 2800
        Me.Detailbereich.Height = Me.WindowHeight - 200
    End If
End Sub

Private Sub Login_generieren_Click()
    Dim retl As Integer
    Dim wr As Recordset
    If Nz(Me!Offizielle.Form!WR_kenn) <> "" Then
        retl = MsgBox("Es gibt bereits ein Login sollen alle �berschieben werden?", vbYesNo)
        If retl = vbNo Then Exit Sub
    End If
    Set wr = Me!Offizielle.Form.RecordsetClone
    wr.MoveFirst
    Do Until wr.EOF
        Randomize
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
                For lo = 2 To get_properties("max_WR")
                    Me!UForm_wr_liste.Form.Controls("Text" & Format(lo, "0#")).Visible = False
                    Me!UForm_wr_liste.Form.Controls("CTRL" & Format(lo, "0#")).Visible = False
                    Me!UForm_wr_liste.Form.Controls("Name" & Format(lo, "0#")).Caption = ""
                    Me!UForm_wr_liste.Form.Controls("Name" & Format(lo, "0#")).BackStyle = 0
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
        Set wr = dbs.OpenRecordset("SELECT * FROM TLP_OFFIZIELLE WHERE Lizenzn = '" & re!WR_Lizenznr & "' AND WName ='" & re!WR_Nachname & "';")
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
    DoCmd.SendObject , , , , , MailAn, Forms![A-Programm�bersicht]!Turnierbez, body, True
    Exit Sub
    
EMail_noSend:
    If err.Number <> 2501 Then MsgBox "Error: " & err
    
End Sub

