Option Compare Database
Option Explicit

Private Sub Befehl12_Click()
On Error GoTo Err_Befehl12_Click


    DoCmd.Close

Exit_Befehl12_Click:
    Exit Sub

Err_Befehl12_Click:
    MsgBox err.Description
    Resume Exit_Befehl12_Click
    
End Sub

Private Sub refresh_Startklassen()
    Form_Wertungsrichter_aufnehmen!currentWR_ID = WR_ID
End Sub

Private Sub Form_AfterDelConfirm(Status As Integer)
    If Status = acDeleteOK Then
        Forms!Wertungsrichter_aufnehmen.Form_Close
    End If
    
End Sub

Private Sub Form_Close()
    Dim re As Recordset
    Dim i As Long
    
    Set re = DBEngine(0)(0).OpenRecordset("Select MAX(WR_Lizenznr) AS maxLiz FROM wert_richter;")
    i = 9000
    If re!maxLiz > i Then i = re!maxLiz
    Set re = Me.RecordsetClone
    If re.RecordCount > 0 Then
        re.MoveFirst
        Do Until re.EOF
            If Nz(re!WR_Lizenznr) = "" Then
                re.Edit
                re!WR_Lizenznr = i
                re.Update
                i = i + 1
            End If
            re.MoveNext
        Loop
    End If
End Sub

Private Sub km_holen_Click()
    Dim db As Database
    Dim wr As Recordset
    Dim re As Recordset
    Dim objIE As Object
    Dim ti, s
    Set db = CurrentDb
    Set wr = db.OpenRecordset("SELECT * FROM TLP_OFFIZIELLE WHERE Lizenzn=""" & Me!Lizenznr & """;")
    Set re = db.OpenRecordset("SELECT * FROM turnier WHERE turniernum=" & get_aktTNr & ";")
    If wr.RecordCount > 0 Then
        Set objIE = CreateObject("InternetExplorer.Application")
        objIE.Navigate2 "https://www.google.de/maps/dir/" & wr!straße & ", " & wr!plz & " " & wr!ort & "/" & re!Veranst_Ort & Chr(13) & Chr(10)   '"About:blank"
        objIE.Visible = True
        
        'ti = Time
        'Do Until ti + 0.00004 < Time
        
        'Loop
        ' 1 str
        ' 2 plz 3 ort
        ' 5 ziel
        'Debug.Print "from: " & wr!straße & ", " & wr!plz & " " & wr!ort & " to: " & Me!Liste1.Column(5) & Chr(13)
        'objIE.Document.Forms.Item(0).elements("q").value = "from: " & wr!straße & ", " & wr!plz & " " & wr!ort & " to: " & Forms![A-Programmübersicht]!Turnierauswahl.Column(6) & Chr(13) & Chr(10)
    End If
End Sub

Private Sub Lizenznr_GotFocus()
    Call refresh_Startklassen
End Sub

Private Sub TL_Nachname_GotFocus()
    Call refresh_Startklassen
End Sub

Private Sub TL_Vorname_GotFocus()
    Call refresh_Startklassen
End Sub

Private Sub Vereinsnr_GotFocus()
    Call refresh_Startklassen
End Sub

Private Sub Wertungsbögen_drucken_Click()
    Dim dbs As Database
    Dim re As Recordset
    Dim WB As String
    Dim anz As Integer
    If MsgBox("Wertungsbögen für " & Me!WR_Nachname & " " & WR_Vorname & " auf" & Chr(13) & Chr(13) & _
                Me.ActiveControl.Application.Printer.DeviceName & " drucken?", vbYesNo) = 6 Then
        Set dbs = CurrentDb
        Set re = dbs.OpenRecordset("SELECT Wert_Richter.Turniernr, Wert_Richter.WR_Kuerzel, Rundentab.Startklasse, Rundentab.Runde, Rundentab.WB FROM (Wert_Richter INNER JOIN Startklasse_Wertungsrichter ON Wert_Richter.WR_ID = Startklasse_Wertungsrichter.WR_ID) INNER JOIN Rundentab ON Startklasse_Wertungsrichter.Startklasse = Rundentab.Startklasse WHERE (((Wert_Richter.Turniernr)=" & [Forms]![A-Programmübersicht]![Akt_Turnier] & ") AND ((Wert_Richter.WR_Kuerzel)=""" & Me!WR_Kuerzel & """) AND ((Rundentab.WB)>0)) ORDER BY Rundentab.Startklasse;")
            
        If re.EOF Then
            MsgBox "Es wurde noch keine Startklassenzuordnung gemacht!"
        Else
            print_wait_close IIf(Forms!Wertungsrichter_aufnehmen!druck, "Deckblatt_quer", "Deckblatt"), acNormal, "Ausdr3 = 'WR_" & re!WR_Kuerzel & "'"
            print_wait_close "Wertungsrichter_Einteilung", acNormal
            re.MoveFirst
            Do Until re.EOF
                If Not IsNull(re!Startklasse) Then
                    Select Case left(Nz(re!Startklasse), 3)
                    Case "BS_"
                        WB = "WertungsbogenEinzelBS"
                    Case "BW_"
                        WB = "WertungsbogenEinzelBW"
                    Case "LH_"
                        WB = "WertungsbogenEinzelLindy"
                    Case "RR_"
                        WB = "WertungsbogenEinzelRR"
                    Case "F_F", "F_R"
                        WB = "WertungsbogenFormRR"
                    Case "F_B"
                        WB = "WertungsbogenFormBW"
                    Case Else
                        Exit Sub
                    End Select
                End If
                For anz = 1 To re!WB
                    print_wait_close WB, acNormal, "WR_Kuerzel = """ & Me!WR_Kuerzel & """"
                Next
                re.MoveNext
            Loop
        End If
    End If


End Sub

Private Sub Wertungen_drucken_Click()
    
        DoCmd.OpenReport "Wertungsbogen", acViewPreview, , "wr_id = " & Me!WR_ID
                                        
End Sub

Private Sub Lizenznr_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub TL_Vorname_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub WR_kenn_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub WR_km_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub WR_Kürzel_GotFocus()
    Call refresh_Startklassen
End Sub

Private Sub WR_Kürzel_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub WR_zeit_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub
