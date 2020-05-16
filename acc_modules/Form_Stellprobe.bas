Option Compare Database
Option Explicit
    Const seku = 1.15740740740741E-05
    Dim count_down As Integer

Private Sub Kombinationsfeld53_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub alle_holen_Click()
    Dim db As Database
    Dim quelle, ziel As Recordset
    Dim i As Integer
    
    Set db = CurrentDb
    i = vbYes
    If Me!Stellprobe_Liste.Form.RecordsetClone.RecordCount > 0 Then
        i = MsgBox("Es werden alle vorhandenen Formationen gelöscht" & vbCrLf & "weitermachen?", vbYesNo, "Turnierprogramm")
    End If
    If i = vbYes Then
        db.Execute "DELETE * FROM stellprobe;"
        Set quelle = db.OpenRecordset("SELECT TP_ID, Name_Team, Verein_Name FROM Paare WHERE (Anwesent_Status=1 AND Paare.Da_Nachname Is Null) ORDER BY Verein_Name, Name_Team;")
        Set ziel = db.OpenRecordset("stellprobe")
        If quelle.RecordCount > 0 Then
            quelle.MoveFirst
            i = 1
            Do Until quelle.EOF
                ziel.AddNew
                ziel!Stell_TP_ID = quelle.TP_ID
                ziel!Stell_Reihe = i
                ziel.Update
                i = i + 1
                quelle.MoveNext
            Loop
        Else
            MsgBox "Keine Formationen vorhanden."
        End If
    End If
    DoCmd.Requery "Stellprobe_Liste"
End Sub

Private Sub Form_Current()
    Dim re As Recordset
    If Not Me.NewRecord Then

        Set re = Me.RecordsetClone
        If Me.RecordsetClone.RecordCount > 0 Then
            re.Bookmark = Me.Bookmark
        End If
        If re.EOF Then
            Me!Danach = ""
        Else
            re.MoveNext
            Me!Danach = IIf(re.EOF, "", re!tName)
            Me!Danach_verein = IIf(re.EOF, "", re!Verein_Name)
            re.MovePrevious
        End If
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowHeight > 7100 Then
        Me.ScrollBars = 0
        Me.Stellprobe_Liste.Height = Me.WindowHeight - 3000
        Me.Formationen.Height = Me.WindowHeight - 2500
        Me!RegisterStr87.Height = Me.WindowHeight - 1500
    Else
        Me.ScrollBars = 2
    End If
End Sub

Private Sub next_rec_Click()
On Error Resume Next
    DoCmd.GoToRecord , , acNext
End Sub

Private Sub RegisterStr87_Click()
    If Nz(Me.stell_starten) = False Then DoCmd.Requery
End Sub

Private Sub schliesssen_Click()
    DoCmd.Close
End Sub

Private Sub btnAktualisieren_Click()
    DoCmd.Requery "Stellprobe_Liste"
End Sub

Private Sub stell_starten_Click()
    If Me.stell_starten Then
        Me.stell_starten.Caption = "Stop"
        count_down = Me!vorgabe
        Folie_anzeigen_Click
        Me.TimerInterval = 1000
        Me.Folie_anzeigen.Enabled = False
        Me.next_rec.Enabled = False
    Else
        Me.TimerInterval = 0
        Me.stell_starten.Caption = "Starten"
        vorgabe_AfterUpdate
        Me.Folie_anzeigen.Enabled = True
        Me.next_rec.Enabled = True
    End If
End Sub

Private Sub Form_Timer()
    Me!stell_zeit.Caption = Int(count_down / 60) & ":" & Format(Int(count_down Mod 60), "00")
    If count_down = 0 Then
        Me.stell_starten.SetFocus
        count_down = Me!vorgabe
        next_rec_Click
        Folie_anzeigen_Click
        If Me!Jetzt = "Pause" Then
            stell_starten_Click
        End If
    Else
        count_down = count_down - 1
        If Me!vorgabe > 225 And count_down < Me!vorgabe - 225 Then
            Me.verkürzen.Visible = True
        Else
            Me.verkürzen.Visible = False
        End If
     End If
End Sub

Private Sub Stellprobe_drucken__Aktualisieren_Click()
    DoCmd.OpenReport "Stellprobe", acViewPreview
End Sub

Private Sub verkürzen_Click()
    count_down = 0
End Sub

Private Sub vorgabe_AfterUpdate()
    Me!stell_zeit.Caption = Int(Me!vorgabe / 60) & ":" & Format(Int(Me!vorgabe Mod 60), "00")
    count_down = Me!vorgabe
End Sub

Private Sub Folie_anzeigen_Click()
    Dim re As Recordset
    Dim out
    Dim startHTML As String
    Dim StellHTML As String
    Dim HTMLtext As String
    Dim next_HTML As String
    Dim ht_pfad As String
    Dim startseite
    Dim line As String
    Dim i, s As Integer
    
    Set re = Me.RecordsetClone
    If Not Me.NewRecord Then
        re.Bookmark = Me.Bookmark
        ht_pfad = getBaseDir & "Apache2\htdocs\beamer\"
        Me!Stell_erst = True
        ' Einstiegsseite scheiben
        startHTML = "<!DOCTYPE html PUBLIC ""-//W3C//DTD HTML 4.01//EN"" ><html><head><meta http-equiv=""refresh"" content=""0; URL=" & _
                     "st" & Format(re!Stell_Reihe, "00000") & ".html""><title></title></head><body></body></html>"
        Set out = file_handle(ht_pfad & "stellprobe.html")
        out.writeline (startHTML)
        out.Close
        ' Countdownseite + Warteseite scheiben
        For i = 0 To 1
            line = get_line("Beamer", "Stellprobe", i)  'holt HTML-Seite aus HTML-Block
            line = Replace(line, "x__turnier", Umlaute_Umwandeln(Forms![A-Programmübersicht]!Turnierbez))
            line = Replace(line, "x__jetzt", Umlaute_Umwandeln(re!tName) & "</strong><br>" & Umlaute_Umwandeln(re!Verein_Name))
            If re.EOF Then
                line = Replace(line, "x__danach", "&nbsp;")
            Else
                re.MoveNext
                If re.EOF Then
                    line = Replace(line, "x__danach", "&nbsp;")
                    next_HTML = "st" & Format(10000, "00000") & ".html"
                Else
                    line = Replace(line, "x__danach", IIf(re.EOF, "", Umlaute_Umwandeln(re!tName) & "</strong><br>" & Umlaute_Umwandeln(re!Verein_Name)))
                    next_HTML = "st" & Format(re!Stell_Reihe, "00000") & ".html"
                End If
                re.MovePrevious
            End If
            
            If i = 0 Then
                Set out = file_handle(ht_pfad & "st" & Format(re!Stell_Reihe, "00000") & ".html")
            Else
                Set out = file_handle(ht_pfad & next_HTML)
            End If
            line = Replace(line, "x__html", next_HTML)
            
            out.writeline (line)
            out.Close
        Next
    End If
    If re.RecordCount > 0 Then re.MoveFirst
    i = 0
    Do Until re.EOF
        If re!Stell_erst = False Then i = i + 1
        
        re.MoveNext
    Loop
    Me!Ende_ca.Caption = "Ende ca.: " & Format(Now() + (i * Me!vorgabe * seku), "hh:mm")
End Sub

Private Sub Zeit_eintragen__drucken_Click()
    Dim db As Database
    Dim re As Recordset
    Dim rt_stellprobe As Date
    Dim anz_form As Integer
    Set db = CurrentDb
    Set re = db.OpenRecordset("SELECT Startzeit from Rundentab WHERE Runde = 'Stellpr';")
    If re.RecordCount = 0 Then
        MsgBox "Es ist keine Stellprobe im Ablaufplan erstellt!", vbCritical, "Turnierprogramm"
    Else
        rt_stellprobe = re!Startzeit
        Set re = Me!Stellprobe_Liste.Form.RecordsetClone
        re.MoveFirst
        anz_form = 0
        Do Until re.EOF
            re.Edit
            re!Stell_Start = rt_stellprobe + (Me!vorgabe * anz_form * seku)
            re.Update
            anz_form = anz_form + 1
            re.MoveNext
        Loop
    End If

End Sub
