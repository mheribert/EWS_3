Option Compare Database

Private Sub Auswahl_AfterUpdate()
gewvnr = Forms!RR_Paare_aufnehmen!auswahl.Column(0)
Me.Refresh
End Sub

Private Sub Befehl0_Click()
 DoCmd.Close
End Sub

Public Sub btnAddOffiziellen_Click()
    Dim dbs As Database
    Set dbs = CurrentDb
    
    Dim rstoff As Recordset, rsCheck As Recordset
    Dim Count As Integer
        
    If Nz(Me!VName) = "" Or Nz(Me!NName = "") Then
        MsgBox "Bitte ganzen Namen ausfüllen"
        Exit Sub
    End If
    
    ' Prüfen, ob der TL schon in der DB vorhanden ist (nur, wenn mit Lizenznr. eingegeben
    If (Not IsNull(Lizenznr) And Lizenznr <> "") Then
        sqlstr = "select count(*) as anzahl from Turnierleitung where turniernr = " & Turnier_Nummer & " and Lizenznr='" & Lizenznr & "';"
        Set rsCheck = dbs.OpenRecordset(sqlstr)
        rsCheck.MoveFirst
        Count = rsCheck!Anzahl
        rsCheck.Close
        
        If (Count > 0) Then
            MsgBox "Dieser Turnierleiter wurde dem Turnier schon hinzugefügt!"
            Exit Sub
        End If
    End If
    
    Set rstoff = dbs.OpenRecordset("select * from Turnierleitung where turniernr = " & Turnier_Nummer & ";")
    With rstoff
        .AddNew
        !Turniernr = Forms![A-Programmübersicht]![Akt_Turnier]
        !Lizenznr = Lizenznr
        !TL_Vorname = VName
        !TL_Nachname = NName
        !Vereinsnr = Club
        !Art = Lizenzart
        .Update
    End With
    Offizielle.Requery
End Sub


Private Sub FilterName_Change()
    off_auswählen.Requery
End Sub

Private Sub FilterNameEingabe_Change()
    FilterName = FilterNameEingabe.text
    off_auswählen.Requery
End Sub

Private Sub Form_Resize()
    If Me.WindowHeight > 5900 Then
        Me.Offizielle.Height = Me.WindowHeight - 4700
        Me.ScrollBars = 0
    Else
        Me.ScrollBars = 2
    End If
End Sub

