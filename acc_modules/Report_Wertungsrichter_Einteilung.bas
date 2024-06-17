Option Compare Database


Private Sub Report_Open(Cancel As Integer)
    Dim db As Database
    Dim re As Recordset
    Dim n As Integer
    
    Set db = CurrentDb
    Set re = db.OpenRecordset("SELECT Turniernr, [WR_Nachname] & "" "" & [WR_Vorname] AS Ausdr1, WR_Azubi, WR_Kuerzel From Wert_Richter WHERE (Wert_Richter.Turniernr=" & [Forms]![A-Programmübersicht]![Akt_Turnier] & ") ORDER BY WR_Kuerzel;")
    If re.RecordCount > 0 Then re.MoveFirst
    n = 1
    Do Until re.EOF
        If re!WR_AzuBi Then
            Me("Kopf" & Trim(str(n))).BackColor = rgb(255, 255, 0)
            Me("Kopf" & Trim(str(n))).BackStyle = 1
        End If
        Me("Kopf" & Trim(str(n))).Caption = re!Ausdr1
        Me("Feld" & Trim(str(n))).ControlSource = "=Get_WR(""" & re!WR_Kuerzel & """,[Startklasse])"
    
        n = n + 1
        re.MoveNext
    Loop

End Sub

