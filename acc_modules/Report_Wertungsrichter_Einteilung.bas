Option Compare Database


Private Sub Report_Open(Cancel As Integer)
    Dim db As Database
    Dim re As Recordset
    Dim n As Integer
    
    Set db = CurrentDb
    Set re = db.OpenRecordset("SELECT Wert_Richter.Turniernr, [WR_Nachname] & "" "" & [WR_Vorname] AS Ausdr1, Wert_Richter.WR_Kuerzel From Wert_Richter WHERE (Wert_Richter.Turniernr=" & [Forms]![A-Programmübersicht]![Akt_Turnier] & " AND Wert_Richter.WR_Azubi = false) ORDER BY Wert_Richter.WR_Kuerzel;")
    re.MoveFirst
    n = 1
    Do Until re.EOF
        Me("Kopf" & Trim(str(n))).Caption = re!Ausdr1
        Me("Feld" & Trim(str(n))).ControlSource = "=Get_WR(""" & re!WR_Kuerzel & """,[Startklasse])"
    
        n = n + 1
        re.MoveNext
    Loop

End Sub

