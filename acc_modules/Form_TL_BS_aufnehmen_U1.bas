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

Private Sub km_holen_Click()
    Dim db As Database
    Dim wr, re As Recordset
    Dim objIE As Object
    Dim ti, s
    Set db = CurrentDb
    Set wr = db.OpenRecordset("SELECT * FROM TLP_OFFIZIELLE WHERE Lizenzn=""" & Me!Lizenznr & """;")
    Set re = db.OpenRecordset("SELECT * FROM turnier WHERE turniernum=" & get_aktTNr & ";")

    If wr.RecordCount > 0 Then
        Set objIE = CreateObject("InternetExplorer.Application")
        objIE.Navigate2 "https://www.google.de/maps/dir/" & wr!straﬂe & ", " & wr!plz & " " & wr!ort & "/" & re!Veranst_Ort & Chr(13) & Chr(10)   '"About:blank"
        objIE.Visible = True
        'Shell ("C:\Program Files\Mozilla Firefox\firefox.exe " & """https://www.google.de/maps/dir/" & wr!straﬂe & ", " & wr!plz & " " & wr!ort & "/" & re!Veranst_Ort & """")
        'ti = Time
        'Do Until ti + 0.00004 < Time
        
        'Loop
        ' 1 str
        ' 2 plz 3 ort
        ' 5 ziel
        'Debug.Print "from: " & wr!straﬂe & ", " & wr!plz & " " & wr!ort & " to: " & Me!Liste1.Column(5) & Chr(13)
        'objIE.Document.Forms.Item(0).elements("q").value = "from: " & wr!straﬂe & ", " & wr!plz & " " & wr!ort & " to: " & Forms![A-Programm¸bersicht]!Turnierauswahl.Column(6) & Chr(13) & Chr(10)
    End If

End Sub
