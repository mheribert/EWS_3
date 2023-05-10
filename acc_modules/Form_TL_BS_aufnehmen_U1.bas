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
        Set objIE = CreateObject("WScript.Shell")
        objIE.Run """https://www.google.de/maps/dir/" & wr!straﬂe & ", " & wr!plz & " " & wr!ort & "/" & re!Veranst_Ort & Chr(13) & Chr(10) & """"  '"About:blank"
    End If

End Sub
