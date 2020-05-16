Option Compare Database

Function db_Ver()
    db_Ver = DLookup("PROP_VALUE", "Properties", "Prop_Key='DB_VERSION'")
End Function

Public Function checkDBVersion()
    
    On Error GoTo wrongDBVersion
    
    Dim db As Database
    Dim rs As Recordset
    Dim value As String
    Set db = CurrentDb
    Set rs = db.OpenRecordset("Select * from Properties where PROP_KEY='DB_VERSION'")
    value = rs!PROP_VALUE
    
    rs.Close
    
    If (value <> db_Ver) Then
        GoTo wrongDBVersion
    End If
    
    Exit Function

wrongDBVersion:
    Dim result As Integer
'    result = MsgBox("Die Datendatei 'TDaten.mdb' hat die falsche Version. Bitte kopieren Sie eine neue gültige Datei in das Installationsverzeichnis. Soll das Turnierprogramm jetzt beendet werden?", vbYesNo)
    If (result = vbYes) Then
        DoCmd.Close
    End If
    
    err.Clear
End Function
