Option Explicit
    Dim dbs As Database
    Dim stDocName As String

Private Sub Rundenreihenfolge_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub Stell_TP_ID_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub Paare_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub Stell_erst_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Dim db As Database
    Dim re As Recordset
    If Me!Stell_Reihe = 0 Then
        Set db = CurrentDb
        Set re = db.OpenRecordset("SELECT stell_Reihe FROM stellprobe ORDER BY stell_Reihe DESC;")
        Me!Stell_Reihe = re!Stell_Reihe + 1
    End If
End Sub

Private Sub Stell_TP_ID_BeforeUpdate(Cancel As Integer)
    Dim re As Recordset
    Set re = Me.RecordsetClone
    re.FindFirst "Stell_tp_id = " & Me!Stell_TP_ID
    If re.NoMatch Or Me!Stell_TP_ID = -1 Then
    Else
        MsgBox "Formation ist schon vorhanden"
        Cancel = True
        Me!Stell_TP_ID.Undo
    End If

End Sub
