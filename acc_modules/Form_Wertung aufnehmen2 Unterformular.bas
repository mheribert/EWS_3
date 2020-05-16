Option Compare Database
Option Explicit

Private Sub Befehl30_Click()
On Error GoTo Err_Befehl30_Click

    Dim stDocName As String

    stDocName = "Platzierungsliste_WR"
    DoCmd.OpenReport stDocName, acPreview

Exit_Befehl30_Click:
    Exit Sub

Err_Befehl30_Click:
    MsgBox err.Description
    Resume Exit_Befehl30_Click
    
End Sub

Private Sub Form_AfterUpdate()
    Form_Paare_ohne_Punkte_UF.Requery
End Sub

Private Sub Punkte_DblClick(Cancel As Integer)
    With Me.Recordset
    '*****AB****** V13.02 FEHLER, beim Kompilieren, deshalb auskommentiert - 1  Zeile
        'show_wertung .Fields("PR_ID").value, .Fields("Startnr").value, .Fields("wr_id").value
    End With
End Sub
