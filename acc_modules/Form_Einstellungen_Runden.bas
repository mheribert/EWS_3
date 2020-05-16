Option Compare Database
Option Explicit

Private Sub Form_Current()
    If Me!Runden_ID < 13 Then
        Me.Rundentext.Locked = True
    Else
        Me.Rundentext.Locked = False
    End If
End Sub

Private Sub Form_Delete(Cancel As Integer)
    Dim res As String
    If Me!Runden_ID < 15 Then
        Cancel = True
        MsgBox Me!Rundentext & " <- ist eine Standartvorgabe und kann nicht gelöscht werden."
    Else
        Dim re As Recordset
        res = "SELECT COUNT(Runde) AS Anz FROM Rundentab where Runde= '" & Me!Runde & "'"
        Set re = DBEngine(0)(0).OpenRecordset(res)
        If re!anz > 0 Then
            MsgBox Me!Rundentext & " <- wird verwendet und kann nicht gelöscht werden."
            Cancel = True
        End If
    End If
End Sub

Private Sub Rundentext_Change()
    Me!R_NAME_ABLAUF = Me!Rundentext.text
    If Nz(Me!Runde) = "" Then Me!Runde = "Erg_" & Me!Runden_ID
End Sub
