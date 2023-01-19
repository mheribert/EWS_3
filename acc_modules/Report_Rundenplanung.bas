Option Compare Database
Option Explicit

Private Sub Report_Open(Cancel As Integer)
    
    Dim re As Recordset
    Dim dbs As Database
    Dim ctlC As Control
    Dim fld As String
    Dim retl As Integer
    
    Set dbs = CurrentDb()
    fld = Replace(Me.RecordSource, "[Formulare]![A-Programmübersicht]![akt_Turnier]", [Forms]![A-Programmübersicht]![Akt_Turnier])
    Set re = dbs.OpenRecordset(fld)
    
    For Each ctlC In Me.Controls
        If ctlC.ControlType = acTextBox And Mid(ctlC.Name, 1, 2) = "f_" Then
            fld = Mid(ctlC.Name, 3)
            If tst_fl(fld, re) Then
                ctlC.ControlSource = IIf(fld = "Vorrunde" Or fld = "Endrunde", "=[" & fld & "] & [" & fld & " Akrobatik]", fld)
            End If
        End If
    Next ctlC


End Sub

Private Function tst_fl(fld, re As Recordset)
    On Error GoTo open_err
    tst_fl = True
open_err:
End Function
