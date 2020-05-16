Option Compare Database
Option Explicit

Function Get_W(fld, PR_ID, Cgi_Input)
    Dim Trennlinien
    Dim i As Integer
    
    Get_W = rep_fill_fields(Reports!Wertungsbogen, fld, PR_ID, Cgi_Input, Me!Runde)
    rep_show_lines Me, Split(Reports!Wertungsbogen!Trennlinien, ",")
End Function

Private Sub Detailbereich_Format(Cancel As Integer, FormatCount As Integer)
    If Reports!Wertungsbogen.Report!WR_AzuBi = True Then
        Me.Detailbereich.BackColor = 10092543
    Else
        Me.Detailbereich.BackColor = 16777215
    End If

End Sub
