Option Compare Database
Option Explicit

Function Get_W(fld, PR_ID, Cgi_Input)
    Dim Trennlinien
    Dim i As Integer
    
    Get_W = rep_fill_fields(Reports!Giveaway, fld, PR_ID, Cgi_Input, Me!Runde)
    rep_show_lines Me, Split(Reports!Giveaway!Trennlinien, ",")
    
End Function

Private Sub Detailbereich_Format(Cancel As Integer, FormatCount As Integer)
    'If Reports!wertungsbogen.Report!WR_AzuBi = True Then
       ' Me.Seitenkopfbereich.BackColor = 6750207
       ' Me.Detailbereich.BackColor = 6750207
    'Else
       ' Me.Seitenkopfbereich.BackColor = 16777215
        'Me.Detailbereich.BackColor = 16777215
   ' End If

End Sub
