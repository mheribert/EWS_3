Option Compare Database
Option Explicit

Private Sub Seitenkopfbereich_Format(Cancel As Integer, FormatCount As Integer)
    rep_show_lines Reports!Giveaway, Split(Me!Trennlinien, ",")
End Sub
