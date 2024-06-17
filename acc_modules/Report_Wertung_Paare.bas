Option Compare Database
Option Explicit

Private Sub Seitenkopfbereich_Format(Cancel As Integer, FormatCount As Integer)
    If Not IsNull(Me!Trennlinien) Then
        rep_show_lines Reports!Wertung_Paare, Split(Me!Trennlinien, ",")
    End If
End Sub
