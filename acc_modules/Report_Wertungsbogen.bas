Option Compare Database
Option Explicit

Private Sub Report_Load()

'    Me.Filter = wr_wb_filter
'    Me.FilterOn = True
End Sub

Private Sub Seitenkopfbereich_Format(Cancel As Integer, FormatCount As Integer)
    If Not IsNull(Me!Trennlinien) Then
        rep_show_lines Reports!wertungsbogen, Split(Me!Trennlinien, ",")
    End If
End Sub
