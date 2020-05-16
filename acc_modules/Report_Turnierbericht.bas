Option Compare Database
Option Explicit

Private Sub Report_Close()
Dim stDocName As String
stDocName = "unentschuldigt_gefehlte_paare"
DoCmd.OpenReport stDocName, acPreview
End Sub
