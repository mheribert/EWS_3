Option Compare Database
Option Explicit


Private Sub Report_Open(Cancel As Integer)
    Dim lae, ver As String
    
    lae = Nz(Forms![A-Programmübersicht]!Turnierausw.Column(8))
    ver = get_properties("LAENDER_VERSION")
    Me!Dokumentation.ControlSource = IIf(lae = "", ver, lae) & "_Dokumentation"
End Sub
