Option Compare Database

Private Sub Report_Open(Cancel As Integer)
    If Forms![A-Programm�bersicht]!Turnierausw.Column(8) = "SL" Then
        Me!Tanz.Visible = True
    End If
End Sub

