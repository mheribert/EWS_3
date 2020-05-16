Option Compare Database

Private Sub Report_Load()
    Dim fil As String
    Dim stkl As String
    stkl = Nz(Forms![Ausdrucke]![Startklasse_einstellen])
    
    
'    stDocName = "gemeldete_Akrobatiken"

    If stkl = "" Then
        Me.FilterOn = False
    Else
        Me.Filter = "Startkl = '" & stkl & "'"
        Me.FilterOn = True
    End If

End Sub
