Option Compare Database
Option Explicit

Function Musik_W(Musik, Paar_name, rd)
    Dim vars
    If Not IsNull(Musik) Then
        vars = Split(Musik, "_")
        Musik_W = vars(UBound(vars)) & " " & Paar_name & ".mp3"
    End If
End Function

