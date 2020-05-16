Option Compare Database
Option Explicit

Private Sub Detailbereich_Format(Cancel As Integer, FormatCount As Integer)
On Error Resume Next
    If left(Me!Startkl, 4) = "F_RR" Or left(Me!Startkl, 4) = "F_BW" Then
        Me!FT_Musik.Visible = False
        Me!AK_Musik.Visible = False
        Me!St_Musik.Visible = True
        Me!Fo_Musik.Visible = True
        Me!Si_Musik.Visible = True
    Else
        Me!FT_Musik.Visible = True
        Me!AK_Musik.Visible = True
        Me!St_Musik.Visible = False
        Me!Fo_Musik.Visible = False
        Me!Si_Musik.Visible = False
    End If
    
End Sub



Function musik_titel(Musik, Name)
    If Not IsNull(Musik) Then
        Dim vars
        vars = Split(Musik, "_")
        musik_titel = vars(UBound(vars)) & "_" & Name & ".mp3"
    End If
End Function
