Option Compare Database
Option Explicit

Private Sub Befehl23_Click()
    Dim rgb As Long
    DoCmd.OpenForm "Einstellungen_Color"
End Sub

Private Sub Befehl24_Click()
    Me!PPT_Font = "Arial"
    Me!PPT_Size = 0
    Me!PPT_Color = 0
    Me!PPT_Datei = False
    Me!PPT_Suffix = ".ppt"
    Me!PPT_Pfad = ""
    Me!PPT_Pfad.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
    Me!Feld11.BackColor = Me!PPT_Color
End Sub

Private Sub Form_Timer()
    Me!Feld11.BackColor = Me!PPT_Color

End Sub

Private Sub get_Pfad_Click()
    Dim nPfad As String
    nPfad = GetFolder("Ordner für Folien", Screen.ActiveForm.hwnd)
    Me!PPT_Pfad = nPfad
    Me!PPT_Pfad.Requery
End Sub
