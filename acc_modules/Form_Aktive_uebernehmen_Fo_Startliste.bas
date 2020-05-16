Option Compare Database
Option Explicit

Private Sub Form_Click()
    Call Copy_Fo_Textfelder
End Sub

Private Sub Formation_GotFocus()
    Call Copy_Fo_Textfelder
End Sub

Private Sub Startbuch_GotFocus()
    Call Copy_Fo_Textfelder
End Sub

Private Sub Startnr_GotFocus()
    Call Copy_Fo_Textfelder
End Sub

Private Sub Text24_GotFocus()
    Call Copy_Fo_Textfelder
End Sub

Private Sub Verein_GotFocus()
    Call Copy_Fo_Textfelder
End Sub

Private Sub Copy_Fo_Textfelder()
    On Error GoTo Copy_Fo_Textfelder_exit
    Form_Aktive_uebernehmen!formationsname = Formation
    Form_Aktive_uebernehmen!Clubname_kurz = Verein
    Form_Aktive_uebernehmen!FBuch = Startbuch
    Form_Aktive_uebernehmen!FStartklasse = Text24
Copy_Fo_Textfelder_exit:
End Sub
