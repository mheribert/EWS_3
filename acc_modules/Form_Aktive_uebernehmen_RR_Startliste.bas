Option Compare Database
Option Explicit

Private Sub Bezahlt_GotFocus()
    Call CopyTP2Textfelder
End Sub

Private Sub Dame_GotFocus()
    Call CopyTP2Textfelder
End Sub

Private Sub Form_Click()
    Call CopyTP2Textfelder
End Sub

Private Sub Herr_GotFocus()
    Call CopyTP2Textfelder
End Sub

Private Sub Kombinationsfeld26_LostFocus()
    Call CopyTP2Textfelder
End Sub

Private Sub Startbuch_GotFocus()
    Call CopyTP2Textfelder
End Sub

Private Sub Startkl_GotFocus()
    Call CopyTP2Textfelder
End Sub

Private Sub Startnr_GotFocus()
    Call CopyTP2Textfelder
End Sub

Private Sub CopyTP2Textfelder()
On Error GoTo CopyTP2Textfelder_Error
    Form_Aktive_uebernehmen!STBuchnum = Startbuch
    Form_Aktive_uebernehmen!VName_Dame = Da_Vorname
    Form_Aktive_uebernehmen!NName_Dame = Da_NAchname
    Form_Aktive_uebernehmen!Alter_Dame = Da_Alterskontrolle
    Form_Aktive_uebernehmen!VName_Herr = He_Vorname
    Form_Aktive_uebernehmen!NName_Herr = He_Nachname
    Form_Aktive_uebernehmen!Alter_Herr = He_Alterskontrolle
    Exit Sub
CopyTP2Textfelder_Error:
    'MsgBox ("Erroe")
End Sub
