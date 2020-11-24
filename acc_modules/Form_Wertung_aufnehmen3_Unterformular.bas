Option Compare Database
Option Explicit

Dim focus_field
Dim vars

Private Sub Befehl30_Click()
On Error GoTo Err_Befehl30_Click

    Dim stDocName As String

    stDocName = "Platzierungsliste_WR"
    DoCmd.OpenReport stDocName, acPreview

Exit_Befehl30_Click:
    Exit Sub

Err_Befehl30_Click:
    MsgBox err.Description
    Resume Exit_Befehl30_Click
    
End Sub

Private Sub Anzeige_mk_th_GotFocus()
    Me!Text38 = anzeige(" ", "mk_th")
    Me!Text38.SetFocus
End Sub

Private Sub Anzeige_mk_dh_GotFocus()
    Me!Text39 = anzeige(" ", "mk_dh")
    Me!Text39.SetFocus
End Sub

Private Sub Anzeige_mk_td_GotFocus()
    Me!Text40 = anzeige(" ", "mk_td")
    Me!Text40.SetFocus
End Sub

Private Sub Anzeige_mk_dd_GotFocus()
    Me!Text41 = anzeige(" ", "mk_dd")
    Me!Text41.SetFocus
End Sub

Private Sub Text38_LostFocus()
    write_back "mk_th", Me!Text38.text
End Sub

Private Sub Text39_LostFocus()
    write_back "mk_dh", Me!Text39.text
End Sub

Private Sub Text40_LostFocus()
    write_back "mk_td", Me!Text40.text
End Sub

Private Sub Text41_LostFocus()
    write_back "mk_dd", Me!Text41.text
End Sub

Private Sub Form_AfterUpdate()
    Form_Paare_ohne_Punkte_UF.Requery
End Sub

Function write_back(fld, wert)
    If IsNull([Cgi_Input]) Then
        [Cgi_Input] = "PR_ID1=" & [PR_ID] & "&rh1=" & [Reihenfolge] & "&rt_ID=" & [RT_ID] & "&wmk_th1=&wmk_dh1=&wmk_td1=&wmk_dd1=&wtim=" & Format(Now(), "hh_mm_ss") & "&WR_ID=7&Punkte1="
    End If
    Set vars = zerlege([Cgi_Input])
    vars("w" & fld & "1") = wert
End Function

Private Function anzeige(PR_ID, fld)
    If IsNull(Cgi_Input) Then
        anzeige = ""
    Else
        Set vars = zerlege([Cgi_Input])
        If fld = "Punkte" Then
            anzeige = vars("Punkte" & "1")
        Else
            anzeige = vars("w" & fld & "1")
        End If
    End If
End Function
