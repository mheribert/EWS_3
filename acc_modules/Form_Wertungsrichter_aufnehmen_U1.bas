Option Compare Database
Option Explicit
Dim dbs As Database

Function Einteil()

    Dim sqlcmd As String
    Dim sel As String
    Dim wr, re As Recordset
    Dim left, top As Integer
    Dim ctl As String
    Dim Art, wr_art, i
    Dim of
    Set dbs = CurrentDb
        
    wr_art = Split(Me!stkl_w, ", ")
    Art = Nz(Screen.ActiveControl)
    For i = 0 To UBound(wr_art)
        If wr_art(i) = Art Then
            Exit For
        End If
    Next
    sel = Screen.ActiveControl.Name
    ctl = sel
    sel = Me(sel).ControlSource
    Set wr = dbs.OpenRecordset("SELECT Wert_Richter.WR_ID FROM Wert_Richter WHERE (((Wert_Richter.WR_Kuerzel)=""" & Right(sel, 1) & """) AND ((Wert_Richter.Turniernr)=" & [Forms]![A-Programmübersicht]![Akt_Turnier] & "));")
    left = Me.ActiveControl.Parent.SelLeft
    top = Me.ActiveControl.Parent.SelTop
    i = i + 1
    If i > UBound(wr_art) Then i = 0
        dbs.Execute "DELETE skwr.WR_ID, skwr.Startklasse FROM Startklasse_wertungsrichter AS skwr WHERE (((skwr.WR_ID)=(SELECT TOP 1 Wert_Richter.WR_ID FROM Wert_Richter WHERE (((Wert_Richter.WR_Kuerzel)=""" & Right(sel, 1) & """) AND ((Wert_Richter.Turniernr)=" & [Forms]![A-Programmübersicht]![Akt_Turnier] & " ));)) AND ((skwr.Startklasse)= """ & Me!Startklasse & """));"
        If i > 0 Then
            dbs.Execute "INSERT into Startklasse_wertungsrichter( WR_ID, startklasse, WR_function)" & _
                     " values(" & wr!WR_ID & ", '" & Me!Startklasse & "', '" & wr_art(i) & "');"
        End If
    Me.Requery
    Me.SelTop = top
    Me(ctl).SetFocus
    Set wr = Nothing
    Set dbs = Nothing

End Function

' ***** HM14.05 *****
' an WR-Einteilung angepasst erst FT/BW dann AK
Function WR_Anzeige(Startkl, sle)
    Dim re As Recordset
    Dim anzeige As String
    Set dbs = CurrentDb
    
    Set re = dbs.OpenRecordset("SELECT Count(WR_ID) AS Ak_WR FROM Startklasse_wertungsrichter WHERE Startklasse = '" & Startkl & "' AND (WR_function = 'Ak' OR WR_function = 'X');", DB_OPEN_DYNASET)
    If re.RecordCount > 0 Then
        anzeige = re!Ak_WR
        Set re = dbs.OpenRecordset("SELECT Count(WR_ID) AS Ft_WR FROM Startklasse_wertungsrichter WHERE Startklasse = '" & Startkl & "' AND (WR_function = 'Ft');", DB_OPEN_DYNASET)
        If re.RecordCount > 0 And (InStr(1, Startkl, "BW") = 0 And InStr(1, Startkl, "BS") = 0) Then
            anzeige = re!Ft_WR & " + " & anzeige
        End If
    End If
    WR_Anzeige = anzeige
End Function

Private Sub CTRL01_KeyDown(KeyCode As Integer, Shift As Integer)
    taste_up_down KeyCode, Shift, Me.ActiveControl.Name, Me.SelTop
End Sub

Private Sub CTRL02_KeyDown(KeyCode As Integer, Shift As Integer)
    taste_up_down KeyCode, Shift, Me.ActiveControl.Name, Me.SelTop
End Sub

Private Sub CTRL03_KeyDown(KeyCode As Integer, Shift As Integer)
    taste_up_down KeyCode, Shift, Me.ActiveControl.Name, Me.SelTop
End Sub

Private Sub CTRL04_KeyDown(KeyCode As Integer, Shift As Integer)
    taste_up_down KeyCode, Shift, Me.ActiveControl.Name, Me.SelTop
End Sub

Private Sub CTRL05_KeyDown(KeyCode As Integer, Shift As Integer)
    taste_up_down KeyCode, Shift, Me.ActiveControl.Name, Me.SelTop
End Sub

Private Sub CTRL06_KeyDown(KeyCode As Integer, Shift As Integer)
    taste_up_down KeyCode, Shift, Me.ActiveControl.Name, Me.SelTop
End Sub

Private Sub CTRL07_KeyDown(KeyCode As Integer, Shift As Integer)
    taste_up_down KeyCode, Shift, Me.ActiveControl.Name, Me.SelTop
End Sub

Private Sub CTRL08_KeyDown(KeyCode As Integer, Shift As Integer)
    taste_up_down KeyCode, Shift, Me.ActiveControl.Name, Me.SelTop
End Sub

Private Sub CTRL09_KeyDown(KeyCode As Integer, Shift As Integer)
    taste_up_down KeyCode, Shift, Me.ActiveControl.Name, Me.SelTop
End Sub

Private Sub CTRL10_KeyDown(KeyCode As Integer, Shift As Integer)
    taste_up_down KeyCode, Shift, Me.ActiveControl.Name, Me.SelTop
End Sub

Private Sub CTRL11_KeyDown(KeyCode As Integer, Shift As Integer)
    taste_up_down KeyCode, Shift, Me.ActiveControl.Name, Me.SelTop
End Sub

Private Sub CTRL12_KeyDown(KeyCode As Integer, Shift As Integer)
    taste_up_down KeyCode, Shift, Me.ActiveControl.Name, Me.SelTop
End Sub

Private Sub CTRL13_KeyDown(KeyCode As Integer, Shift As Integer)
    taste_up_down KeyCode, Shift, Me.ActiveControl.Name, Me.SelTop
End Sub

Private Sub CTRL14_KeyDown(KeyCode As Integer, Shift As Integer)
    taste_up_down KeyCode, Shift, Me.ActiveControl.Name, Me.SelTop
End Sub

Private Sub CTRL15_KeyDown(KeyCode As Integer, Shift As Integer)
    taste_up_down KeyCode, Shift, Me.ActiveControl.Name, Me.SelTop
End Sub

Private Sub CTRL16_KeyDown(KeyCode As Integer, Shift As Integer)
    taste_up_down KeyCode, Shift, Me.ActiveControl.Name, Me.SelTop
End Sub

Private Sub CTRL17_KeyDown(KeyCode As Integer, Shift As Integer)
    taste_up_down KeyCode, Shift, Me.ActiveControl.Name, Me.SelTop
End Sub

Function taste_up_down(KeyCode, Shift, ctl, top)
On Error GoTo Fehlerout
    Dim sqlcmd As String
    Dim Art, sel As String
    Dim wr, re As Recordset
    Dim wr_art
    Dim i As Integer
     
    Set dbs = CurrentDb
    
    sel = Me(ctl).ControlSource
    Set wr = dbs.OpenRecordset("SELECT Wert_Richter.WR_ID FROM Wert_Richter WHERE (((Wert_Richter.WR_Kuerzel)=""" & Right(sel, 1) & """) AND ((Wert_Richter.Turniernr)=" & [Forms]![A-Programmübersicht]![Akt_Turnier] & "));")
    Set re = dbs.OpenRecordset("SELECT * FROM Startklasse_wertungsrichter WHERE WR_ID=" & wr!WR_ID & " AND startklasse ='" & Me!Startklasse & "';")
        
    wr_art = Split(Me!stkl_w, ", ")
    Art = Chr(KeyCode)
    For i = 0 To UBound(wr_art)
        If left(wr_art(i), 1) = Art Then
            Exit For
        End If
    Next
    If KeyCode = 32 Or KeyCode = 46 Then
        DoCmd.CancelEvent
        dbs.Execute "DELETE skwr.WR_ID, skwr.Startklasse FROM Startklasse_wertungsrichter AS skwr WHERE (((skwr.WR_ID)=(SELECT TOP 1 Wert_Richter.WR_ID FROM Wert_Richter WHERE (((Wert_Richter.WR_Kuerzel)=""" & Right(sel, 1) & """) AND ((Wert_Richter.Turniernr)=" & [Forms]![A-Programmübersicht]![Akt_Turnier] & " ));)) AND ((skwr.Startklasse)= """ & Me!Startklasse & """));"
    End If
    If i <= UBound(wr_art) Then
        DoCmd.CancelEvent
        If Screen.ActiveControl = "MA" Then i = i + 1
        dbs.Execute "DELETE skwr.WR_ID, skwr.Startklasse FROM Startklasse_wertungsrichter AS skwr WHERE (((skwr.WR_ID)=(SELECT TOP 1 Wert_Richter.WR_ID FROM Wert_Richter WHERE (((Wert_Richter.WR_Kuerzel)=""" & Right(sel, 1) & """) AND ((Wert_Richter.Turniernr)=" & [Forms]![A-Programmübersicht]![Akt_Turnier] & " ));)) AND ((skwr.Startklasse)= """ & Me!Startklasse & """));"
        dbs.Execute "INSERT into Startklasse_wertungsrichter( WR_ID, startklasse, WR_function)" & _
                     " values(" & wr!WR_ID & ", '" & Me!Startklasse & "', '" & wr_art(i) & "');"
    End If
        
    Me.Requery
    Me.SelTop = top
    Me(ctl).SetFocus
    If KeyCode = 40 And Shift = 0 Then
        DoCmd.GoToRecord , , acNext
        KeyCode = 0
    End If
    If KeyCode = 38 And Shift = 0 Then
        DoCmd.GoToRecord , , acPrevious
        KeyCode = 0
    End If

Fehlerout:
    If err = 2105 Then Resume Next
End Function

Function update_insert(WR_ID, st_kl, anz, func)
    Dim sqlcmd As String
    Set dbs = CurrentDb
    If anz > 0 Then
        sqlcmd = "UPDATE Startklasse_wertungsrichter SET WR_function='" & func & "' WHERE WR_ID=" & WR_ID & " AND startklasse ='" & st_kl & "';"
    Else
        sqlcmd = "INSERT into Startklasse_wertungsrichter( WR_ID, Startklasse, WR_function)" & " values(" & WR_ID & ", '" & st_kl & "', '" & func & "');"
    End If
    dbs.Execute sqlcmd
End Function

