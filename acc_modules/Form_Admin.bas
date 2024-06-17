Option Compare Database
Option Explicit

Private Sub close_Click()
    DoCmd.Close
End Sub

Private Sub lauf_Click()
    Dim st As String
    Dim lz_start, lz_stop As Single
    If get_properties("EWS") = "EWS3" Then
        lz_start = Timer
        st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=Serverlaufzeit&text=" & WR_ID)
        lz_stop = Timer
        If st = "server_retour" Then Me!laufzeit = (lz_stop - lz_start) * 1000
    End If

End Sub

Private Sub WR_Laufzeit_Click()
    Dim st As String
    If get_properties("EWS") = "EWS3" Then
        st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=WRlaufzeit&text=" & WR_ID)
    End If

End Sub

Private Sub Umschaltfläche61_Click()
    Dim st As String
    st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=status_wr&text=")
End Sub

Private Sub eine_Runde_zurück_Click()
    Dim st
    Dim back

    back = MsgBox("Wirklich eine Runde zurück?" & vbCrLf & " Es werden alle Wertungen überschrieben!", vbYesNo)
    
    If back = vbYes Then
        st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=eingriff&text=runde_mi")
    End If
End Sub

Private Sub nochmal_werten_Click()
    Dim st
    Dim back

    back = MsgBox("Nocheinmal werten?", vbYesNo)
    
    If back = vbYes Then
        st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=nochmal werten&text=" & WR_ID)
        If st = "alle werten" Then
            
        End If
    End If
End Sub

Private Sub alle_nochmal_werten_Click()
    Dim re As Recordset
    Dim back
    Dim st
    back = MsgBox("Alle nocheinmal werten?", vbQuestion + vbYesNo)
    
    If back = vbYes Then
        Set re = Me.RecordsetClone
        re.MoveFirst
        Do Until re.EOF
            st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=nochmal werten&text=" & re!WR_ID)
            re.MoveNext
        Loop
    End If
End Sub
