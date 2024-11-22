Option Compare Database
Option Explicit

Private Sub werb_Datei_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub werb_reihe_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub Befehl12_Click()
    DoCmd.Close acForm, "beamer_Werbung"
End Sub

Private Sub Datei_laden1_Click()
    Dim fPfad As String
    Dim fName As String
    
'    fPfad = getBaseDir & "\Apache2\htdocs\htdocs\beamer\res"
    fName = get_Filename("Bild-Dateien", "*.jpg; *.gif; *.png", fPfad)

    If fName <> "" Then
        Me!werb_Datei = Mid(fName, InStrRev(fName, "\") + 1)
        DoCmd.Requery
    End If

End Sub

Private Sub Form_Timer()
    DoCmd.GoToRecord acDataForm, "beamer_werbung", acNext
    If Me!werb_Datei = "Zeitplan" Then
        
        get_url_to_string_check ("http://" & GetIpAddrTable() & "/hand?msg=beamer_zeitplan&text=10")      ' & Tanzrunde)
    Else
        If Nz(Me!werb_Datei) = "" Then
            DoCmd.GoToRecord acDataForm, "beamer_werbung", acFirst
        End If
    
        Call Werbung_anzeigen_Click
    End If

End Sub

Private Sub Text151_KeyDown(KeyCode As Integer, Shift As Integer)
    Pfeil_up_down KeyCode, Shift
End Sub

Private Sub Umschaltfläche148_Click()

    If Me!Umschaltfläche148.Caption = "Automatik" Then
        Me!Umschaltfläche148.Caption = "Stop"
        Me.TimerInterval = Me!intervall_anzeige * 1000
    Else
        Me!Umschaltfläche148.Caption = "Automatik"
        Me.TimerInterval = 0
    End If
    
    
End Sub

Private Sub Werbung_anzeigen_Click()
    Dim st As String
    Dim cont As String
    Dim i As Integer
'    cont = "<td id=""bild"" class=""kopf"" width=""300px""><img src=""http://motion4:8082/"" alt=""DRBV"" width=""1024"" height=""768""></td>"   'stream von video
'    st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=beamer&bereich=beamer_bild&cont=<td id=""bild"" class=""kopf"" width=""300px""><img src=""BSW Boogie-Woogie Silver-Cup.jpg"" width=""290"" height=""180""></td>")
    cont = "<td id=""bild"" style=""text-align: center;""><img src=""" & Me!werb_Datei & """ width=""" & Me!werb_width & """ height=""" & Me!werb_height & """></td>"
    
    For i = 0 To Me!Beamerlist.ListCount - 1
        If Me!Beamerlist.Selected(i) Then
            st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=" & LCase(Me!Beamerlist.Column(0, i)) & "&bereich=beamer_kopf&cont=Danke an unsere Sponsoren")
            st = get_url_to_string_check("http://" & GetIpAddrTable() & "/hand?msg=" & LCase(Me!Beamerlist.Column(0, i)) & "&bereich=beamer_inhalt&cont=" & cont)
        End If
    Next

End Sub
