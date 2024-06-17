Option Compare Database
Option Explicit

Private Sub Beenden_Click()
    If check_valid_ip = True Then
        DoCmd.Close
    Else
        MsgBox "Es wurde keine gültige IP-Adresse in Serveradresse EWS2.0 eingegeben!"
    End If
End Sub

Private Sub Bezeichnungsfeld100_Click()
    Me!Einstellungen_Runden.SetFocus
    DoCmd.GoToRecord , , acNewRec
End Sub

Private Sub Bezeichnungsfeld101_Click()
    Me!Einstellungen_Deckblatt.SetFocus
    DoCmd.GoToRecord , , acNewRec
End Sub

Private Sub Bezeichnungsfeld102_Click()
    Me!Einstellungen_Rundeneinteilung.SetFocus
    DoCmd.GoToRecord , , acNewRec
End Sub

Private Sub Bezeichnungsfeld21_Click()
    Me!Untergeordnet72.Form!PROP_VALUE = ""
    Me!Untergeordnet88.Form!PROP_VALUE = ""
    Me!Untergeordnet75.Form!PROP_VALUE = ""
End Sub

Private Sub Form_Current()

    If Me!Untergeordnet96.Form!PROP_VALUE = "EWS2" Then
        Me!Einstellungen_Properties.Visible = True
        Me!Text18.Visible = True
        Me!Untergeordnet66.Visible = True
        Me!Text19.Visible = True
    Else
        Me!Einstellungen_Properties.Visible = False
        Me!Text18.Visible = False
        Me!Untergeordnet66.Visible = False
        Me!Text19.Visible = False
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If check_valid_ip = False Then
        MsgBox "Es wurde keine gültige IP-Adresse in Serveradresse EWS2.0 eingegeben!"
        Cancel = True
    End If
End Sub

Private Sub Befehl26_Click()
    start_config_webserver
End Sub

Private Sub Befehl52_Click()
    updateTLP False, True
End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim retl As Integer
    retl = EnumRegistryValues(&H80000002, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")
    Select Case Forms![A-Programmübersicht]!Turnierausw.Column(8)
        Case "SL"
            Me.Einstellungen_PPT.Visible = True
            Me.Präsentationen.Visible = True
        Case "BW"
            Me.Einstellungen_PPT.Visible = True
            Me.Präsentationen.Visible = True
        Case "BY"
            Me.Einstellungen_PPT.Visible = True
            Me.Präsentationen.Visible = True
        Case Else
    End Select
    
End Sub

Private Sub Form_Resize()
    Me!Rechteck29.Width = Me.InsideWidth - 5
End Sub

Function check_valid_ip()
    Dim strEWS2 As String
    Dim vars
    Dim i As Integer
    check_valid_ip = True
    strEWS2 = get_properties("EWS20_Adresse")
    vars = Split(strEWS2, ".")
    If strEWS2 <> "" Then
        If UBound(vars) = 3 Then
            For i = 0 To 3
                If Val(vars(i)) < 0 Or Val(vars(i)) > 255 Or Not IsNumeric(vars(i)) Then
                    check_valid_ip = False
                    Exit For
                End If
            Next
        Else
            check_valid_ip = False
        End If
    End If
End Function

Private Sub IPAddr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Then
        Forms!Einstellungen!Untergeordnet72.SetFocus
    End If
    DoCmd.CancelEvent
End Sub


