Option Compare Database
Option Explicit

Private Sub alle_Akrobatiken_Click()
    
    If Me.PaarAkrobatiken Then
        AkrobatikenStartklasseRowsource
        Me.alle_Akrobatiken.Caption = "nur Paar Akrobatiken freischalten"
        Me.PaarAkrobatiken = False
    Else
        AkrobatikenPaarRowsource
        Me.alle_Akrobatiken.Caption = "alle Akrobatiken freischalten"
        Me.PaarAkrobatiken = True
    End If

End Sub

Private Sub Beenden_Click()
On Error GoTo Err_Beenden_Click
    Dim i As Integer
    Dim VR As Integer
    Dim ZR As Integer
    Dim ER As Integer
    
    For i = 1 To 8
        If Nz(Me("Akro" & i & "_VR")) <> "" Then
            VR = VR + 1
        End If
        If Nz(Me("Akro" & i & "_ZR")) <> "" Then
            ZR = ZR + 1
        End If
        If Nz(Me("Akro" & i & "_ER")) <> "" Then
            ER = ER + 1
        End If
    Next
    If (ZR = 0 Or ER = 0) And VR > 0 Then
        MsgBox "Es gibt keine Akrobatiken in Zwischenrunde und/oder Endrunde!"
    End If
    DoCmd.Close

Exit_Beenden_Click:
    Exit Sub

Err_Beenden_Click:
    MsgBox err.Description
    Resume Exit_Beenden_Click
    
End Sub

Private Sub btn_Paare_Akrobatiken_1_Click()
    Dim i As Integer
    For i = 1 To 8
        Me("Akro" & i & "_ZR") = Me("Akro" & i & "_VR")
        Me("Wert" & i & "_ZR") = Me("Wert" & i & "_VR")
        Me("Akro" & i & "_ER") = Me("Akro" & i & "_VR")
        Me("Wert" & i & "_ER") = Me("Wert" & i & "_VR")
    Next
    DoCmd.Requery
End Sub

Private Sub btn_Paare_Akrobatiken_2_Click()
    DoCmd.OpenForm "Tanzpaare_neueAkro"

End Sub

Private Sub Form_Load()
        
    AkrobatikenPaarRowsource
    MusikListenFuellen
    AnzahlTaenzereinstellen

End Sub

Function fill_akro()
    Dim db As Database
    Dim re As Recordset
    Dim Tanzrunde
    Dim fld
    Dim i, j, gr_id As Integer
    Dim a_id As Variant
    Dim f_text As String
    Dim ctl As Control
    Dim grid_text As String
    Dim Gruppen_ID(40)
    Dim idcheck(12)     ' FD=11   FF=12
    Dim alteWerte(2)

    fld = Replace(Me.ActiveControl.Name, "Akro", "Wert")
    alteWerte(1) = Me.ActiveControl.OldValue
    alteWerte(2) = Me(fld)
    Me(fld) = Me(ActiveControl.Name).Column(2)
    
    Set db = CurrentDb
    Tanzrunde = Right(Me.ActiveControl.Name, 2)
    For i = 1 To 8
        For j = 1 To 3
            a_id = Mid(Me("Akro" & i & "_" & Tanzrunde).Column(1), j, 1)
            If IsNumeric(a_id) Then Exit For
        Next j
        If IsNumeric(a_id) Then
            If left(Me("Akro" & i & "_" & Tanzrunde), 2) = "FD" Then
                idcheck(11) = idcheck(11) + 1
            End If
            If left(Me("Akro" & i & "_" & Tanzrunde), 2) = "FF" Then
                idcheck(12) = idcheck(12) + 1
            End If
            Me("ID" & i & "_" & Tanzrunde) = a_id
            If Not IsNull(a_id) Then
                idcheck(a_id) = idcheck(a_id) + 1
                Set re = db.OpenRecordset("SELECT * FROM Akrobatiken WHERE Akrobatik='" & Me("Akro" & i & "_" & Tanzrunde) & "';")
                For j = 1 To 5
                    If re("Gruppen_ID_" & j) <> 0 Then
                        Gruppen_ID(gr_id) = re("Gruppen_ID_" & j)
                        grid_text = grid_text & re("Gruppen_ID_" & j) & " "
                        gr_id = gr_id + 1
                    End If
                Next
                Me("GR_ID" & i & "_" & Tanzrunde) = grid_text
                grid_text = ""
            End If
        End If
    Next
    If Me!Startkl = "RR_A" Or Me!Startkl = "RR_B" Then
        If idcheck(0) = 0 Then
            f_text = "Die Kategorie (0) Rotationen ist nicht belegt worden!" & vbCrLf & vbCrLf
        End If
        If idcheck(3) = 0 Then
            f_text = f_text & "Die Kategorie (3) Vorwärtselement ist nicht belegt worden!" & vbCrLf & vbCrLf
        End If
        If idcheck(4) = 0 Then
            f_text = f_text & "Die Kategorie (4) Rückwärtselement ist nicht belegt worden!" & vbCrLf & vbCrLf
        End If
        If idcheck(5) = 0 Then
            f_text = f_text & "Die Kategorie (5) Kopfüberelement ist nicht belegt worden!" & vbCrLf & vbCrLf
        End If
        If idcheck(5) > 3 Then
            f_text = f_text & "Die max. Anzahl an Kopfüberelementen wurde überschritten!" & vbCrLf & vbCrLf
        End If
        If idcheck(8) > 2 Then
           f_text = f_text & "Die max. Anzahl der erlaubten Kombinationen (8) wurde überschritten!" & vbCrLf & vbCrLf
        End If
        If idcheck(9) > 2 Then
            f_text = f_text & "Die max. Anzahl der erlaubten Rotationen (9) wurde überschritten!" & vbCrLf & vbCrLf
        End If
    End If
    If Me!Startkl = "F_RR_M" Then
        If idcheck(11) > 2 Then
            f_text = f_text & "Die max. Anzahl an Dualakrobatiken wurde überschritten!" & vbCrLf & vbCrLf
        End If
        If idcheck(12) > 2 Then
            f_text = f_text & "Die max. Anzahl an formationsspezifischen Akrobatiken wurde überschritten!" & vbCrLf & vbCrLf
        End If
    End If
    If check_doppelte(gr_id, Gruppen_ID) Then f_text = f_text & "Es gibt min. 2 Akrobatiken mit gleicher Gruppen ID!" & vbCrLf & vbCrLf
    If Me("Summe_" & Tanzrunde) > 70 Then
        f_text = f_text & "Der max. zulässige akrobatische Vorwert wurde überschritten!"
    End If
    If f_text <> "" Then
        Me!Gruppen_ID.Visible = True
        MsgBox "Der Tausch ist nicht erlaubt!" & vbCrLf & vbCrLf & f_text, , "Achtung"
        Me(ActiveControl.Name) = alteWerte(1)
        Me(fld) = alteWerte(2)
        For Each ctl In Controls
            If (left(ctl.Name, 5) = "GR_ID" Or left(ctl.Name, 2) = "ID") And Right(ctl.Name, 3) = Right(fld, 3) Then
                ctl.Value = ""
            End If
        Next
        Me!Gruppen_ID.Visible = False
    End If

End Function

Function check_doppelte(max, Gruppen_ID)
    Dim i, j As Integer
    check_doppelte = False
    For i = 0 To max - 1
        For j = i + 1 To max - 1
'            Debug.Print Gruppen_ID(i), Gruppen_ID(j)
            If Gruppen_ID(i) = Gruppen_ID(j) Then
                check_doppelte = True
                Exit Function
            End If
        Next
    Next

End Function

Private Sub AkrobatikenPaarRowsource____neu()
Dim sql As Variant
Dim Startklasse, Tanzrunde As String
Dim akronummer, TP_ID As Integer

Startklasse = Forms!Tanzpaare_aufnehmen!Startkl
TP_ID = Me.TP_ID

    For akronummer = 1 To 8
        Tanzrunde = "ER"
        sql = "SELECT Akrobatiken.Akrobatik, [Nr#] & ' ' & [Langtext] AS Ausdr2, Akrobatiken.RR_A FROM Akrobatiken;"


        Me("Akro" & akronummer & "_" & Tanzrunde).RowSource = sql


    Next
End Sub

Private Sub AkrobatikenPaarRowsource()
' diese Funktion sucht mit den Parametern das Paar und erstellt ein SQL Skript für die RowSource mit den jeweiligen Akrobatiken und Ersatzakrobatiken

    Dim sql As Variant
    Dim Startklasse, Tanzrunde As String
    Dim akronummer, TP_ID As Integer
    
    Select Case Forms!Tanzpaare_aufnehmen!Startkl
        Case "BS_BW_EI"
            Startklasse = "RR_S"
        Case "BS_BW_FO"
            Startklasse = "RR_C"
        Case "BS_BW_HA"
            Startklasse = "RR_A"
        Case "BS_F_RR_JF", "BS_F_RR_EF"
            Startklasse = "F_RR_M"
        Case Else
            Startklasse = Forms!Tanzpaare_aufnehmen!Startkl
    End Select
    TP_ID = Me.TP_ID
    
    For akronummer = 1 To 8
        Tanzrunde = "VR"
        
        sql = "SELECT TOP 1 '' AS Ausdr1, ' < keine > ' AS Ausdr2, '' AS Ausdr3 FROM Tanz_Runden UNION SELECT Akrobatik, [nr#] & ' ' & Langtext, " & Startklasse & " FROM Akrobatiken  WHERE [Nr#]='ALL' UNION "
        sql = sql & "SELECT Akrobatiken.Akrobatik, [Nr#] & ' ' & [Langtext] AS Ausdr2, Akrobatiken." & Startklasse & " FROM Paare INNER JOIN Akrobatiken ON Paare.Akro" & akronummer & "_" & Tanzrunde & " = Akrobatiken.Akrobatik WHERE (((Paare.TP_ID)=" & TP_ID & " and " & Startklasse & " >='0')) UNION "
        sql = sql & "SELECT Akrobatiken.Akrobatik, [Nr#] & ' ' & [Langtext] AS Ausdr2, Akrobatiken." & Startklasse & " FROM Paare INNER JOIN Akrobatiken ON Paare.E_Akro1_" & Tanzrunde & " = Akrobatiken.Akrobatik WHERE (((Paare.TP_ID)=" & TP_ID & " and " & Startklasse & " >='0')) UNION "
        sql = sql & "SELECT Akrobatiken.Akrobatik, [Nr#] & ' ' & [Langtext] AS Ausdr2, Akrobatiken." & Startklasse & " FROM Paare INNER JOIN Akrobatiken ON Paare.E_Akro2_" & Tanzrunde & " = Akrobatiken.Akrobatik WHERE (((Paare.TP_ID)=" & TP_ID & " and " & Startklasse & " >='0')) ORDER BY Ausdr2;"

    
        Me("Akro" & akronummer & "_" & Tanzrunde).RowSource = sql
'        Me("Wert" & Akronummer & "_" & Tanzrunde).ControlSource = "=[Akro" & Akronummer & "_VR].[column](2)"
    
        Tanzrunde = "ZR"
        
        sql = "SELECT TOP 1 '' AS Ausdr1, ' < keine > ' AS Ausdr2, '' AS Ausdr3 FROM Tanz_Runden UNION SELECT Akrobatik, [nr#] & ' ' & Langtext, " & Startklasse & " FROM Akrobatiken  WHERE [Nr#]='ALL' UNION "
        sql = sql & "SELECT Akrobatiken.Akrobatik, [Nr#] & ' ' & [Langtext] AS Ausdr2, Akrobatiken." & Startklasse & " FROM Paare INNER JOIN Akrobatiken ON Paare.Akro" & akronummer & "_" & Tanzrunde & " = Akrobatiken.Akrobatik WHERE (((Paare.TP_ID)=" & TP_ID & " and " & Startklasse & " >='0')) UNION "
        sql = sql & "SELECT Akrobatiken.Akrobatik, [Nr#] & ' ' & [Langtext] AS Ausdr2, Akrobatiken." & Startklasse & " FROM Paare INNER JOIN Akrobatiken ON Paare.E_Akro1_" & Tanzrunde & " = Akrobatiken.Akrobatik WHERE (((Paare.TP_ID)=" & TP_ID & " and " & Startklasse & " >='0')) UNION "
        sql = sql & "SELECT Akrobatiken.Akrobatik, [Nr#] & ' ' & [Langtext] AS Ausdr2, Akrobatiken." & Startklasse & " FROM Paare INNER JOIN Akrobatiken ON Paare.E_Akro2_" & Tanzrunde & " = Akrobatiken.Akrobatik WHERE (((Paare.TP_ID)=" & TP_ID & " and " & Startklasse & " >='0')) ORDER BY Ausdr2;"

    
        Me("Akro" & akronummer & "_" & Tanzrunde).RowSource = sql
'        Me("Wert" & Akronummer & "_" & Tanzrunde).ControlSource = "=[Akro" & Akronummer & "_ZR].[column](2)"
        
        Tanzrunde = "ER"
        
        sql = "SELECT TOP 1 '' AS Ausdr1, ' < keine > ' AS Ausdr2, '' AS Ausdr3 FROM Tanz_Runden UNION SELECT Akrobatik, [nr#] & ' ' & Langtext, " & Startklasse & " FROM Akrobatiken  WHERE [Nr#]='ALL' UNION "
        sql = sql & "SELECT Akrobatiken.Akrobatik, [Nr#] & ' ' & [Langtext] AS Ausdr2, Akrobatiken." & Startklasse & " FROM Paare INNER JOIN Akrobatiken ON Paare.Akro" & akronummer & "_" & Tanzrunde & " = Akrobatiken.Akrobatik WHERE (((Paare.TP_ID)=" & TP_ID & " and " & Startklasse & " >='0')) UNION "
        sql = sql & "SELECT Akrobatiken.Akrobatik, [Nr#] & ' ' & [Langtext] AS Ausdr2, Akrobatiken." & Startklasse & " FROM Paare INNER JOIN Akrobatiken ON Paare.E_Akro1_" & Tanzrunde & " = Akrobatiken.Akrobatik WHERE (((Paare.TP_ID)=" & TP_ID & " and " & Startklasse & " >='0')) UNION "
        sql = sql & "SELECT Akrobatiken.Akrobatik, [Nr#] & ' ' & [Langtext] AS Ausdr2, Akrobatiken." & Startklasse & " FROM Paare INNER JOIN Akrobatiken ON Paare.E_Akro2_" & Tanzrunde & " = Akrobatiken.Akrobatik WHERE (((Paare.TP_ID)=" & TP_ID & " and " & Startklasse & " >='0')) ORDER BY Ausdr2;"

        Me("Akro" & akronummer & "_" & Tanzrunde).RowSource = sql
'        Me("Wert" & Akronummer & "_" & Tanzrunde).ControlSource = "=[Akro" & Akronummer & "_ER].[column](2)"
    Next
End Sub

Public Sub AkrobatikenStartklasseRowsource()
' diese Funktion füllt die Akrobatiken mit allen in der Startklasse möglichen Akrobatiken

    Dim sql As String
    Dim Startklasse As String
    Dim i As Integer
        
    Select Case Forms!Tanzpaare_aufnehmen!Startkl
        Case "BS_BW_EI"
            Startklasse = "RR_S"
        Case "BS_BW_FO"
            Startklasse = "RR_C"
        Case "BS_BW_HA"
            Startklasse = "RR_A"
        Case "BS_F_RR_JF", "BS_F_RR_EF"
            Startklasse = "F_RR_M"
        Case Else
            Startklasse = Forms!Tanzpaare_aufnehmen!Startkl
    End Select
    
    sql = "SELECT TOP 1 '' AS Ausdr1, ' < keine > ' AS Ausdr2, '' AS Ausdr3 FROM Tanz_Runden UNION SELECT Akrobatik, [nr#] & ' ' & Langtext, " & Startklasse & " FROM Akrobatiken WHERE Nz([" & Startklasse & "])>='0' ORDER BY Ausdr2;"
    For i = 1 To 8
        Me("Akro" & i & "_VR").RowSource = sql
        Me("Akro" & i & "_ZR").RowSource = sql
        Me("Akro" & i & "_ER").RowSource = sql
    Next

End Sub

Private Sub MusikListenFuellen()
    Dim db As Database
    Dim Paare As DAO.Recordset
    Dim RowSourceString As String
    
    Set db = CurrentDb()
    Set Paare = db.OpenRecordset("Select * from paare where TP_ID = " & Me.TP_ID)
    
    If Not Paare.EOF Then
        RowSourceString = ";"
        
        If (Not IsNull(Paare!Musik_FT)) And Not Paare!Musik_FT = "" Then RowSourceString = RowSourceString & Paare!Musik_FT & ";"
        If (Not IsNull(Paare!Musik_Akro)) And Not Paare!Musik_Akro = "" Then RowSourceString = RowSourceString & Paare!Musik_Akro & ";"
        If (Not IsNull(Paare!Musik_Stell)) And Not Paare!Musik_Stell = "" Then RowSourceString = RowSourceString & Paare!Musik_Stell & ";"
        If (Not IsNull(Paare!Musik_Form)) And Not Paare!Musik_Form = "" Then RowSourceString = RowSourceString & Paare!Musik_Form & ";"
        If (Not IsNull(Paare!Musik_Sieg)) And Not Paare!Musik_Sieg = "" Then RowSourceString = RowSourceString & Paare!Musik_Sieg & ";"
    
        
        Me.MusikAkrobatik.RowSource = RowSourceString
        Me.MusikFusstechnik.RowSource = RowSourceString
        Me.MusikStellprobe.RowSource = RowSourceString
        Me.MusikFormation.RowSource = RowSourceString
        Me.MusikSiegertanz.RowSource = RowSourceString
    End If

End Sub

Sub AnzahlTaenzereinstellen()
    Dim f As Formationswerte
    Dim werte As String
    Dim i As Integer
    f = Faktor_Formation_Abzuege(Me!Startkl)
    werte = " "
    For i = f.min To f.max
        werte = werte & ";" & i
    Next
    Me!AnzahlTaenzerInnen.RowSource = werte
End Sub

Private Sub Form_Open(Cancel As Integer)
    setzte_buttons Me.Name, Me.Name, Forms![A-Programmübersicht]!Turnierausw.Column(8)
End Sub

Private Sub Form_Resize()
    Me.Rechteck191.Width = Me.InsideWidth - 100
    Me.Rechteck192.Width = Me.InsideWidth - 100
End Sub
