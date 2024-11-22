Option Compare Database
Option Explicit

    Dim st_kl As String
  
Private Sub Akro_speichern_Click()
    Dim db As Database
    Dim re As Recordset
    Dim fehlertext As String
    Dim retl As Integer
    Set db = CurrentDb
    Set re = db.OpenRecordset("Akrobatiken")
    Me!ak_einstufung = UCase(Me!ak_einstufung)
    Me!ak_nummer = UCase(Me!ak_nummer)
    If Nz(Me!ak_nummer) = "" Or Nz(Me!ak_text) = "" Or IsNull(Me!ak_wert) Or Nz(Me!ak_einstufung) = "" Then
        fehlertext = "Bitte alle Felder füllen! "
    End If
'    If Me!A_Startklasse.Column(2) <> left(Me!ak_nummer, 2) Then
'        fehlertext = fehlertext & "Diese Nummer passt nicht zur Startklasse! " & vbCrLf
'    End If
    If Me!ak_wert < 0.1 Or Me!ak_wert > 10 Then
        fehlertext = fehlertext & "Wertigkeit stimmt nicht! " & vbCrLf
    End If
    If InStr("LMS", Me!ak_einstufung) = 0 Or Len(Me!ak_einstufung) > 1 Then
        fehlertext = fehlertext & "Einstufung stimmt nicht! " & vbCrLf
    End If
    
    If Len(fehlertext) = 0 Then
        re.FindFirst "[Nr#] = '" & Me!ak_nummer & "'"
        If Not re.NoMatch Then
            retl = MsgBox("Diese Nummer existiert bereits! ", vbOKCancel)
        End If
        If retl = 0 Or retl = vbOK Then
            If retl = 0 Then re.AddNew
            If retl = 1 Then re.Edit
            re![Nr#] = Me!ak_nummer
            re!akrobatik = Me!ak_nummer
            re!langtext = Me!ak_text
            re!einstufung = Me!ak_einstufung
            If st_kl = "RR_A" Then
                re!RR_A = Me!ak_wert
            Else
                re!RR_C = Me!ak_wert
            End If
            
            re.Update
            DoCmd.Requery
            DoCmd.RepaintObject acForm, "Paare_Akrobatiken"
        End If
    Else
        MsgBox fehlertext
    End If
End Sub

Private Sub Akrobatiken_AfterUpdate()

    Me!ak_nummer = Me!Akrobatiken.Column(0)
    Me!ak_text = Replace(Me!Akrobatiken.Column(1), Me!Akrobatiken.Column(0) & " ", "")
    Me!ak_wert = Me!Akrobatiken.Column(2)
    Me!ak_einstufung = Me!Akrobatiken.Column(3)
    
    Me!Akrobatiken = ""
End Sub

Private Sub Form_Load()
    Select Case Forms!paare_akrobatiken!Text97  'Startklasse
        Case "BS_BW_FO"
            st_kl = "RR_C"
        Case "BS_BW_HA"
            st_kl = "RR_A"
    End Select
        
    Me!Akrobatiken.RowSource = "SELECT Akrobatik, [nr#] & ' ' & Langtext, " & st_kl & ", Akrobatiken.Einstufung FROM Akrobatiken WHERE Nz([" & st_kl & "])>='0' ORDER BY 1;"
End Sub

Private Sub Schliessen_Click()
    DoCmd.Close
End Sub

