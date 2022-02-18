Option Compare Database
Option Explicit

Private Sub schliessen_Click()
    DoCmd.Close
End Sub

Private Sub Befehl19_Click()
    ' Zuerst �berpr�fen, ob der Benutzer schon die richtigen Daten ausgew�hlt hat
    If (IsNull(Forms![Paare_in erste Runde nehmen]!Startklasse)) Then
        MsgBox "Bitte w�hlen Sie zuerst eine Startklasse aus!"
        Exit Sub
    End If
    
    If (IsNull(n�chste_Runde)) Then
        MsgBox "Bitte w�hlen Sie zuerst die n�chste Runde aus!"
        Exit Sub
    End If

    Dim dbs As Database
    Dim rstErste, rstpaare, rs As Recordset
    Dim sk As String
    Dim sqlString As String
    Dim MehrkampfrundenAnzahl, Zaehler, i As Integer
    Dim MehrkampfEndrunde As Boolean
    
    Set dbs = CurrentDb
    
    sk = Forms![Paare_in erste Runde nehmen]!Startklasse

    Set rstErste = dbs.OpenRecordset("select * from paare_rundenqualifikation")
    
    ' Den Eintrag in der Tabelle Rundentab ermitteln
    Set rs = dbs.OpenRecordset("select * from Rundentab where startklasse = '" & sk & "' and turniernr = " & Me!T_Nr & " and runde='" & n�chste_Runde & "'")
    
    sqlString = "select * from paare p1 where startkl='" & sk & "' and turniernr=" & T_Nr & " and (Anwesent_Status = 1 Or Anwesent_Status = 2)"
    sqlString = sqlString & " and not exists (select 1 from paare_rundenqualifikation pr where pr.rt_id=" & rs!RT_ID & " AND pr.tp_id=p1.tp_id)"
    Set rstpaare = dbs.OpenRecordset(sqlString)
    fill_Paare_rundenquali rstErste, rstpaare, rs!RT_ID
    
    ' bei geteilter End/Vorrunde die Paare in alle Runden aufnehmen
    If Me!n�chste_Runde.Column(3) <> "" Then
        Set rs = dbs.OpenRecordset("select * from Rundentab where startklasse = '" & sk & "' and turniernr = " & T_Nr & " and runde = '" & Me!n�chste_Runde.Column(3) & "';")
        sqlString = "select * from paare p1 where startkl='" & sk & "' and turniernr=" & T_Nr & " and (Anwesent_Status = 1 Or Anwesent_Status = 2)"
        sqlString = sqlString & " and not exists (select 1 from paare_rundenqualifikation pr where pr.rt_id=" & rs!RT_ID & " AND pr.tp_id=p1.tp_id)"
        Set rstpaare = dbs.OpenRecordset(sqlString)
        fill_Paare_rundenquali rstErste, rstpaare, rs!RT_ID
    End If
        
    '***** Mehrkampf ******
    If left(Me!n�chste_Runde, 4) = "MK_5" Then
        Set rs = dbs.OpenRecordset("SELECT * FROM Rundentab WHERE startklasse = '" & sk & "' AND turniernr = " & Me!T_Nr & " AND runde LIKE 'MK_*' AND runde<>'MK_5_TNZ'")
        rs.MoveFirst
        Do Until rs.EOF()
            sqlString = "select * from paare p1 where startkl='" & sk & "' and turniernr=" & T_Nr & " and (Anwesent_Status = 1 Or Anwesent_Status = 2)"
            sqlString = sqlString & " and not exists (select 1 from paare_rundenqualifikation pr where pr.rt_id=" & rs!RT_ID & " AND pr.tp_id=p1.tp_id)"
            Set rstpaare = dbs.OpenRecordset(sqlString)
            fill_Paare_rundenquali rstErste, rstpaare, rs!RT_ID
            
            rs.MoveNext
        Loop
    
    End If
    
    '********* HM V14.03 check ob Anzahl der T�nzer bei Formationen richtig eingetragen sind
    If InStr(1, sk, "F_") > 0 And rstpaare.RecordCount > 0 Then
        Dim AnzahlCheck As Formationswerte
        Dim isFault As Boolean
        rstpaare.MoveFirst
        AnzahlCheck = Faktor_Formation_Abzuege(sk)
        Do Until rstpaare.EOF
            If rstpaare!Anz_Taenzer < AnzahlCheck.min Or rstpaare!Anz_Taenzer > AnzahlCheck.max Or Nz(rstpaare!Anz_Taenzer) = "" Then
                MsgBox "Die Anzahl der T�nzer bei >" & rstpaare!Name_Team & "< stimmt nicht!", vbOKOnly
                isFault = True
            End If
            rstpaare.MoveNext
        Loop
        If isFault Then Exit Sub
    End If
       
    Me.Refresh
End Sub

Public Function fill_Paare_rundenquali(ziel, quelle, rt As Integer)
    ' �berz�hlige l�schen
    Dim db As Database
    Dim sqlcmd As String
    
    Set db = CurrentDb
    sqlcmd = "DELETE FROM Paare_Rundenqualifikation pr WHERE pr.rt_id=" & rt
    sqlcmd = sqlcmd & " and not exists (select 1 from Paare p where pr.tp_id=p.tp_id and p.anwesent_status>0)"
    db.Execute (sqlcmd)
    ' neue hinzuf�gen
    If quelle.RecordCount > 0 Then quelle.MoveFirst
    
    Do Until quelle.EOF()
        ziel.AddNew
        ziel!TP_ID = quelle!TP_ID
        ziel!RT_ID = rt
        ziel!Anwesend_Status = quelle!Anwesent_Status
        ziel!Verein_Name = quelle!Verein_Name
        ziel!Rundennummer = Null
        ziel.Update
        quelle.MoveNext
    Loop
    make_a_startlist rt
End Function

Private Sub Befehl20_Click()
On Error GoTo Err_Befehl20_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Monitor_Runden"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Befehl20_Click:
    Exit Sub

Err_Befehl20_Click:
    MsgBox err.Description
    Resume Exit_Befehl20_Click
    
End Sub

Private Sub Form_Resize()
    If Me.InsideHeight > 3000 Then
        Me![Paare Unterformular].Height = Me.InsideHeight - 1800
        Me![Paare_Rundenqualifikation Unterformular].Height = Me.InsideHeight - 1800
    End If
End Sub

Private Sub n�chste_Runde_Change()
    Paare_Rundenqualifikation_Unterformular.Requery
End Sub

Private Sub Startklasse_AfterUpdate()
    
    Me!n�chste_Runde = Null
    DoCmd.RepaintObject acForm, "Paare_in erste Runde nehmen"
    DoCmd.GoToRecord , "", acFirst
    Me.Refresh
    
End Sub
