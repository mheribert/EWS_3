Option Compare Database
Option Explicit

Function fill_observer_verstoesse(line, re, ppr, RT_nr, sei_1, sei_2, WR_ID)
    Dim db As Database
    Dim varBookmark As Variant
    Dim vars, t
    Dim color As String
    Dim rde As Recordset
    Dim verstoss
    Dim i, seite As Integer
    Dim anzahl_paare As Integer
    varBookmark = re.Bookmark
    verstoss = fill_verstoss()
    For anzahl_paare = 1 To IIf(ppr, 1, 2)
        For i = 0 To 8
            verstoss(i, 1) = Null
        Next
        Set db = CurrentDb
        Set rde = db.OpenRecordset("Select * FROM rundentab WHERE RT_ID = " & RT_nr & ";", DB_OPEN_DYNASET)
        
        Set rde = db.OpenRecordset("SELECT Cgi_Input FROM Rundentab INNER JOIN (Auswertung INNER JOIN Paare_Rundenqualifikation ON Auswertung.PR_ID = Paare_Rundenqualifikation.PR_ID) ON Rundentab.RT_ID = Paare_Rundenqualifikation.RT_ID WHERE ((Auswertung.WR_ID=" & WR_ID & ") AND (Paare_Rundenqualifikation.TP_ID=" & re!TP_ID & ") AND (Rundentab.Startklasse=""" & re!Startkl & """) AND (Rundentab.Rundenreihenfolge<" & rde!Rundenreihenfolge & "));", DB_OPEN_DYNASET)
        If rde.RecordCount <> 0 Then
            rde.MoveFirst
            Do Until rde.EOF
                Set vars = zerlege(rde!Cgi_Input)
                If CSng(vars.Item("PR_ID1")) = re!TP_ID Then
                    seite = 1
                ElseIf CSng(vars.Item("PR_ID2")) = re!TP_ID Then
                    seite = 2
                Else
                    MsgBox "Zuordnung Paar und Verstösse stimmt nicht"
                End If
                For t = 0 To 7
                    If vars.Item("w" & verstoss(t, 0) & seite) <> "" Then
                        If Val(vars.Item("w" & verstoss(t, 0) & seite)) >= CSng(Nz(verstoss(t, 1))) Then
                            verstoss(t, 1) = CSng(vars.Item("w" & verstoss(t, 0) & seite))
                        End If
                    End If
                Next
                rde.MoveNext
            Loop
        End If
        For t = 0 To 7
            Select Case verstoss(t, 1)
                Case 0
                    color = "yell"
                Case 30
                    color = "red"
                Case 100
                    color = "black"
                Case Else
                    color = "leer"
            End Select
            ' hier wird das kriterium in die HTML--Seite gebracht abhängig von Tausch
            line = Replace(line, "x__" & verstoss(t, 0) & IIf(anzahl_paare = 1, sei_1, sei_2), color)
        Next
        re.MoveNext
    Next
    fill_observer_verstoesse = line
    re.Bookmark = varBookmark
End Function

Public Function fill_verstoss()
Dim vArray(8, 1)
    vArray(0, 0) = "sidebysidevw"
    vArray(1, 0) = "akrovw"
    vArray(2, 0) = "highlightvw"
    vArray(3, 0) = "juniorvw"
    vArray(4, 0) = "kleidungvw"
    vArray(5, 0) = "tanzbereichvw"
    vArray(6, 0) = "tanzzeitvw"
    vArray(7, 0) = "aufrufvw"
    fill_verstoss = vArray
End Function


