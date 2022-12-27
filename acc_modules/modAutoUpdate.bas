Option Compare Database

    #If Win64 And VBA7 Then
        Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As LongPtr, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As LongPtr, ByVal lpfnCB As LongPtr) As Long
        Private Declare PtrSafe Function DeleteUrlCacheEntry Lib "Wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
    #Else
        Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
        Private Declare Function DeleteUrlCacheEntry Lib "Wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
    #End If

Function DirExists(fileName As String) As Boolean
    DirExists = (Len(Dir(fileName, vbDirectory)) <> 0)
End Function

Function tes_dld()
    updateTLP False, True
End Function

Public Function updateTLP(dl_data, rmldg)
    'Erst nachfragen, ob im Internet nach einem Update gesucht werden soll
    Dim result As Integer
    Dim Version As String
    If get_properties("update_TLP") = True Then
        result = MsgBox("Soll das Turnierleiterpaket aktualisiert werden?", vbYesNo)
    End If
    updateTLP = 0
    If (result = vbYes) Then
        Dim dateien, tbls As Variant
        Dim llRetVal As Long
        Dim downloadTP As String
        Dim fMsg As String
        Dim destDir As String
        Dim i As Integer
        Dim cnt As Integer
        
        Select Case get_properties("LAENDER_VERSION")
            Case "SL"
                dateien = Array("Termine-Start-Daten.txt", "WR-TL-Start-Daten.txt", "DRBV-Akrotabelle.txt")
                tbls = Array("TLP_TERMINE", "TLP_OFFIZIELLE", "Akrobatiken")
        
            Case Else
                dateien = Array("BW-Start-Daten.txt", "RR-Start-Daten-Paare.txt", "Formationen.txt", "WR-TL-Start-Daten.txt", "Termine-Start-Daten.txt", "DRBV-Akrotabelle-12P.txt")
                tbls = Array("TLP_BW_PAARE", "TLP_RR_PAARE", "TLP_FORMATIONEN", "TLP_OFFIZIELLE", "TLP_TERMINE", "MSys__Akrobatiken")
        End Select
        
        
        destDir = getBaseDir() & "Turnierleiterpaket\"
        gen_Ordner destDir
        
        If dl_data Then
            For i = 0 To UBound(dateien)
                downloadTP = destDir & dateien(i)
                If get_url_to_file("http://www.drbv.de/cms/images/Download/TurnierProgramm/" & dateien(i), downloadTP) = 0 Then
                    cnt = cnt + 1
                End If
            Next
            If cnt = UBound(dateien) + 1 Then
                fMsg = "Das Turnierleiterpaket wurde erfolgreich aktualisiert."
            Else
                fMsg = "Es konnten nicht alle Dateien vom DRBV-Server geladen werden."
            End If
        
            If cnt <> 0 And dl_data Then   'nichts heruntergeladen
                ' Check neues TLP
                aktVersion = Replace(db_Ver, "-", ".")
                Version = get_url_to_string("http://www.drbv.de/cms/index.php/aktivenportal/downloads/turnierprogramm")
                off = InStr(1, Version, "/cms/images/Download/TurnierProgramm/TLP-V20/")
                If off <> 0 Then
                    Version = Replace(Mid(Version, off + 53, Len(aktVersion)), "-", ".")
                    If Version - aktVersion > 0 Then
                        If Len(fMsg) > 1 Then fMsg = vbCrLf + fMsg
                        fMsg = "Es gibt eine neue Version (" & Version & ") des Turnierprogramms." & fMsg
                    End If
                End If
            End If
        End If
        
        cnt = 0
        For i = 0 To UBound(dateien)
            llRetVal = update_drbv_tables(tbls(i), dateien(i), destDir)
            cnt = cnt + llRetVal
        Next i
        If Len(fMsg) > 1 Then fMsg = fMsg + vbCrLf
        If rmldg = True Then
            MsgBox fMsg & "Es wurden " & cnt & " Tabellen aktualisiert"
        End If
        If cnt > 0 Then updateTLP = cnt
    End If
End Function

Function get_url_to_file(file_URL, file_dest)
    On Error Resume Next
    Dim lRet As Integer
    lRet = DeleteUrlCacheEntry(file_URL)
    Kill file_dest
    get_url_to_file = URLDownloadToFile(0, file_URL, file_dest, 0, 0)
    
End Function

Function get_url_to_string_check(url)
    If get_properties("EWS") = "EWS3" Then
        get_url_to_string_check = get_url_to_string(url)
    End If
End Function

Function get_url_to_string(url)
    On Error GoTo exit_sub
    Dim winHttpReq As Object

    Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    winHttpReq.Open "GET", url, False
    winHttpReq.send
    get_url_to_string = winHttpReq.responseText
exit_sub:
End Function

Function post_url_string()
    Dim winHttpReq As Object
    Dim url As String
    Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    url = "http://192.168.1.101/login"
    winHttpReq.Open "POST", url, False
    winHttpReq.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    winHttpReq.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    winHttpReq.send ("wr_id=4&passwort=1234")
    Debug.Print winHttpReq.responseText
    

End Function

Public Function update_drbv_tables(tbl, fName, destDir)
    Dim db As Database
    Dim re As Recordset
    Dim impo As String
    Dim sql As String
    Dim strZeile As String
    Dim he, da As Variant
    Dim i, st, en As Integer
        
    Set db = CurrentDb
    If InStr(fName, "Anmeldung_2.txt") = 0 Then
        sql = "DELETE FROM " & tbl
        db.Execute sql
    End If
    Set re = db.OpenRecordset(tbl, DB_OPEN_DYNASET)
   
    If Len(Dir(destDir & fName)) <> 0 Then
        update_drbv_tables = 1
        Open destDir & fName For Input As #1
        Line Input #1, strZeile
        strZeile = del_kochkomma(strZeile)
        he = Split(strZeile, ";")
        Do While Not EOF(1)
            Line Input #1, strZeile
            strZeile = del_kochkomma(strZeile)
            da = Split(strZeile, ";")
            If InStr(fName, "Anmeldung_2.txt") > 0 Then
                If da(9) <> "" Then
                    re.FindFirst he(9) & " = " & da(9)
                Else
                    re.FindFirst he(10) & " = " & da(10)
                End If
                If re.NoMatch Then
                    MsgBox "Diese Startkarte existiert nicht"
                    update_drbv_tables = 0
                    Exit Do
                Else
                    re.Edit
                End If
                ' Name und Verein aktualisieren oder neu schreiben
                st = 12
                en = 71
            Else
                re.AddNew
                st = 0
                en = UBound(he)
            End If
            For i = st To en
                If da(i) <> "" Then re(he(i)) = Nz(da(i))
            Next i
            re.Update
        Loop
        Close #1
    End If
    Set re = Nothing
End Function

Function del_kochkomma(str)
    If left(str, 1) = Chr(34) Then str = Mid(str, 2)
    If Right(str, 1) = "," Then str = Mid(str, 1, Len(str) - 1)
    If Right(str, 1) = Chr(34) Then str = Mid(str, 1, Len(str) - 1)
    str = Replace(str, Chr(34) & ";" & Chr(34), ";")
    del_kochkomma = str
End Function

Private Sub Endrunden_Musik_herunterladen()
    Dim db As Database
    Dim re As Recordset
    Dim vars
    Dim pfad As String
    Dim file_URL As String
    Dim dest_file As String
    Dim retl As Integer
    
    Set db = CurrentDb
    Set re = db.OpenRecordset("Musik", DB_OPEN_DYNASET)
    
    pfad = gen_Ordner(getBaseDir() & "Turnierleiterpaket\" & get_TerNr() & "_Musik")
    re.MoveFirst
    Do Until re.EOF
        If Nz(re!lieder) <> "" Then
            gen_Ordner (pfad & "\" & re!Startkl)
            gen_Ordner (pfad & "\" & re!Startkl & "\" & re!pfad)
            vars = Split(re!lieder, "_")
            dest_file = pfad & "\" & re!Startkl & "\" & re!pfad & "\" & vars(UBound(vars)) & "_" & re!Musik_Name & ".mp3"
            file_URL = "http://www.drbv.de/turniermusik/index.php?file=" & re!lieder '& ".mp3"
            retl = get_url_to_file(file_URL, dest_file)
        End If
        re.MoveNext
    Loop

End Sub

Private Sub DRBV_Musik_herunterladen()
    Dim vars
    Dim pfad As String
    Dim file_URL As String
    Dim dest_file As String
    Dim retl As Integer
    Const takte = 51
    
    pfad = "C:\DRBV-V16\Musik\turniermusik"
'    pfad = gen_Ordner(getBaseDir() & "Turnierleiterpaket\" & get_TerNr() & "_Musik")
    If Len(Dir(pfad & ".csv")) <> 0 Then
        Open pfad & fName & ".csv" For Input As #1
        Line Input #1, strZeile
        Do While Not EOF(1)
            Line Input #1, strZeile
            If strZeile <> "" Then
                strZeile = del_kochkomma(strZeile)
                da = Split(strZeile, ";")
                If da(5) = takte And left(da(2), 6) = "boogie" And da(8) = "swing" Then
                    gen_Ordner (getBaseDir() & "Musik")
                    dest_file = gen_Ordner(getBaseDir() & "Musik" & "\" & da(5) & da(8)) & "\" & da(4) & " - " & da(3) & ".mp3"
                    If InStr(da(1), "&") Then
'                        Debug.Print
                        da(1) = Replace(da(1), "&", "&teil2=")
                    End If
                    file_URL = "http://www.drbv.de/turniermusik/musikdb.php?pfad=/" & da(2) & "&file=" & da(1)

                    retl = get_url_to_file(file_URL, dest_file)
'                    Pause 1

                End If
            End If
        Loop
        Close #1
    End If
    Debug.Print "Fertig " & takte
End Sub

Private Sub Musik_prüfen()
    Dim db As Database
    Dim re As Recordset
    Dim vars
    Dim pfad As String
    Dim retl As Long
    
    Set db = CurrentDb
    Set re = db.OpenRecordset("Musik", DB_OPEN_DYNASET)
    
    pfad = gen_Ordner(getBaseDir() & "Turnierleiterpaket\" & get_TerNr() & "_Musik")
    re.MoveFirst
    Do Until re.EOF
        If Nz(re!lieder) <> "" Then
            vars = Split(re!lieder, "_")
            dest_file = pfad & "\" & re!Startkl & "\" & re!pfad & "\" & vars(UBound(vars)) & "_" & re!Musik_Name & ".mp3"
            retl = FileLen(dest_file)
            If retl < 1000000 Then
                MsgBox "Das Lied " & vars(UBound(vars)) & "_" & re!Musik_Name & ".mp3 scheint zu kurz zu sein!"""
            End If
        End If
        re.MoveNext
    Loop

End Sub
