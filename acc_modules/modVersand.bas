Option Compare Database
Option Explicit
    
    #If Win64 And VBA7 Then
         Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    #Else
        Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    #End If

    
Sub send_zeitplan(Turniernr)
    Dim re As Recordset
    Dim Zeitplan As String
    Dim zFileName As String
    Dim zFile
    
    Set re = CurrentDb().OpenRecordset("SELECT r.Rundenreihenfolge, Startzeit, Rundentext, Startklasse_text FROM Tanz_Runden t RIGHT JOIN (Rundentab r LEFT JOIN Startklasse s ON r.Startklasse = s.Startklasse) ON t.Runde = r.Runde ORDER BY r.Rundenreihenfolge;")
    zFileName = getBaseDir() & "Turnierleiterpaket\" & Turniernr & "_Zeitplan.csv"
    
    re.MoveFirst
    Zeitplan = """Uhrzeit"";""Runde"";""Startklasse""" & vbCrLf
    Do Until re.EOF
        If re!Rundenreihenfolge < 999 Then
            Zeitplan = Zeitplan & """" & Format(re!Startzeit, "hh:mm") & """;""" & re!Rundentext & """;""" & re!Startklasse_text & """" & vbCrLf
        End If
        re.MoveNext
    Loop
    Set zFile = file_handle(zFileName)
    zFile.writeline (Zeitplan)
    zFile.Close

    If OutlookInstalliert Then
        Dim objOutlook, objOLAtt, objOutMsg As Object
        Dim oApp As Object
    
        Set objOutlook = CreateObject("Outlook.Application")
        Set objOutMsg = objOutlook.CreateItem(0)
        With objOutMsg
            .To = "turnierueberwachung@drbv.de"
            .Subject = "Zeitplan " & Turniernr
            .body = "Zeitplan " & Turniernr
        End With
        Set objOLAtt = objOutMsg.Attachments.Add(zFileName)
        objOutMsg.Display
        'objOutMsg.Send
    
        Set objOutMsg = Nothing
        Set objOutlook = Nothing
    End If
    

End Sub

Sub Gen_Mail()
    Dim DefPath As String
    Dim retl As Integer
    Dim tName As String
    Dim FileToZip
    Dim empf As String
    Dim tur_ber As Boolean
    Dim ZipFileName
    Dim i As Integer
    
    DefPath = gen_Ordner(getBaseDir & "_Versand\") & get_TerNr
    
    If Info_Laufwerke(left(DefPath, 3)) Then
        ZipFileName = DefPath & "_Versand.zip"
        NewZip ZipFileName
        
        i = 1
        FileToZip = DefPath & "_Turnierbericht.rtf"
        DoCmd.OutputTo acOutputReport, "Turnierbericht", acFormatRTF, FileToZip, False, ""
        zip_file ZipFileName, FileToZip, i
        
        FileToZip = DefPath & "_Rangliste.xls"
        DoCmd.OutputTo acQuery, "Ergebnisliste_Text", "MicrosoftExcel(*.xls)", FileToZip, False, ""
        zip_file ZipFileName, FileToZip, i
        
        FileToZip = getBaseDir & get_TerNr & "_TDaten.mdb"
        zip_file ZipFileName, FileToZip, i
        
        FileToZip = DefPath & "_Ergebnisliste.txt"
        Call writeErgebnisliste(CStr(FileToZip))
        zip_file ZipFileName, FileToZip, i
        
        FileToZip = DefPath & "_Ergebnisliste.html"
        zip_file ZipFileName, FileToZip, i
        
        Select Case Forms![A-Programmübersicht]!Turnierausw.Column(8)
            Case "SL"
                empf = "breitensport@bwrrv.de"
                tur_ber = False
            Case Else
                empf = "turnierueberwachung@drbv.de"
                tur_ber = True
                FileToZip = DefPath & "_Abgegebene_Wertungen.csv"
                export_tabelle "Abgegebene_Wertungen", FileToZip
                zip_file ZipFileName, FileToZip, i
                
                FileToZip = DefPath & "_Paare.csv"
                DoCmd.TransferText acExportDelim, "Paare Exportspezifikation", "Paare", FileToZip, True
                zip_file ZipFileName, FileToZip, i
                
                FileToZip = DefPath & "_Majoritaet.csv"
                DoCmd.TransferText acExportDelim, "Majoritaet Exportspezifikation", "Majoritaet", FileToZip, True
                zip_file ZipFileName, FileToZip, i
                
                FileToZip = DefPath & "_Rundentab.csv"
                DoCmd.TransferText acExportDelim, "Rundentab Exportspezifikation", "Rundentab", FileToZip, True
                zip_file ZipFileName, FileToZip, i
                
                FileToZip = DefPath & "_Turnier.csv"
                DoCmd.TransferText acExportDelim, "Turnier Exportspezifikation", "Turnier", FileToZip, True
                zip_file ZipFileName, FileToZip, i
                
                FileToZip = DefPath & "_Turnierleitung.csv"
                DoCmd.TransferText acExportDelim, "Turnierleitung Exportspezifikation", "Turnierleitung", FileToZip, True
                zip_file ZipFileName, FileToZip, i
                
                FileToZip = DefPath & "_Wert_Richter.csv"
                DoCmd.TransferText acExportDelim, "Wert_Richter Exportspezifikation", "Wert_Richter", FileToZip, True
                zip_file ZipFileName, FileToZip, i
                
                FileToZip = DefPath & "_Paare_Rundenqualifikation.csv"
                DoCmd.TransferText acExportDelim, "Paare_Rundenqualifikation Exportspezifikation", "Paare_Rundenqualifikation", FileToZip, True
                zip_file ZipFileName, FileToZip, i
                
                FileToZip = DefPath & "_Auswertung.csv"
                DoCmd.TransferText acExportDelim, "Auswertung Exportspezifikation", "Auswertung", FileToZip, True
                zip_file ZipFileName, FileToZip, i
        
'                FileToZip = DefPath & "_Mehrkampfplatzierungen.csv"
'                DoCmd.TransferText acExportDelim, "Mehrkampfauswertung", "view_Mehrkampfauswertung", FileToZip, True
'                zip_file ZipFileName, FileToZip, i
'
        End Select
        
        
        If OutlookInstalliert Then
            Dim betreff As String
            Dim text As String
            
            'Turnierunterlagen an die Turnierüberwachung
            betreff = Forms![A-Programmübersicht]!Turnierbez & " _ " & Forms![A-Programmübersicht]!Turnierveranstalter & " _ " & Forms![A-Programmübersicht]!Tur_Datum
            text = "Hallo," & vbCrLf & vbCrLf & "es wurde " & db_Ver() & " verwendet." & _
                    vbCrLf & vbCrLf & "Gruß "
                                   
            send_outlook empf, "", betreff, text, ZipFileName
            
            If tur_ber Then
                'Ergebnisliste an Mailliste
                empf = "geschaeftsstelle@drbv.de"
                text = "Hallo," & vbCrLf & vbCrLf & "hier der Turnierbericht " & Forms![A-Programmübersicht]![Turnierbez] & _
                        vbCrLf & vbCrLf & "Gruß "
                send_outlook empf, "", betreff, text, DefPath & "_Ergebnisliste.txt;" & DefPath & "_Turnierbericht.rtf"
            End If
        
        End If
    Else
        MsgBox "Erstellen einer ZIP-Datei funktioniert nicht auf einem Wechseldatenträger!"
    End If
End Sub

Sub send_outlook(empf, bcc, betreff, text, anhang)
    Dim objOutlook, objOLAtt, objOutMsg As Object
    Dim oApp As Object
    Dim i
    Dim vars
    
    Set objOutlook = CreateObject("Outlook.Application")
    'Turnierunterlagen an die Turnierüberwachung
    Set objOutMsg = objOutlook.CreateItem(0)
    With objOutMsg
        .To = empf
        If bcc <> "" Then .bcc = bcc
        .Subject = betreff
        .body = text
    End With
    If anhang <> "" Then
        vars = Split(anhang, ";")
        For Each i In vars
            Set objOLAtt = objOutMsg.Attachments.Add(i)
        Next
    End If
    objOutMsg.Display
    Set objOutMsg = Nothing
    Set objOutlook = Nothing
   
End Sub

Sub NewZip(sPath)
    'Create empty Zip File
    If Len(Dir(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub

Function zip_file(ZipFileName, fName, i)
    ' copy File in ZIP und warte
    Dim oApp As Object
    Set oApp = CreateObject("Shell.Application")
    
    If Len(Dir(fName)) > 0 Then
        oApp.Namespace(ZipFileName).CopyHere fName
        Sleep 1000
        Do Until oApp.Namespace(ZipFileName).items.Count = i
        Loop
        i = i + 1
    Else
        MsgBox fName & " wurde noch nicht erzeugt", vbOKOnly
    End If

End Function

Function OutlookInstalliert()
    ' testen ob Outlook installiert ist
    Dim olapp As Object
    On Error Resume Next
    OutlookInstalliert = False
    Set olapp = GetObject(, "Outlook.Application")
    If olapp Is Nothing Then _
        Set olapp = CreateObject("Outlook.Application")
    On Error GoTo 0
    
    If Not olapp Is Nothing Then
        OutlookInstalliert = True
    End If
    Set olapp = Nothing
End Function

Function Info_Laufwerke(pfad)
    ' bei Wechseldatenträgern funktioniert copy in ZIP nicht
    On Error Resume Next
    Dim fso, lw
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set lw = fso.GetDrive(pfad)

     If lw.DriveType = 2 Then Info_Laufwerke = True
End Function

Function export_tabelle(tbl, FileToZip)
    Dim db As Database
    Dim re As Recordset
    Dim out As Object
    Dim fld As Field
    Dim flds()
    Dim line As String
    Dim fld_count, i As Integer
    
    Set db = CurrentDb
    fld_count = 1
    Set re = db.OpenRecordset(tbl)
    For Each fld In db(tbl).Fields
        ReDim Preserve flds(fld_count)
        flds(fld_count) = fld.Name
        line = line & """" & fld.Name & """;"
        fld_count = fld_count + 1
    Next
    line = left(line, Len(line) - 1) & vbCrLf
    If re.RecordCount > 0 Then re.MoveFirst
    
    Do Until re.EOF
        For i = 1 To fld_count - 1
            If InStr(1, flds(i), "_Text") > 0 Then
                line = line & """" & re(flds(i)) & """;"
            Else
                line = line & re(flds(i)) & ";"
            End If
        Next
        line = left(line, Len(line) - 1) & vbCrLf
        re.MoveNext
    Loop
    Set out = file_handle(FileToZip)
    out.writeline (line)
    out.Close
End Function

Sub alle_Paare_anschreiben()
    Dim fName As String
    Dim line As String
    Dim mails As String
    Dim empf As String
    Dim vars
    Dim indexMail As Integer
    
    OpenFile.lpstrFilter = "Turnierdatenbanken (*.csv)" & Chr(0) & "*.csv" & Chr(0)
    OpenFile.lpstrInitialDir = "C:\"
    fName = get_Filename(0)

    Open fName For Input As #1
    Line Input #1, line
    line = del_kochkomma(line)
    vars = Split(line, ";")
    Do Until EOF(1)
        Line Input #1, line
        line = del_kochkomma(line)
        vars = Split(line, ";")
        mails = mails & vars(8) & "; "
    Loop
    Close #1
    empf = DLookup("Lizenznr", "Turnierleitung", "Art = 'TL'")
    empf = DLookup("[e-mail]", "TLP_OFFIZIELLE", "Lizenzn = '" & empf & "'")
    send_outlook empf, mails, "Betreff", "Text", ""
End Sub

