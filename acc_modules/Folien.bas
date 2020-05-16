Option Compare Database
Option Explicit
    Dim v_Font As String
    Dim v_Size As Integer
    Dim v_Color As Long
    Dim v_Datei As Boolean
    Dim v_Suffix As String
    Dim v_Pfad As String
    Public FolienMaster As String
    
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Public Declare Sub wlib_AccChooseColor Lib "msaccess.exe" Alias "#53" (ByVal hwnd As Long, rgb As Long)
    
    Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
    Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
    Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
    Const KEY_READ = &H20019
    Const ERROR_MORE_DATA = 234

Public Sub gen_Vorstellung() ' aller Paare nach Verein
    Dim oPPTPres As Object
    Dim oPPTsli As Object
    Dim oPPTTBox As Object
    Dim db As Database
    Dim re As Recordset
    Dim tex As String
    Dim cou As Integer
    
    Set db = CurrentDb
    Set re = db.OpenRecordset("SELECT View_Paare.Verein_Name, IIf([Name_Team] Is Not Null, [Name_Team],[Da_Vorname] & "" "" & [Da_Nachname] & "" - "" & [He_Vorname] & "" "" & [He_Nachname]) AS VollerName, View_Paare.Startklasse_text, View_Paare.Startnr, Startklasse.isTeam FROM Startklasse INNER JOIN View_Paare ON Startklasse.Startklasse = View_Paare.Startkl WHERE (((View_Paare.Anwesent_Status)>0)) ORDER BY View_Paare.Verein_Name, Startklasse.isTeam, View_Paare.Startnr, IIf([Name_Team] Is Not Null,""  "" & [Name_Team],[Da_Vorname] & "" "" & [Da_Nachname] & "" - "" & [He_Vorname] & "" "" & [He_Nachname]);")

    tex = open_Master()
    If tex = "" Then Exit Sub        ' Check ob Folienmaster vorhanden
    Set oPPTPres = open_Pres(tex)

    Call erste_Folie(oPPTPres.Slides(1), "Vorstellung der Vereine")
    re.MoveFirst
    
    Do Until re.EOF
        If re!Verein_Name <> tex Or cou = 13 Then
            tex = re!Verein_Name
            Set oPPTsli = oPPTPres.Slides.Add(oPPTPres.Slides.Count + 1, 12)
            Set oPPTTBox = oPPTsli.Shapes.AddTextbox(1, 30, 180, 655, 130)
            schrZeile oPPTTBox, 1, re!Verein_Name, 30, True, 1
            oPPTTBox.TextFrame.WordWrap = False
            If oPPTTBox.Width > 670 Then red_Width oPPTTBox, 1
            Set oPPTTBox = oPPTsli.Shapes.AddTextbox(1, 30, 225, 660, 230)
            oPPTTBox.TextFrame.Ruler.TabStops.Add 3, 33     'ppTabStopRight
            oPPTTBox.TextFrame.Ruler.TabStops.Add 1, 45     'ppTabStopLeft
           cou = 1
        End If
        schrZeile oPPTTBox, 1, Chr(9) & re!Startnr & Chr(9), 20, False, 1
        schrZeile oPPTTBox, 1, re!VollerName & Chr(13), 20, re!isTeam, 1, re!isTeam
        cou = cou + 1
        re.MoveNext
    Loop
    Call addRRClogo(oPPTPres)
    Call save_Pres(oPPTPres, "_Vorstellung der Paare nach Verein")
    oPPTPres.Application.Quit
    Set oPPTTBox = Nothing
    Set oPPTsli = Nothing
    Set oPPTPres = Nothing
    
End Sub
 
Public Sub gen_Folien(re As Recordset, st_klasse As String, Runde As String, runde_id As String) 'Rundeneinteilung
    Dim oPPTPres As Object
    Dim oPPTsli As Object
    Dim oPPTTBox As Object
    Dim yPos As Single
    Dim tex As String
    Dim max_Runde As Integer
    Dim Runden_anz As Integer
    
    re.MoveFirst
    If IsNull(re!Rundennummer) Then
        MsgBox "Scheinbar wurde die Runde noch nicht ausgelost."
    Else
        re.MoveLast
        tex = open_Master()
        If tex = "" Then Exit Sub        ' Check ob Folienmaster vorhanden
        Set oPPTPres = open_Pres(tex)
    
        Call erste_Folie(oPPTPres.Slides(1), st_klasse & Chr(13) & Runde)
        max_Runde = Int(re.RecordCount / re!Anz_Paare) + re.RecordCount Mod re!Anz_Paare
        re.MoveFirst
        Do Until re.EOF
            If Nz(re!Rundennummer) <> Runden_anz Then
                Runden_anz = Runden_anz + 1
                Set oPPTsli = oPPTPres.Slides.Add(oPPTPres.Slides.Count + 1, 12)
                Set oPPTTBox = oPPTsli.Shapes.AddTextbox(1, 30, 180, 550, 130)
                    schrZeile oPPTTBox, 1, st_klasse & " " & Runde, 30, False, 1
                    yPos = IIf(oPPTTBox.Height > 70, 270, 240)
                Set oPPTTBox = oPPTsli.Shapes.AddTextbox(1, 520, 180, 170, 130)
                    schrZeile oPPTTBox, 1, Runden_anz & "/" & max_Runde, 30, False, 3
            End If
            Set oPPTTBox = oPPTsli.Shapes.AddTextbox(1, 30, yPos, 630, 130)
                oPPTTBox.TextFrame.WordWrap = False
                If IsNull(re!Name_Team) Then    'Dame u Herr
                    schrZeile oPPTTBox, 1, re!Startnr & Chr(9) & get_Dame(re) & Chr(13), 37 + v_Size, True, 1
                    schrZeile oPPTTBox, 2, Chr(9) & get_Herr(re) & Chr(13), 37 + v_Size, True, 1
                Else                            'Verein
                    schrZeile oPPTTBox, 1, re!Startnr & Chr(9) & re!Name_Team & Chr(13) & Chr(13), 37 + v_Size, True, 1
                    If oPPTTBox.Width > 670 Then red_Width oPPTTBox, 1
                End If
                schrZeile oPPTTBox, 3, Chr(9) & re!Verein_Name, 27 + v_Size, False, 1, True
                If oPPTTBox.Width > 670 Then red_Width oPPTTBox, 3
                yPos = yPos + oPPTTBox.Height + 5
            
            re.MoveNext
        Loop
        
        Call addRRClogo(oPPTPres)
        Call save_Pres(oPPTPres, IIf(v_Datei, Format(runde_id, "00") & "_", "") & st_klasse & "_" & Runde & "_" & "Rundeneinteilung")
        oPPTPres.Application.Quit
        Set oPPTTBox = Nothing
        Set oPPTsli = Nothing
        Set oPPTPres = Nothing
    End If
End Sub

Sub gen_Ergebnisliste(re As Recordset, st_klasse As String, Runde As String)
    Dim oPPTPres As Object
    Dim oPPTsli As Object
    Dim oPPTTBox As Object
    Dim tex As String
    Dim i As Integer
    
    tex = open_Master()
    If tex = "" Then Exit Sub        ' Check ob Folienmaster vorhanden
    Set oPPTPres = open_Pres(tex)
    Call erste_Folie(oPPTPres.Slides(1), st_klasse & Chr(13) & "Siegerehrung")
    
    re.MoveLast
    Set oPPTsli = oPPTPres.Slides.Add(oPPTPres.Slides.Count + 1, 12)
    Set oPPTTBox = oPPTsli.Shapes.AddTextbox(1, 30, 180, 646, 230)
        schrZeile oPPTTBox, 1, "Ergebnis " & st_klasse, 30, False, 1
        
    Set oPPTTBox = oPPTsli.Shapes.AddTextbox(1, 30, 240, 660, 230)
    Set oPPTTBox = oPPTsli.Shapes.AddTextbox(1, 30, 240, 660, 230)
        oPPTTBox.TextFrame.Ruler.TabStops.Add 3, 14     'ppTabStopRight
        oPPTTBox.TextFrame.Ruler.TabStops.Add 3, 57     'ppTabStopLeft
        oPPTTBox.TextFrame.Ruler.TabStops.Add 1, 68     'ppTabStopLeft

        re.MoveFirst
            For i = 1 To re.RecordCount
                schrZeile oPPTTBox, Val(i), Chr(9) & re!Platz & Chr(9) & getPaarname(re) & Chr(13), 20 + v_Size, IIf(i = 1, True, False), 1
                re.MoveNext
            Next i

        oPPTTBox.TextFrame.TextRange.ParagraphFormat.LineRuleAfter = False
        oPPTTBox.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 12
        oPPTTBox.TextFrame.WordWrap = False
        oPPTTBox.AnimationSettings.EntryEffect = 3329 ' ppEffectFlyFromLeft
        oPPTTBox.AnimationSettings.TextLevelEffect = 1 ' ppAnimateByFirstLevel
        oPPTTBox.AnimationSettings.AnimateTextInReverse = True
        If oPPTTBox.Width > 670 Then red_Width oPPTTBox, 0

    re.MoveFirst
    Call addRRClogo(oPPTPres)
    Call save_Pres(oPPTPres, get_sieger(re!Startklasse) & st_klasse & "_Siegerehrung")
    
    oPPTPres.Application.Quit
    Set oPPTTBox = Nothing
    Set oPPTsli = Nothing
    Set oPPTPres = Nothing

End Sub

Sub gen_NächsteRunde(re As Recordset, st_klasse As String, Runde As String, runde_id)    'Qualifikation für nächste Runde
    Dim oPPTPres As Object
    Dim oPPTsli As Object
    Dim oPPTTBox As Object
    Dim tex As String
    Dim i As Integer
    
    tex = open_Master()
    If tex = "" Then Exit Sub        ' Check ob Folienmaster vorhanden
    Set oPPTPres = open_Pres(tex)
    
    Call erste_Folie(oPPTPres.Slides(1), " ")
    re.MoveLast
    Set oPPTsli = oPPTPres.Slides.Add(oPPTPres.Slides.Count + 1, 12)
    Set oPPTTBox = oPPTsli.Shapes.AddTextbox(1, 30, 180, 660, 230)
        schrZeile oPPTTBox, 1, "Qualifikation für ", 40, True, 2
        schrZeile oPPTTBox, 2, Chr(13) & st_klasse & " " & Runde & " ", 40, True, 2
   
    Set oPPTTBox = oPPTsli.Shapes.AddTextbox(1, 40, 320, 640, 230)
        re.MoveFirst
            For i = 1 To re.RecordCount
                schrZeile oPPTTBox, Val(i), re!Startnr & ", ", 40 + v_Size, False, 2
                re.MoveNext
            Next i
        oPPTTBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2  'ppAlignCenter
        oPPTTBox.TextFrame.TextRange.ParagraphFormat.LineRuleAfter = False
        oPPTTBox.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 12
        oPPTTBox.TextFrame.WordWrap = True
    
    Call addRRClogo(oPPTPres)
    Call save_Pres(oPPTPres, IIf(v_Datei, Format(runde_id, "00") & "__", "") & st_klasse & "_" & "Qualifikation für " & Runde)
    
    oPPTPres.Application.Quit
    Set oPPTTBox = Nothing
    Set oPPTsli = Nothing
    Set oPPTPres = Nothing

End Sub

'1  ppAlignLeft
'2  ppAlignCenter
'3  ppAlignRight
Private Function schrZeile(tBox, zl As Integer, atext As String, gr As Integer, fett As Boolean, orient As Integer, Optional kursiv As Boolean)
    Dim sha As Object
    Dim ln As Integer
    With tBox.TextFrame.TextRange
        ln = .Length + 1
        .InsertAfter atext
        .font.color = v_Color
        .font.Name = v_Font
        .ParagraphFormat.Alignment = orient
        Set sha = .Characters(ln, .Length - ln + 1)
        sha.font.Size = gr
        sha.font.Bold = fett
        sha.font.Italic = kursiv
    End With
End Function

Function red_Width(tBox, ln)
    Dim st As Single
    With tBox.TextFrame.TextRange.Paragraphs(ln).font
    st = .Size
    For st = 1 To 24
        .Size = .Size - 0.5
        If tBox.Width < 670 Then Exit For
    Next
    End With
End Function

Function open_Master()
    If Len(Dir(getBaseDir & "FolienMaster.ppt")) > 0 Then
        open_Master = getBaseDir & "FolienMaster.ppt"
    Else
        MsgBox "Folienmaster fehlt!"
    End If
End Function

Sub erste_Folie(sli, text As String)
    Dim oPPTbox As Object
    Set oPPTbox = sli.Shapes.AddTextbox(1, 30, 220, 660, 130)
    schrZeile oPPTbox, 1, text, 60, True, 2
End Sub

Sub addRRClogo(oPPTPres)
    oPPTPres.Slides.Add oPPTPres.Slides.Count + 1, 12
    If Len(Dir(getBaseDir & "Logo.jpg")) > 0 Then
        oPPTPres.Slides(oPPTPres.Slides.Count).Shapes.AddPicture getBaseDir & "logo.jpg", False, True, 220, 230, 300, 250
    End If
End Sub

Function get_sieger(st_kl)
    Dim db As Database
    Dim re As Recordset
    Set db = CurrentDb
    Set re = db.OpenRecordset("SELECT Rundenreihenfolge, Startklasse FROM rundentab WHERE (rundentab.runde=""Sieger"" AND rundentab.Turniernr= " & [Forms]![A-Programmübersicht]![Akt_Turnier] & ") ORDER BY Rundenreihenfolge;", DB_OPEN_DYNASET)
    If re.EOF Then
        get_sieger = ""
    Else
        re.MoveLast
        If re.RecordCount > 1 Then
            re.FindFirst "startklasse  = '" & st_kl & "'"
            If re.NoMatch Then
                get_sieger = ""
                Exit Function
            End If
        End If
        get_sieger = IIf(v_Datei, Format(re!Rundenreihenfolge, "00") & "_", "")
    End If
End Function

Sub save_Pres(oPPTPres, fName)
    oPPTPres.SaveAs v_Pfad & fName & v_Suffix
End Sub

Public Function getPaarname(re As Recordset) As String
    If IsNull(re!Name_Team) Then
        getPaarname = re!Startnr & Chr(9) & left(re!Da_Vorname, 1) & "." & re!Da_NAchname & " - " & left(re!He_Vorname, 1) & "." & re!He_Nachname & " / " & re!Verein_Name
    Else
        getPaarname = re!Startnr & Chr(9) & re!Name_Team & " / " & re!Verein_Name
    End If
End Function

Function get_Dame(re As Recordset)
    get_Dame = re!Da_Vorname & " " & re!Da_NAchname
End Function

Function get_Herr(re As Recordset)
    get_Herr = re!He_Vorname & " " & re!He_Nachname
End Function

Public Function open_Pres(f_Pres As String)
    Dim oPPTApp As Object
    Dim Pres As Object
    Dim db As Database
    Dim re As Recordset
    
    Set db = CurrentDb
    Set re = db.OpenRecordset("SELECT * FROM Turnier WHERE Turniernum = " & get_aktTNr)
    v_Font = re!PPT_Font
    v_Size = re!PPT_Size
    v_Color = re!PPT_Color
    v_Datei = re!PPT_Datei
    v_Suffix = re!PPT_Suffix
    If Nz(re!PPT_Pfad) <> "" Then
        If Len(Dir(re!PPT_Pfad, vbDirectory)) > 0 Then
            v_Pfad = IIf(Right(re!PPT_Pfad, 1) = "\", re!PPT_Pfad, re!PPT_Pfad & "\")
        Else
            MsgBox "Der angegebene Ordner existiert nicht!" & Chr(13) & Chr(13) & "Es wird der Standartpfad gesetzt"
            v_Pfad = gen_Ordner(left(getBaseDir, 2) & "\Foliengenerator") & "\"
        End If
    Else
        v_Pfad = gen_Ordner(left(getBaseDir, 2) & "\Foliengenerator") & "\"
    End If

    Set oPPTApp = CreateObject("PowerPoint.Application")
    oPPTApp.Visible = True
    Set open_Pres = oPPTApp.Presentations.Open(f_Pres)
    
    Set re = Nothing
    Set db = Nothing
    Set oPPTApp = Nothing
End Function

Public Function EnumRegistryValues(ByVal hKey As Long, ByVal keyname As String) 'As Collection
    Dim handle As Long
    Dim index As Long
    Dim valueType As Long
    Dim Name As String
    Dim nameLen As Long
    Dim retVal As Long
    Dim merk, font As String
    Dim db As Database
    Set db = CurrentDb
    Dim re As Recordset
    Call db.Execute("DELETE FROM Show")
    Set re = db.OpenRecordset("Show")
    
    If Len(keyname) Then
        If RegOpenKeyEx(hKey, keyname, 0, KEY_READ, handle) Then Exit Function
        hKey = handle
    End If
    
    Do
        nameLen = 260
        Name = Space$(nameLen)
        ReDim resBinary(0 To 4096 - 1) As Byte
        retVal = RegEnumValue(hKey, index, Name, nameLen, ByVal 0&, valueType, resBinary(0), 4096)
        
        If retVal = ERROR_MORE_DATA Then
            ReDim resBinary(0 To 4096 - 1) As Byte
            retVal = RegEnumValue(hKey, index, Name, nameLen, ByVal 0&, valueType, resBinary(0), 4096)
        End If
        If retVal Then Exit Do
        font = left$(Name, nameLen)
        font = Replace(font, "(TrueType)", "")
        font = Replace(font, "Bold", "")
        font = Replace(font, "Fett", "")
        font = Replace(font, "Italic", "")
        font = Replace(font, "Kursiv", "")
        font = Trim(font)
        If merk <> font Then
            merk = font
            re.AddNew
            re!SH_Wert = font
            re.Update
        End If
        index = index + 1
    Loop
    If handle Then RegCloseKey handle
    Set re = Nothing
    Set db = Nothing
    
End Function

