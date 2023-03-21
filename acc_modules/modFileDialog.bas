Option Compare Database
Option Explicit

    #If Win64 And VBA7 Then
        Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
        Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
        Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
        Declare PtrSafe Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
    
        Private Type BROWSEINFO
             hOwner As LongPtr
             pidlRoot As Long
             pszDisplayName As String
             lpszTitle As String
             ulFlags As Long
             lpFn As LongPtr
             lParam As LongPtr
             iImage As Long
         End Type
        Public Type OPENFILENAME
            lStructSize As Long
            hwndOwner As LongPtr
            hInstance As LongPtr
            lpstrFilter As String
            lpstrCustomFilter As String
            nMaxCustFilter As Long
            nFilterIndex As Long
            lpstrFile As String
            nMaxFile As Long
            lpstrFileTitle As String
            nMaxFileTitle As Long
            lpstrInitialDir As String
            lpstrTitle As String
            flags As Long
            nFileOffset As Integer
            nFileExtension As Integer
            lpstrDefExt As String
            lCustData As LongPtr
            lpfnHook As LongPtr
            lpTemplateName As String
        End Type
    #Else
        Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
        Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
        Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
        Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
        
        Public Type BROWSEINFO
           hOwner As Long
           pidlRoot As Long
           pszDisplayName As String
           lpszTitle As String
           ulFlags As Long
           lpFn As Long
           lParam As Long
           iImage As Long
        End Type
        Public Type OPENFILENAME
            lStructSize As Long
            hwndOwner As Long
            hInstance As Long
            lpstrFilter As String
            lpstrCustomFilter As String
            nMaxCustFilter As Long
            nFilterIndex As Long
            lpstrFile As String
            nMaxFile As Long
            lpstrFileTitle As String
            nMaxFileTitle As Long
            lpstrInitialDir As String
            lpstrTitle As String
            flags As Long
            nFileOffset As Integer
            nFileExtension As Integer
            lpstrDefExt As String
            lCustData As Long
            lpfnHook As Long
            lpTemplateName As String
        End Type
    
    #End If

    Const OFN_READONLY           As Long = &H1
    Const OFN_EXPLORER           As Long = &H80000
    Const OFN_LONGNAMES          As Long = &H200000
    Const OFN_CREATEPROMPT       As Long = &H2000
    Const OFN_NODEREFERENCELINKS As Long = &H100000
    Const OFN_OVERWRITEPROMPT    As Long = &H2
    Const OFN_HIDEREADONLY       As Long = &H4
    Const OFS_FILE_OPEN_FLAGS    As Long = OFN_EXPLORER _
                                        Or OFN_LONGNAMES _
                                        Or OFN_CREATEPROMPT _
                                        Or OFN_NODEREFERENCELINKS
    Const OFS_FILE_SAVE_FLAGS    As Long = OFN_EXPLORER _
                                        Or OFN_LONGNAMES _
                                        Or OFN_OVERWRITEPROMPT _
                                        Or OFN_HIDEREADONLY
    Const BIF_RETURNONLYFSDIRS = &H1
    Public OpenFile As OPENFILENAME

Public Function GetFolder(szDialogTitle As String, XHwnd As Long) As String
  Dim retl As Long, bi As BROWSEINFO, dwIList As Long
  Dim szPath As String, wPos As Integer
  
    With bi
        .hOwner = XHwnd
        .lpszTitle = szDialogTitle
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    
    dwIList = SHBrowseForFolder(bi)
    szPath = Space$(512)
    retl = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
    
    If retl Then
        wPos = InStr(szPath, Chr(0))
        GetFolder = left$(szPath, wPos - 1)
    Else
        GetFolder = ""
    End If
End Function

Function FileSaveAs(sModul, sType, sFilters As String) As String
    sModul = Replace(sModul, "/", "_")
    sModul = Replace(sModul, "\", "_")
    sModul = Replace(sModul, ":", "_")
    sModul = Replace(sModul, "*", "_")
    sModul = Replace(sModul, "?", "_")
    sModul = Replace(sModul, "<", "_")
    sModul = Replace(sModul, ">", "_")
    sModul = Replace(sModul, """", "_")
    
    Dim intError As Integer
    Dim OFName As OPENFILENAME
    ' Formattyp-Filter festlegen
    With OFName
        'Setzt die Größe der OPENFILENAME Struktur
        .lStructSize = Len(OFName)
        'Der Window Handle ist bei VBA fast immer &O0
        .hwndOwner = &O0
        ' Formattyp-Filter setzen
        .lpstrFilter = sFilters
        ' Auswerten des Dateityps zur Auswahl des Filers
        Select Case sType
        Case ".txt"
        .nFilterIndex = 1
        Case Else
        .nFilterIndex = 2
        End Select
        ' Buffer für Dateinamen erzeugen
        .lpstrFile = sModul & Space$(1024) & vbNullChar & vbNullChar
        ' Maximale Anzahl der Dateinamen-Zeichen
        .nMaxFile = Len(.lpstrFile)
        ' Buffer für Titel erzeugen
        .lpstrFileTitle = Space$(254)
        ' Maximale Anzahl der Titel-Zeichen
        .nMaxFileTitle = 255
        ' Anfangsverzeichnis vorgeben
        .lpstrInitialDir = "c:\temp"
        .lpstrDefExt = sType & vbNullChar & vbNullChar
        ' Titel des Dialogfester festlegen
        .lpstrTitle = "Speichern unter..."
        ' Flags zum Festlegen eines bestimmten Verhaltens,
        ' OFN_LONGNAMES = lange Dateinamen verwenden
        ' OFN_OVERWRITEPROMPT = Abfrage vorm Überschreiben
        .flags = OFN_LONGNAMES Or OFN_OVERWRITEPROMPT
    End With
    ' API aufrufen und evtl. Fehler abfangen
    intError = GetSaveFileName(OFName)
    If intError <> 0 Then
        FileSaveAs = left(OFName.lpstrFile, InStr(1, OFName.lpstrFile, Chr(0)) - 1)
    ElseIf intError = 0 Then
        ' Abbruch durch Benutzer oder Fehler
    End If
End Function

Public Function gen_Ordner(pfad)
    Dim fs
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Len(Dir(pfad, vbDirectory)) = 0 Then
        fs.createfolder (pfad)
    End If
    gen_Ordner = pfad

End Function

Public Sub Bilderspeichern()
    Dim db As Database
    Dim ht As Recordset
    Dim out
    Dim tr_nr As String
    Dim base As String
    Dim t As Integer
    
    Set db = CurrentDb

    If get_properties("EWS") = "EWS3" Then
        base = getBaseDir & "webserver\views"
        Set ht = db.OpenRecordset("Select * FROM HTML_Block WHERE Seite = 'All' and F1 = 'favicon.ico';")
        ht.MoveFirst
        out_bild ht, base
    Else
        base = getBaseDir & "Apache2\htdocs\"
        Set ht = db.OpenRecordset("Select * FROM HTML_Block WHERE Seite = 'All' and Bereich = 'Bild';")
        ht.MoveFirst
        gen_default base
        gen_Ordner base & "gifs"
        gen_Ordner base & "beamer"
        gen_default base & "beamer\"
        gen_Ordner base & "moderator"
        gen_default base & "moderator\"
        gen_Ordner base & "observer"
        gen_default base & "observer\"
    

        Do Until ht.EOF
            out_bild ht, base
            ht.MoveNext
        Loop
        
        tr_nr = "T" & Forms![A-Programmübersicht]!Turnier_Nummer
        Set ht = open_re("All", "Zahl")
        For t = 1 To 7
            Set out = file_handle(base & tr_nr & "_" & t & ".html")
            out.writeline Replace(ht!f1, "x__zahl", t)
        Next t
    End If

End Sub

Function out_bild(ht, base)
    Dim BilddateiID As Long
    Dim Dateigroesse As Long
    Dim Buffer() As Byte
On Error Resume Next
    
    BilddateiID = FreeFile
    Dateigroesse = Nz(LenB(ht!f3), 0)

    ReDim Buffer(Dateigroesse)
    Open base & Trim(ht!f1) For Binary Access Write As BilddateiID
    Buffer = ht!f3.GetChunk(0, Dateigroesse)
    Put BilddateiID, , Buffer
    Close BilddateiID

 End Function

Function sp_mk()
    Dim db As Database
    Dim re As Recordset
    Dim lngDateigroesse As Long
    Dim Buffer() As Byte
    Dim dateiID As Integer
    Dim mkPfad As String
    Dim mkFile As String
    Dim mkArt As String
    Dim i As Integer
    
    Set db = CurrentDb
    Set re = db.OpenRecordset("analyse")
    mkArt = get_mk()
    mkPfad = getBaseDir & "Turn und Athletik-WB\"
    If mkArt = "Bodenturnen und Trampolin" Then
        mkFile = Dir(mkPfad & "1*.xlsx")
    ElseIf mkArt = "Kondition und Koordination" Then
        mkFile = Dir(mkPfad & "2*.xlsx")
    End If
    
    Do While mkFile <> ""
        re.FindFirst "CGI_Input = '" & mkFile & "'"
        If re.NoMatch Then
            re.AddNew
        Else
            re.Edit
        End If
        re!Cgi_Input = mkFile
        re!zeit = Time
        dateiID = FreeFile
        Open mkPfad & mkFile For Binary Access Read Lock Read Write As dateiID
        lngDateigroesse = FileLen(mkPfad & mkFile)
        ReDim Buffer(lngDateigroesse)
        re!datei = Null
        Get dateiID, , Buffer
        Close dateiID
        re!datei.AppendChunk Buffer

        re.Update
        mkFile = Dir
    Loop

End Function
