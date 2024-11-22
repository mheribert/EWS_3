Option Compare Database
Option Explicit
    
    Const msoFileDialogOpen = 1         ' Dialogfeld Öffnen
    Const msoFileDialogSaveAs = 2       ' Dialogfeld Speichern unter
    Const msoFileDialogFilePicker = 3   ' Dialogfeld Dateiauswahl
    Const msoFileDialogFolderPicker = 4 ' Dialogfeld Ordnerauswahl
    

Public Function get_Filename(fName As String, sfilter As String, pfad As String)
'    fname = get_Filename("Datenbanken", "*.mdb", "C:\")
'    fname = get_Filename("Ausschreibung", "*.pdf", getBaseDir)

    Dim f As Object
    Set f = Application.FileDialog(msoFileDialogFilePicker)
    f.Filters.Add fName, sfilter, 1
    f.InitialFileName = pfad
    
    If f.Show = -1 Then
        get_Filename = f.SelectedItems(1)
    End If

End Function

Public Function get_Folder(fName As String, sfilter, pfad) As String
    Dim f As Object
    Set f = Application.FileDialog(msoFileDialogFolderPicker)
    f.InitialFileName = pfad

    If f.Show = -1 Then
        get_Folder = f.SelectedItems(1)
    End If
End Function

Function FileSaveAs(fName, sfilter As String, pfad As String) As String
    fName = Replace(fName, "/", "_")
    fName = Replace(fName, "\", "_")
    fName = Replace(fName, ":", "_")
    fName = Replace(fName, "*", "_")
    fName = Replace(fName, "?", "_")
    fName = Replace(fName, "<", "_")
    fName = Replace(fName, ">", "_")
    fName = Replace(fName, """", "_")

    Dim f As Object
    Set f = Application.FileDialog(msoFileDialogSaveAs)
    f.InitialFileName = pfad & fName
    If f.Show = -1 Then
        f.Execute
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

Public Function verst(pfad)
    Dim fs As Object
    Dim a As Object
    Dim fName As String
    fName = pfad & "obs.html"
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(fName) Then fs.deletefile fName
    Set a = fs.CreateTextFile(fName, True)
    a.WriteLine ("")
    a.Close
    Set a = fs.GetFile(fName)
    a.Attributes = 2
End Function

Public Sub Bilderspeichern(EWS)
    Dim db As Database
    Dim ht As Recordset
    Dim out
    Dim tr_nr As String
    Dim base As String
    Dim t As Integer
    
    Set db = CurrentDb

    If EWS = "EWS3" Then
        base = getBaseDir & "webserver\views\"
        Set ht = db.OpenRecordset("Select * FROM HTML_Block WHERE Seite = 'All' and F1 = 'favicon.ico';")
        ht.MoveFirst
        out_bild ht, base
    Else
        base = gen_Ordner(getBaseDir & "Apache2")
        gen_Ordner (base & "\" & "cgi-bin")
        base = gen_Ordner(base & "\" & "htdocs") & "\"
'        getBaseDir & "Apache2\htdocs\"
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
            out.WriteLine Replace(ht!F1, "x__zahl", t)
        Next t
    End If

End Sub

Function out_bild(ht, base)
    Dim BilddateiID As Long
    Dim Dateigroesse As Long
    Dim Buffer() As Byte
On Error Resume Next
    
    BilddateiID = FreeFile
    Dateigroesse = Nz(LenB(ht!F3), 0)

    ReDim Buffer(Dateigroesse)
    Open base & Trim(ht!F1) For Binary Access Write As BilddateiID
    Buffer = ht!F3.GetChunk(0, Dateigroesse)
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
