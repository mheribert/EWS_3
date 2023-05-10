Option Compare Database

Public Function bindTables()
    
    ' Alle verknüpften Tabellen dynamisch neu einbinden
    On Error GoTo MyError
    
    Dim i As Integer
    Dim dbName As String
    Dim dirName As String
    Dim dirTLP As String
    
    dirName = getBaseDir()
       
    
    dirTLP = dirName & "Turnierleiterpaket\"
    ' Setzen des ICONs relativ zum Pfad

    CurrentDb.Properties("AppIcon") = dirName & "DRBV.ico"
    Application.RefreshTitleBar
    
    Application.SetOption ("Auto Compact"), 1
MyExit:
      Exit Function
    
MyError:
      MsgBox "Beim Starten der Applikation ist ein Fehler aufgetreten. ", 16, "Fehler"
      Resume MyExit
    
End Function

Public Sub bind_exttbl(mdb_Nr)
On Error Resume Next
    Dim db As DAO.Database
    Dim dbName As String
    Dim strDaten As String
    Dim strsql As String
    Dim i As Integer

    strDaten = getBaseDir & "T" & mdb_Nr & "_TDaten.mdb"
    
'Die Nächten 3 Zeilen dienen dem Schutz vor doppelten Feldname WR_func  als Bugfixing
    Set db = DBEngine.Workspaces(0).OpenDatabase(strDaten)
    strsql = "ALTER TABLE Startklasse_Wertungsrichter DROP COLUMN WR_func;"
    db.Execute strsql
    
    Set db = CurrentDb()
    For i = 0 To db.TableDefs.Count - 1
        If left(db.TableDefs(i).Connect, 9) = ";DATABASE" And left(db.TableDefs(i).Name, 1) <> "~" Then
            dbName = Mid(db.TableDefs(i).Connect, 11)
            db.TableDefs(i).Connect = ";database=" & strDaten
            db.TableDefs(i).RefreshLink
        End If
    Next i
End Sub
    
Public Function bindExcel(newDirName As String, tableName As String, fileName As String)
    
    Dim db As DAO.Database
    Set db = CurrentDb()
    
    Dim excelFileURL As String
    excelFileURL = "Excel 5.0;HDR=YES;IMEX=2;DATABASE=" & newDirName & fileName
    
    db.TableDefs(tableName).Connect = excelFileURL
    db.TableDefs(tableName).RefreshLink

End Function

Public Sub groessenkomprimierung(maxMegaByte As Integer)
On Error GoTo Err_groessenkomprimierung

    If (FileLen(CurrentDb.Name) / 1024 / 1024) > maxMegaByte Then
        Application.SetOption ("Auto Compact"), 1
    Else
        Application.SetOption ("Auto Compact"), 0
    End If
    
Exit_groessenkomprimierung:
    Exit Sub
    
Err_groessenkomprimierung:
    MsgBox err.Number & " - " & err.Description
    Resume Exit_groessenkomprimierung
    
End Sub

Public Function getBaseDir()
    
    ' Alle Hyperlinks in Buttons mit absolutem Pfad neu setzen
    On Error GoTo MyError
    
    Dim db As DAO.Database
    Set db = CurrentDb()
    Dim dirName As String
    dirName = left(db.Name, Len(db.Name) - Len(Dir(db.Name)))
    
    getBaseDir = dirName
    
    Exit Function
MyError:
      MsgBox "BaseDir konnte nicht ermittelt werden.", 16, "Fehler"
    
End Function
