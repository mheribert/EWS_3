Option Compare Database

Public Sub writeErgebnisliste(fileName As String)

    Dim dbs As Database
    Set dbs = CurrentDb
    
    Dim rs As Recordset
    Set rs = dbs.OpenRecordset("Ergebnisliste_Text")
    
    Dim Akt_Turnier As Integer
    Akt_Turnier = [Forms]![A-Programmübersicht]![Akt_Turnier]
    
    If (Not rs.NoMatch) Then
        Dim fs, out, HTML
        Dim line As String
        Dim Turniername As String
        Dim Startklasse As String
        
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set out = fs.CreateTextFile(fileName, True)
        Set HTML = fs.CreateTextFile(Replace(fileName, ".txt", ".html"), True)
        
        line = "Ergebnisliste " & rs!Turnier_Name
        
        out.writeline (line)
        out.writeline (String(Len(line), "-"))
        out.writeline ("Version " & db_Ver())
        HTML.writeline ("Version " & db_Ver())
        
        HTML.writeline ("<p>&nbsp;")
        Startklasse = ""
        
        Do While (Not rs.EOF)
            ' Paar nur ausgeben, wenn es auch im aktuellen Turnier enthalten ist
            If (rs!Turniernr = Akt_Turnier) Then
                If (Startklasse <> rs!Startklasse_text) Then
                    Startklasse = rs!Startklasse_text
                    
                    out.writeline ("")
                    out.writeline (String(Len(Startklasse), "-"))
                    out.writeline (Startklasse)
                    out.writeline (String(Len(Startklasse), "-"))
                    
                    HTML.writeline ("</p>" & vbCrLf & "<h3><br />" & Startklasse & "</h3>" & vbCrLf & "<p>")

                End If
                If InStr(1, fileName, "Rang") > 0 Then
                    out.writeline (rs!Platz & ". " & rs!Name & "  " & rs!Verein_nr & " " & rs!Verein_Name & "  " & rs!Boogie_Startkarte_H & "  " & rs!Boogie_Startkarte_D)
                    HTML.writeline ("<br />" & rs!Platz & ". " & rs!Name & "  " & rs!Verein_nr & " " & rs!Verein_Name & "  " & rs!Boogie_Startkarte_H & "  " & rs!Boogie_Startkarte_D)
                Else
                    out.writeline (rs!Platz & ". " & rs!Name & " (" & rs!Verein_Name & ")")
                    HTML.writeline ("<br />" & rs!Platz & ". " & rs!Name & " (" & rs!Verein_Name & ")")
                End If
            End If
            rs.MoveNext
        Loop
        out.Close
        HTML.writeline ("</p>" & vbCrLf & "<p>&nbsp;</p>")
        HTML.Close
    End If
    
End Sub

Function print_wait_close(rpt, mo, Optional fi)
    DoCmd.OpenReport rpt, mo, , fi
    Do While SysCmd(acSysCmdGetObjectState, acReport, rpt) = 1
        DoEvents
    Loop
End Function

