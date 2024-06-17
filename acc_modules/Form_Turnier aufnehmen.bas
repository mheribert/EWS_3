Option Compare Database
Option Explicit

Private Sub Befehl8_Click()

    DoCmd.Close
    
End Sub

Private Sub Befehl9_Click()
    
    DoCmd.Close
    
End Sub

Private Sub btnTurnierbericht_Click()
On Error GoTo Err_btnTurnierbericht_Click

    [Form_A-Programmübersicht]![Report_Turniernum] = Turnier_Nummer
    Dim stDocName As String
    stDocName = "Turnierbericht"
    DoCmd.OpenReport stDocName, acPreview

Exit_btnTurnierbericht_Click:
    Exit Sub

Err_btnTurnierbericht_Click:
    MsgBox err.Description
    Resume Exit_btnTurnierbericht_Click
    
End Sub

Private Sub Form_Close()
    Dim re As Recordset
    Dim vars
    Dim fehler As String
    Dim i, anzWR, anzMK As Integer
    Set re = Forms![Turnier aufnehmen]![Startklasse_Turnier Unterformular].Form.RecordsetClone
    If re.RecordCount <> 0 Then re.MoveFirst
    fehler = "Bei " & vbCrLf
    Do Until re.EOF
        anzWR = 0
        vars = Split(Nz(re!SelectWR), "+")
        For i = 0 To UBound(vars)
            anzWR = anzWR + vars(i)
        Next
        If anzWR <> re!AnzahlWR Then
            fehler = fehler & re!Startklasse_text & vbCrLf
        End If
        re.MoveNext
    Loop

    If Len(fehler) > 6 Then
        MsgBox fehler & "stimmt die Anzahl der Wertungsrichter nicht!" & vbCrLf & "Bitte neu eingeben!", vbInformation, "Achtung!"
    End If
    vars = Array("MK_11", "MK_12", "MK_13", "MK_21", "MK_22", "MK_23")
    anzMK = 0
    For i = 0 To 5
        If Nz(Me(vars(i))) <> "" Then
            anzMK = anzMK + 1
        End If
    Next
    If (anzMK < 4 Or (anzMK Mod 2) = 1) And InStr(Me!MehrkampfStationen, "Koordi") > 0 Then
        MsgBox "Die Anzahl der Mehrkampfstationen stimmt nicht", vbInformation, "Achtung!"
    End If
End Sub

Private Sub Form_Current()
    MehrkampfStationen_AfterUpdate
End Sub

Private Sub Form_Open(Cancel As Integer)
    If (Not IsNull([Form_A-Programmübersicht]![Akt_Turnier]) And [Form_A-Programmübersicht]![Akt_Turnier] <> 0 And [Form_A-Programmübersicht]![Akt_Turnier] <> "") Then
        Me.RecordsetClone.FindFirst "Turniernum=" & [Form_A-Programmübersicht]![Akt_Turnier]
        Me.Bookmark = Me.RecordsetClone.Bookmark
    End If
    Select Case Forms![A-Programmübersicht]!Turnierausw.Column(8)
        Case "SL"
            Me!MehrkampfStationen.Visible = False
        Case "BW"
            Me!MehrkampfStationen.Visible = False
        Case "BY"
            Me!MehrkampfStationen.Visible = False
        Case "HE"
            Me!MehrkampfStationen.Visible = False
            
        Case Else
    End Select
End Sub

Private Sub Form_Resize()
    If Me.InsideHeight > 7000 Then
        Me![Startklasse_Turnier Unterformular].Height = Me.InsideHeight - 6000
        Me![besondere_Vorkommnisse].Height = Me.InsideHeight - 6000
        Me.ScrollBars = 0
    Else
        Me.ScrollBars = 2
    End If
End Sub

Sub Kombinationsfeld35_AfterUpdate()
    ' Den mit dem Steuerelement übereinstimmenden Datensatz suchen.
    Me.RecordsetClone.FindFirst "Turniernum=" & Me![Kombinationsfeld35]
    Me.Bookmark = Me.RecordsetClone.Bookmark
End Sub

Private Sub MehrkampfStationen_AfterUpdate()
    mk_visible False
    Select Case Me!MehrkampfStationen
        Case "Bodenturnen und Trampolin"
            Me!Trampolin.Visible = True
            Me!Bodenturnen.Visible = True
            clr_MK_Felder
        Case "Kondition und Koordination"
            mk_visible True
        Case "Breitensportwettbewerb"
            Me!Kraft.Visible = True
            Me!Balance.Visible = True
            Me!Kondition.Visible = True
        Case Else
            clr_MK_Felder
    End Select
End Sub

Function clr_MK_Felder()
    Me!MK_11 = ""
    Me!MK_12 = ""
    Me!MK_13 = ""
    Me!MK_21 = ""
    Me!MK_22 = ""
    Me!MK_23 = ""
End Function

Function mk_visible(vi)
    Me!MK_11.Visible = vi
    Me!MK_12.Visible = vi
    Me!MK_13.Visible = vi
    Me!MK_21.Visible = vi
    Me!MK_22.Visible = vi
    Me!MK_23.Visible = vi
    Me!Bodenturnen.Visible = False
    Me!Trampolin.Visible = False
    Me!Kraft.Visible = False
    Me!Balance.Visible = False
    Me!Kondition.Visible = False
End Function

Function MK_test(fld)
    Dim flds, i
    flds = Array("MK_11", "MK_12", "MK_13", "MK_21", "MK_22", "MK_23")

    For i = 0 To 5
        If flds(i) <> "mk_" & fld Then
            If Me(flds(i)) = Me("mk_" & fld) Then
                MsgBox "Station ist schon vorhanden!"
            End If
        End If
    Next
End Function

Private Sub Mo_Name_Enter()
    Me!Mo_Name.Format = ""
End Sub

Private Sub Mo_Name_Exit(Cancel As Integer)
    Me!Mo_Name.Format = "@;""Vorname Name""[Blue]"
End Sub

Private Sub Tanzfläche_Enter()
    Me!Tanzfläche.Format = ""
End Sub

Private Sub Tanzfläche_Exit(Cancel As Integer)
    Me!Tanzfläche.Format = "@;""z.B. 6m x 6 m""[Blue]"
End Sub

Private Sub Belag_Enter()
    Me!Belag.Format = ""
End Sub

Private Sub Belag_Exit(Cancel As Integer)
    Me!Belag.Format = "@;""z.B. Parkett""[Blue]"
End Sub

