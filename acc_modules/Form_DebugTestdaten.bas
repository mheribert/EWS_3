Option Compare Database

Private Sub Befehl0_Click()
Dim DateinameEinlesen As String
Dim DateinameAusgabe As String
Dim pfad As String
Dim Turniernummer As Long
Dim RT_ID As Long
Dim AusgabeString As String
Dim oFile As Object
Dim sLines() As String
Dim oFSO As Object

Dim Zeile As String


If IsNull(Me.Zeile) Then Me.Zeile = 0

Turniernummer = [Forms]![A-Programmübersicht]![Turnier_Nummer]
RT_ID = Me.Tanzrunde

DateinameEinlesen = getBaseDir & "Testdaten\T" & Turniernummer & "_RT" & RT_ID & ".txt"
DateinameAusgabe = getBaseDir & "T" & Turniernummer & "_RT" & RT_ID & ".txt"

Open DateinameEinlesen For Input As #1

AusgabeString = ""

Do While Not EOF(1)

    Line Input #1, Zeile

    AusgabeString = AusgabeString & Zeile & vbNewLine
Loop
    
Close #1

    
Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    Set oFile = oFSO.OpenTextFile(DateinameEinlesen)
 
    ' Alles lesen und in Array zerlegen
    sLines = Split(oFile.ReadAll, vbCrLf)
 
    ' Datei schließen
    oFile.Close
    
If Me.Zeile <= UBound(sLines) Then
    Open DateinameAusgabe For Append As #2

    Print #2, sLines(Me.Zeile)
    Close #2

    Me.Zeile = Me.Zeile + 1
Else
    Me.Zeile = 0
    Me.Tanzrunde.SetFocus
    Me.Befehl0.Enabled = False
End If

End Sub

Private Sub Tanzrunde_AfterUpdate()
    Me.Zeile = 0
    Me.Befehl0.Enabled = True
End Sub
