Option Compare Database

Private Sub Report_Load()

Dim X As String

'x = Get_Akroname(45, "Vor_r", 6)
End Sub

Public Function Get_Akroname(TP_ID, Runde, Akronummer)

    Dim db As Database
    Dim Paare As Recordset
    Dim Akrobatiken As Recordset
    Dim RundTxt, AkroText As String
    Set db = CurrentDb()
    
    Set Paare = db.OpenRecordset("select * from Paare where TP_ID = " & TP_ID, DB_OPEN_DYNASET)
    'Set paare = db.OpenRecordset("SELECT Paare.*, Paare.TP_ID FROM Paare WHERE (((Paare.TP_ID)=2));", DB_OPEN_DYNASET)

    RundTxt = "_" & ch_runde(Runde)
    
    AkroText = "Akro" & Akronummer & RundTxt
    
    Set Akrobatiken = db.OpenRecordset("SELECT Akrobatiken.Akrobatik, Akrobatiken.Langtext FROM Akrobatiken WHERE (((Akrobatiken.Akrobatik) Like '" & Paare(AkroText) & "'));")
    
    'MsgBox (Paare(AkroText) & " - " & Akrobatiken("Langtext"))
    
    If IsNull(Paare(AkroText)) Then
        Get_Akroname = " "
    Else
        Get_Akroname = Akrobatiken("Langtext")
    End If

End Function

