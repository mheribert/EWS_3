Option Compare Database
Option Explicit

Private Sub Detailbereich_Format(Cancel As Integer, FormatCount As Integer)
On Error Resume Next
    ' grün hell rgb(230, 255, 230)
    ' grün = rgb(183, 255, 183)
    ' rot dunkler rgb(239, 195, 160)
    
    Dim ctrl
    Dim fld
    Dim i As Integer
    fld = Array("Herr_Grundtechnik_wert", "Herr_Haltung_Drehtechnik_wert", "Dame_Grundtechnik_wert", "Choreographie_wert", "Tanzfiguren_wert", "Tänzerische_Darbietung_wert")

    For i = 0 To UBound(fld)
        Set ctrl = Me(fld(i))

        If ((ctrl > 2) Or (ctrl < -2)) Then
            'rot
            ctrl.BackColor = rgb(237, 28, 36)
            ctrl.ForeColor = rgb(255, 255, 255)
            ctrl.FontBold = True
        Else
        
            If ((ctrl > 1.5) Or (ctrl < -1.5)) Then
                ' hell rot
                ctrl.BackColor = rgb(239, 195, 160)
                ctrl.ForeColor = rgb(0, 0, 0)
                ctrl.FontBold = False
            Else
                If ((ctrl > 1) Or (ctrl < -1)) Then
                    'gelb
                    ctrl.BackColor = rgb(255, 255, 171)
                    ctrl.ForeColor = rgb(0, 0, 0)
                    ctrl.FontBold = False
                Else
                    If (IsNull(ctrl)) Then
                        ' Normal Weiss
                        ctrl.BackColor = rgb(255, 255, 255)
                        ctrl.ForeColor = rgb(0, 0, 0)
                        ctrl.FontBold = False
                     Else
                        ' Normal Grün
                        ctrl.BackColor = rgb(255, 255, 255)
                        ctrl.ForeColor = rgb(0, 0, 0)
                        ctrl.FontBold = False
                     
                     End If
                End If
            End If
        End If
    Next
        

    If ((PunkteDiff > 20) Or (PunkteDiff < -20)) Then
            'rot
            PunkteDiff.BackColor = rgb(237, 28, 36)
            PunkteDiff.ForeColor = rgb(255, 255, 255)
            PunkteDiff.FontBold = True
    Else
    
        If ((PunkteDiff > 10) Or (PunkteDiff < -10)) Then
            ' hell rot
            PunkteDiff.BackColor = rgb(239, 195, 160)
            PunkteDiff.ForeColor = rgb(0, 0, 0)
            PunkteDiff.FontBold = False
        Else
            If ((PunkteDiff > 5) Or (PunkteDiff < -5)) Then
                'gelb
                PunkteDiff.BackColor = rgb(255, 255, 171)
                PunkteDiff.ForeColor = rgb(0, 0, 0)
                PunkteDiff.FontBold = False
            Else
                If (IsNull(PunkteDiff)) Then
                    ' Normal Weiss
                    PunkteDiff.BackColor = rgb(255, 255, 255)
                    PunkteDiff.ForeColor = rgb(0, 0, 0)
                    PunkteDiff.FontBold = False
                 Else
                    ' Normal Grün
                    PunkteDiff.BackColor = rgb(255, 255, 255)
                    PunkteDiff.ForeColor = rgb(0, 0, 0)
                    PunkteDiff.FontBold = False
                 
                 End If
            End If
                
        End If
    
    End If
        
        
    If ((Punkte - PunkteDurchSchnitt > 20) Or (Punkte - PunkteDurchSchnitt < -20)) Then
            'rot
            InAndOut.BackColor = rgb(237, 28, 36)
            InAndOut.ForeColor = rgb(255, 255, 255)
            InAndOut.FontBold = True
    Else
    
        If ((Punkte - PunkteDurchSchnitt > 10) Or (Punkte - PunkteDurchSchnitt < -10)) Then
            ' hell rot
            InAndOut.BackColor = rgb(239, 195, 160)
            InAndOut.ForeColor = rgb(0, 0, 0)
            InAndOut.FontBold = False
        Else
            If ((Punkte - PunkteDurchSchnitt > 5) Or (Punkte - PunkteDurchSchnitt < -5)) Then
                'gelb
                InAndOut.BackColor = rgb(255, 255, 171)
                InAndOut.ForeColor = rgb(0, 0, 0)
                InAndOut.FontBold = False
            Else
                If (IsNull(Punkte - PunkteDurchSchnitt)) Then
                    ' Normal Weiss
                    InAndOut.BackColor = rgb(255, 255, 255)
                    InAndOut.ForeColor = rgb(0, 0, 0)
                    InAndOut.FontBold = False
                 Else
                    ' Normal Grün
                    InAndOut.BackColor = rgb(255, 255, 255)
                    InAndOut.ForeColor = rgb(0, 0, 0)
                    InAndOut.FontBold = False
                 
                 End If
            End If
                
        End If
    
    End If

End Sub

Private Sub Report_Activate()
    If Nz(Forms![Ausdrucke]![Runde_einstellen]) = "" Then
        MsgBox ("Bitte Runde auswählen")
        DoCmd.Close acReport, Me.Name
    End If
End Sub

Private Sub Report_Load()
    Dim fil As String
    If InStr(1, Forms![Ausdrucke]![Runde_einstellen], "schnell") > 0 Then
        'Left([runde],4)
        fil = "Runde LIKE '" & left(Forms![Ausdrucke]![Runde_einstellen], 4) & "*' AND Startklasse = '" & Forms![Ausdrucke]!Startklasse_einstellen & "'"
    Else
        fil = "Runde = '" & Forms![Ausdrucke]![Runde_einstellen] & "' AND Startklasse = '" & Forms![Ausdrucke]!Startklasse_einstellen & "'"
    End If
    Me.Filter = fil
    Me.FilterOn = True

End Sub
