Option Compare Database

Private Sub Report_Load()
On Error GoTo raus
    Dim sum As Currency
    Me!VorZuname = Forms!reisekostenabrechnung!Funktionär
    Me!Funktion = Forms!reisekostenabrechnung!Funktion
    Me!Anschrift = Forms!reisekostenabrechnung!Anschrift
    Me!von_nach = Forms!reisekostenabrechnung!Reise_von & ("  /  " + Forms!reisekostenabrechnung!Reise_nach)
'    Me!grund_reise = Forms!reisekostenabrechnung!GrundReise
    Me!ReiseBD = Forms!reisekostenabrechnung!BeginnR
    Me!ReiseBU = Format(Forms!reisekostenabrechnung!UhrzeitBR, "hh:mm")
    Me!ReiseED = Forms!reisekostenabrechnung!EndeR
    Me!ReiseEU = Format(Forms!reisekostenabrechnung!UhrzeitER, "hh:mm")
    Me!DienstBD = Forms!reisekostenabrechnung!BeginnD
    Me!DienstBU = Format(Forms!reisekostenabrechnung!UhrzeitBD, "hh:mm")
    Me!DienstED = Forms!reisekostenabrechnung!EndeD
    
    Me!DienstEU = Format(Forms!reisekostenabrechnung!UhrzeitED, "hh:mm")
    If Nz(Forms!reisekostenabrechnung!PKW_km) > 0 Then
        Me!PKW_km = Nz(Forms!reisekostenabrechnung!PKW_km)
        Me!KM300 = IIf(Me!PKW_km > 300, 300, Me!PKW_km)
        Me!KM400 = IIf(Me!PKW_km > 300, Me!PKW_km - 300, Null)
        Me!bis_300 = IIf(Me!KM300 = "", "", Me!KM300 * 0.3)
        Me!ab_300 = IIf(Me!KM400 = Null, Null, Me!KM400 * 0.15)
        Me!alle_km = Nz(Me!bis_300) + Nz(Me!ab_300)
        Me!erg_km = Nz(Me!alle_km) * 2
    End If
    sum = Nz(Forms!reisekostenabrechnung!Bahn_Flug) + Nz(Forms!reisekostenabrechnung!Zuschläge) + Nz(Forms!reisekostenabrechnung!An_Abfahrt) + Nz(Forms!reisekostenabrechnung!anf_PKW) * 0.3
    If sum > 0 Then
        Me!Bahn_Flug = Forms!reisekostenabrechnung!Bahn_Flug
        Me!Zuschläge = Forms!reisekostenabrechnung!Zuschläge
        Me!An_Abfahrt = Forms!reisekostenabrechnung!An_Abfahrt
        Me!anf_PKW = Forms!reisekostenabrechnung!anf_PKW * 0.3
        Me!erg_bahn = sum
    End If
    If Nz(Forms!reisekostenabrechnung!Stunden8Tage) > 0 Or Nz(Forms!reisekostenabrechnung!Stunden14Tage) > 0 Or Nz(Forms!reisekostenabrechnung!Stunden24Tage) > 0 Then
        Me!Stunden8 = Forms!reisekostenabrechnung!Stunden8Tage
        Me!Stunden14 = Forms!reisekostenabrechnung!Stunden14Tage
        Me!Stunden24 = Forms!reisekostenabrechnung!Stunden24Tage
        Me!Frühstück_Tg = Forms!reisekostenabrechnung!Frühstück
        Me!Stunden8sum = Me!Stunden8 * 12
        Me!Stunden14sum = Me!Stunden14 * 12
        Me!Stunden24sum = Me!Stunden24 * 24
        Me!Frühstücksum = Me!Frühstück_Tg * 4.5
        Me!Tagegeld = Nz(Me!Stunden8sum) + Nz(Me!Stunden14sum) + Nz(Me!Stunden24sum) - Nz(Me!Frühstücksum)
    End If
    If Nz(Forms!reisekostenabrechnung!ÜKosten) > 0 Then
        Me!ÜKostentext = Forms!reisekostenabrechnung!Ü_Text
        Me!ÜKosten = Forms!reisekostenabrechnung!ÜKosten
    End If
    If (Me!VorZuname = "Heribert Mießlinger" And Not IsNull(Me!ÜKosten)) Then
        Me.Bezeichnungsfeld83.Caption = "sonstige Kosten"
        Me.Bezeichnungsfeld193.Visible = False
        Me!ÜKostentext = "Laptop,Drucker, Router, Papier und Kleinmaterial"
        Me.Bezeichnungsfeld88.Visible = False
    End If
    If (Me!VorZuname = "Christian Punk" And Not IsNull(Me!ÜKosten)) Then
        Me.Bezeichnungsfeld83.Caption = "sonstige Kosten"
        Me.Bezeichnungsfeld193.Visible = False
        Me!ÜKostentext = "Aufwandsentschädigung"
        Me.Bezeichnungsfeld88.Visible = False
    End If
    sum = Nz(Me!erg_km) + Nz(Me!erg_bahn) + Nz(Me!Tagegeld) + Nz(Me!ÜKosten)
    If sum > 0 Then
        Me!end_Bet = sum
        Me!gef_Bet = sum
    End If
raus:
End Sub

