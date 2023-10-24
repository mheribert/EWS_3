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
    If Nz(Forms!reisekostenabrechnung!Stunden8Tage) > 0 Or Nz(Forms!reisekostenabrechnung!Stunden24Tage) > 0 Then
        Me!Stunden8 = Forms!reisekostenabrechnung!Stunden8Tage
        Me!Stunden8sum = Me!Stunden8 * 14
        Me!Stunden24 = Forms!reisekostenabrechnung!Stunden24Tage
        Me!Stunden24sum = Me!Stunden24 * 28
        Me!Frühstück = Forms!reisekostenabrechnung!Frühstück
        Me!Frühstück_sum = Me!Frühstück * 5.6
        Me!Mittagessen = Forms!reisekostenabrechnung!Mittagessen
        Me!Mittagessen_sum = Me!Mittagessen * 11.2
        Me!Abendessen = Forms!reisekostenabrechnung!Abendessen
        Me!Abendessen_sum = Me!Abendessen * 11.2
        Me!Tagegeld = Nz(Me!Stunden8sum) + Nz(Me!Stunden24sum) - Nz(Me!Frühstück_sum) - Nz(Me!Mittagessen_sum) - Nz(Me!Abendessen_sum)
    End If
    If Nz(Forms!reisekostenabrechnung!ÜKosten) > 0 Then
        Me!ÜKostentext = Forms!reisekostenabrechnung!Ü_Text
        Me!ÜKosten = Forms!reisekostenabrechnung!ÜKosten
    End If
    If Nz(Forms!reisekostenabrechnung!Honorar_s) > 0 Then
        Me!Honorar_t = Forms!reisekostenabrechnung!Honorar_t
        Me!Honorar_s = Forms!reisekostenabrechnung!Honorar_s
    End If
    sum = Nz(Me!erg_km) + Nz(Me!erg_bahn) + Nz(Me!Tagegeld) + Nz(Me!ÜKosten) + Nz(Honorar_s)
    If sum > 0 Then
        Me!end_Bet = sum
        Me!gef_Bet = sum
    End If
raus:
End Sub

