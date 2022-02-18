SELECT Paare.Verein_Name, Max(Paare.Punkte) AS MaxvonPunkte, Paare.Startkl, Max(IIf(Formulare!Ausdrucke!Runde_einstellen Like "*" & "zw_r",(IIf([startkl]="RR_S",[punkte]*2,IIf([startkl]="RR_J",[punkte]*1.6,IIf([startkl]="rr_c",[punkte]*1.33,IIf([startkl]="rr_b",[punkte],IIf([startkl]="rR_A",[punkte],IIf([startkl]="RR_D",0,[punkte]))))))),(IIf([startkl]="RR_S",[punkte]*2,IIf([startkl]="RR_J",[punkte]*1.6,IIf([startkl]="rr_c",[punkte]*1.33,IIf([startkl]="rr_b",[punkte]*0.5,IIf([startkl]="rR_A",[punkte]*0.5,IIf([startkl]="RR_D",0,[punkte]))))))))) AS Team_Wertung
FROM Paare
GROUP BY Paare.Verein_Name, Paare.Startkl, Paare.Turniernr
HAVING (((Paare.Turniernr)=Formulare![A-Programmübersicht]!akt_Turnier))
ORDER BY Paare.Turniernr, Paare.Verein_Name;

