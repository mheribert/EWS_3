SELECT Auswertung.Wert_Ken AS Ausdr1, Wert_Richter.WR_Vorname, Wert_Richter.WR_Nachname
FROM Auswertung, Wert_Richter
WHERE (((Auswertung.Turniernr)=Formulare![A-Programmübersicht]!akt_Turnier) And ((Auswertung.T_Runde)=Formulare!Ausdrucke!Runde_einstellen) And ((Auswertung.Startkl)=Formulare!Ausdrucke![Startklasse einstellen]))
GROUP BY Auswertung.Wert_Ken, Wert_Richter.WR_Vorname, Wert_Richter.WR_Nachname;

