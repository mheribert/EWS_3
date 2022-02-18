SELECT Abgegebene_Wertungen.RundenTab_ID, Abgegebene_Wertungen.rh, Abgegebene_Wertungen.Wertungsrichter_ID, Count(Abgegebene_Wertungen.Wertungsrichter_ID) AS AnzahlvonWertungsrichter_ID
FROM Abgegebene_Wertungen
GROUP BY Abgegebene_Wertungen.RundenTab_ID, Abgegebene_Wertungen.rh, Abgegebene_Wertungen.Wertungsrichter_ID
HAVING (((Abgegebene_Wertungen.rh)=1));

