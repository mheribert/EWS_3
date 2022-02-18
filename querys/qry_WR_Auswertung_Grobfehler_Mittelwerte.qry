SELECT Abgegebene_Wertungen.Paar_ID, Rundentab.RT_ID, Startklasse_Wertungsrichter.WR_function, Avg(Abgegebene_Wertungen.Grobfehler_Summe) AS MittelwertvonGrobfehler_Summe, Count(Abgegebene_Wertungen.Wertungsrichter_ID) AS AnzahlvonWertungsrichter_ID
FROM Startklasse_Wertungsrichter INNER JOIN (Abgegebene_Wertungen INNER JOIN Rundentab ON Abgegebene_Wertungen.RundenTab_ID = Rundentab.RT_ID) ON (Startklasse_Wertungsrichter.WR_ID = Abgegebene_Wertungen.Wertungsrichter_ID) AND (Startklasse_Wertungsrichter.Startklasse = Rundentab.Startklasse)
GROUP BY Abgegebene_Wertungen.Paar_ID, Rundentab.RT_ID, Startklasse_Wertungsrichter.WR_function
ORDER BY Abgegebene_Wertungen.Paar_ID, Rundentab.RT_ID, Startklasse_Wertungsrichter.WR_function;

