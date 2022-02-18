SELECT Auswertung.Startnr AS Ausdr1, Auswertung.Startkl AS Ausdr2, Auswertung.T_Runde AS Ausdr3, Auswertung.Turniernr AS Ausdr4, Sum(Auswertung.Platz) AS [Summe von Platz], Sum(Auswertung.Punkte) AS [Summe von Punkte], Turnier.Turnier_Name, Turnier.T_Datum, Paare.Da_Vorname, Paare.Da_NAchname, Paare.He_Vorname, Paare.He_Nachname, Paare.Verein_Name, Paare.Name_Team, Turnier.Turnier_Name, Turnier.Veranst_Name, Tanz_Runden.Rundentext
FROM Auswertung, Tanz_Runden, Turnier INNER JOIN Paare ON Turnier.Turniernum=Paare.Turniernr
GROUP BY Auswertung.Startnr, Auswertung.Startkl, Auswertung.T_Runde, Auswertung.Turniernr, Turnier.Turnier_Name, Turnier.T_Datum, Paare.Da_Vorname, Paare.Da_NAchname, Paare.He_Vorname, Paare.He_Nachname, Paare.Verein_Name, Paare.Name_Team, Turnier.Turnier_Name, Turnier.Veranst_Name, Tanz_Runden.Rundentext
ORDER BY Sum(Auswertung.Platz), Sum(Auswertung.Punkte) DESC;

