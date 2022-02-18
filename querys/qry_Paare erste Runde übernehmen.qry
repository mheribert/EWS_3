SELECT Turnier.Turnier_Name, Paare.Anwesent_Status, Turnier.T_Datum, Paare.Da_Vorname, Paare.Da_NAchname, Paare.He_Vorname, Paare.He_Nachname, Paare.Verein_Name, Paare.Name_Team
FROM Turnier INNER JOIN Paare ON Turnier.Turniernum=Paare.Turniernr
GROUP BY Turnier.Turnier_Name, Paare.Anwesent_Status, Turnier.T_Datum, Paare.Da_Vorname, Paare.Da_NAchname, Paare.He_Vorname, Paare.He_Nachname, Paare.Verein_Name, Paare.Name_Team
HAVING (((Paare.Anwesent_Status)<>0))
ORDER BY Paare.Anwesent_Status;

