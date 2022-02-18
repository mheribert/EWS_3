PARAMETERS Zufallszahl IEEEDouble;
SELECT Paare.Zufallszahl AS Paarung, Paare.Startnr, Paare.Turniernr, Turnier.Turnier_Name, Turnier.T_Datum, Turnier.Veranst_Name, Paare.Startkl, Paare.Da_Vorname, Paare.Da_NAchname, Paare.He_Vorname, Paare.He_Nachname, Paare.Verein_nr, Paare.Verein_Name, Paare.Name_Team, Paare.Startbuch, Startklasse.Startklasse_text
FROM Turnier INNER JOIN (Startklasse INNER JOIN Paare ON Startklasse.Startklasse=Paare.Startkl) ON Turnier.Turniernum=Paare.Turniernr
WHERE (((Paare.Turniernr)=Formulare![A-Programmübersicht]!Akt_Turnier) And ((Paare.Startkl)="RR_A"))
ORDER BY Paare.Zufallszahl;

