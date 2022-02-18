SELECT DISTINCTROW View_Paare.Turniernr, View_Paare.Startkl, View_Paare.Platz, View_Paare.Punkte, View_Paare.Name, View_Paare.Verein_Name, Startklasse.Reihenfolge, Turnier.Turnier_Name, Turnier.T_Datum, View_Paare.Da_Vorname, View_Paare.Da_Nachname, View_Paare.He_Vorname, View_Paare.He_Nachname, View_Paare.Name_Team, Turnier.Veranst_Name, View_Paare.Startklasse_text, View_Paare.Startnr, View_Paare.Anwesent_Status, View_Paare.RT_ID_Ausgeschieden, View_Paare.Runde_Report
FROM Startklasse INNER JOIN (Turnier INNER JOIN View_Paare ON Turnier.Turniernum=View_Paare.Turniernr) ON Startklasse.Startklasse=View_Paare.Startkl
WHERE (((View_Paare.Turniernr)=Formulare![A-Programmübersicht]!akt_Turnier) And ((View_Paare.Platz)<>0))
ORDER BY View_Paare.Platz, Startklasse.Reihenfolge;

