SELECT IIf([Name_Team] Is Not Null,"  " & [Name_Team],[Da_Vorname] & " " & [Da_Nachname] & " - " & [He_Vorname] & " " & [He_Nachname]) AS VollerName, Turnier.Turnier_Name, Turnier.T_Datum, Turnier.Veranst_Name, View_Paare.Startklasse_text, View_Paare.Anwesent_Status, Turnier.Turniernum, Startklasse.Reihenfolge, View_Paare.Startkl, View_Paare.Platz, View_Paare.Da_Alterskontrolle, View_Paare.He_Alterskontrolle, View_Paare.Startnr, View_Paare.Verein_Name, View_Paare.Name_Team, View_Paare.Da_Vorname, View_Paare.Da_Nachname, View_Paare.He_Vorname, View_Paare.He_Nachname
FROM Startklasse INNER JOIN (Turnier INNER JOIN View_Paare ON Turnier.Turniernum = View_Paare.Turniernr) ON Startklasse.Startklasse = View_Paare.Startkl
WHERE (((View_Paare.Anwesent_Status)>0) AND ((Turnier.Turniernum)=[Formulare]![A-Programmübersicht]![akt_Turnier]))
ORDER BY Startklasse.Reihenfolge;

