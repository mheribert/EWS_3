SELECT Paare.Turniernr, Paare.Startnr, [Da_Vorname] & " " & [Da_Nachname] & " - " & [He_Vorname] & " " & [He_Nachname] AS Name, Paare.Verein_Name, Startbuch_Status.SBS_Bezeichnung, Turnier.Turnier_Name, Turnier.Turnier_Nummer, Startbuch_Status.SBS_ID, Paare.Startkl, Startklasse.Startklasse_text, Turnier.T_Datum, Turnier.Veranst_Name, Paare.Startbuch as Startbuch, Paare.TP_ID
FROM Startbuch_Status INNER JOIN (Turnier INNER JOIN (Startklasse INNER JOIN Paare ON Startklasse.Startklasse = Paare.Startkl) ON Turnier.Turniernum = Paare.Turniernr) ON Startbuch_Status.SBS_ID = Paare.SBS_ID
WHERE (Startbuch_Status.SBS_ID>0) and not Startklasse.isTeam and Startklasse.Anzahl_Startbuecher=1
UNION
SELECT Paare.Turniernr, Paare.Startnr, Name_Team AS Name, Paare.Verein_Name, Startbuch_Status.SBS_Bezeichnung, Turnier.Turnier_Name, Turnier.Turnier_Nummer, Startbuch_Status.SBS_ID, Paare.Startkl, Startklasse.Startklasse_text, Turnier.T_Datum, Turnier.Veranst_Name, Paare.Startbuch, Paare.TP_ID
FROM Startbuch_Status INNER JOIN (Turnier INNER JOIN (Startklasse INNER JOIN Paare ON Startklasse.Startklasse = Paare.Startkl) ON Turnier.Turniernum = Paare.Turniernr) ON Startbuch_Status.SBS_ID = Paare.SBS_ID
WHERE (((Startbuch_Status.SBS_ID)>0) AND Startklasse.isTeam)
union
SELECT Paare.Turniernr, Paare.Startnr, [Da_Vorname] & " " & [Da_Nachname] AS Name, Paare.Verein_Name, Startbuch_Status.SBS_Bezeichnung, Turnier.Turnier_Name, Turnier.Turnier_Nummer, Startbuch_Status.SBS_ID, Paare.Startkl, Startklasse.Startklasse_text, Turnier.T_Datum, Turnier.Veranst_Name, Paare.Boogie_Startkarte_D, Paare.TP_ID
FROM Startbuch_Status INNER JOIN (Turnier INNER JOIN (Startklasse INNER JOIN Paare ON Startklasse.Startklasse = Paare.Startkl) ON Turnier.Turniernum = Paare.Turniernr) ON Startbuch_Status.SBS_ID = Paare.SBS_ID_BW_D
WHERE (Startbuch_Status.SBS_ID>0) and not Startklasse.isTeam and Startklasse.Anzahl_Startbuecher=2
UNION SELECT Paare.Turniernr, Paare.Startnr, [He_Vorname] & " " & [He_Nachname] AS Name, Paare.Verein_Name, Startbuch_Status.SBS_Bezeichnung, Turnier.Turnier_Name, Turnier.Turnier_Nummer, Startbuch_Status.SBS_ID, Paare.Startkl, Startklasse.Startklasse_text, Turnier.T_Datum, Turnier.Veranst_Name, Paare.Boogie_Startkarte_H, Paare.TP_ID
FROM Startbuch_Status INNER JOIN (Turnier INNER JOIN (Startklasse INNER JOIN Paare ON Startklasse.Startklasse = Paare.Startkl) ON Turnier.Turniernum = Paare.Turniernr) ON Startbuch_Status.SBS_ID = Paare.SBS_ID_BW_H
WHERE (Startbuch_Status.SBS_ID>0) and not Startklasse.isTeam and Startklasse.Anzahl_Startbuecher=2;

