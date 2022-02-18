SELECT Rundentab.RT_ID, [Startklasse_text] & " - " & [Rundentext] AS Rundenname, Turnier.Turniernum, Rundentab.Startklasse, Startklasse.Startklasse_text, Tanz_Runden.Rundentext, Turnier.Turnier_Name, Tanz_Runden.Runde, Tanz_Runden.R_IS_ENDRUNDE, Tanz_Runden.R_NAME_ABLAUF, Tanz_Runden.InRundeneinteilung, Tanz_Runden.InAuswertung, Tanz_Runden.InPunkteeingabe, Tanz_Runden.MitStartklasse, Tanz_Runden.R_IS_ENDRUNDE, Startklasse.Reihenfolge, Tanz_Runden.Rundenreihenfolge, Rundentab.Rundenreihenfolge AS RF
FROM Tanz_Runden INNER JOIN (Startklasse INNER JOIN (Turnier INNER JOIN Rundentab ON Turnier.Turniernum = Rundentab.Turniernr) ON Startklasse.Startklasse = Rundentab.Startklasse) ON Tanz_Runden.Runde = Rundentab.Runde
WHERE (((Turnier.Turniernum)=[Formulare]![A-Programmübersicht]![akt_Turnier]))
ORDER BY Startklasse.Reihenfolge, Tanz_Runden.Rundenreihenfolge;

