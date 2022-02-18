SELECT Runden.RT_ID, [R_NAME_ABLAUF] & " " & [Startklasse_Text] AS R_NAME, Runden.InRundeneinteilung, Turnier.Turniernum, Runden.Startklasse, Runden.Startklasse_text, Runden.Rundentext, Runden.Turnier_Name, Runden.Runde, Runden.InAuswertung
FROM Runden INNER JOIN Turnier ON Runden.Turniernum=Turnier.Turniernum
WHERE (((Runden.InRundeneinteilung)=1) And ((Turnier.Turniernum)=Formulare![A-Programmübersicht]!akt_Turnier))
ORDER BY Runden.Reihenfolge, Runden.Rundenreihenfolge;

