SELECT DISTINCT Tanz_Runden.Runde, Tanz_Runden.R_NAME_ABLAUF AS Rundentext, Rundentab.Startklasse AS startkl, Rundentab.Turniernr, Tanz_Runden.Rundenreihenfolge, Tanz_Runden.InRundeneinteilung
FROM Tanz_Runden INNER JOIN Rundentab ON Tanz_Runden.Runde=Rundentab.Runde
WHERE (((Rundentab.Startklasse)=Formulare![Paare_in erste Runde nehmen]!Startklasse) And ((Rundentab.Turniernr)=Formulare![A-Programmübersicht]!akt_Turnier) And ((Tanz_Runden.InRundeneinteilung)=1 Or (Tanz_Runden.InRundeneinteilung)=2))
ORDER BY Tanz_Runden.Rundenreihenfolge;

