SELECT Rundentab.Turniernr, Rundentab.Rundenreihenfolge, Rundentab.Runde, Rundentab.Startklasse, Rundentab.Startzeit, Count(Paare_Rundenqualifikation.TP_ID) AS Anz_Paare, Rundentab.getanzt, Tanz_Runden.Rundentext, Startklasse.Startklasse_text
FROM (Tanz_Runden INNER JOIN (Startklasse INNER JOIN Rundentab ON Startklasse.Startklasse=Rundentab.Startklasse) ON Tanz_Runden.Runde=Rundentab.Runde) INNER JOIN Paare_Rundenqualifikation ON Rundentab.RT_ID=Paare_Rundenqualifikation.RT_ID
GROUP BY Rundentab.Turniernr, Rundentab.Rundenreihenfolge, Rundentab.Runde, Rundentab.Startklasse, Rundentab.Startzeit, Rundentab.getanzt, Tanz_Runden.Rundentext, Startklasse.Startklasse_text
HAVING (((Rundentab.Turniernr)=Formulare![A-Programmübersicht]!akt_Turnier))
ORDER BY Rundentab.Rundenreihenfolge;

