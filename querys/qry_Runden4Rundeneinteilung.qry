SELECT Runden.RT_ID, [Startklasse_Text] & " " & IIf([Getrennte_Auslosung]=True,[Rundentext],[R_NAME_ABLAUF]) AS R_NAME, Turnier.Getrennte_Auslosung, Runden.InRundeneinteilung, Runden.Runde, Turnier.Turniernum, Runden.Startklasse, Runden.Startklasse_text, Runden.Rundentext, Runden.Turnier_Name, Runden.InAuswertung, Runden.Reihenfolge, Runden.RF, Runden.Rundenreihenfolge
FROM Runden INNER JOIN Turnier ON Runden.Turniernum = Turnier.Turniernum
WHERE (((Turnier.Getrennte_Auslosung)=True) AND ((Runden.InRundeneinteilung)=0) AND ((Runden.Runde) Like "*_Fuß") AND ((Turnier.Turniernum)=[Formulare]![A-Programmübersicht]![akt_Turnier]) AND ((([Runden].[Runde])="End_r_schnell" Or ([Runden].[Runde])="End_r_lang")=False)) OR (((Runden.InRundeneinteilung)=1) AND ((Turnier.Turniernum)=[Formulare]![A-Programmübersicht]![akt_Turnier]))
ORDER BY Runden.Startklasse_text, Runden.RF, Runden.Rundenreihenfolge;

