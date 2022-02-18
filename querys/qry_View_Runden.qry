SELECT Rundentab.RT_ID, Rundentab.Startklasse, Rundentab.Runde, Rundentab.Turniernr, Rundentab.Rundenreihenfolge, Rundentab.Startzeit, Rundentab.Anz_Paare, Rundentab.getanzt, Tanz_Runden.Rundentext, Tanz_Runden.R_NAME_ABLAUF, Startklasse.Startklasse_text, Startklasse.isStartklasse, [R_NAME_ABLAUF] & " " & [Startklasse_Text] AS Name, [Rundentext] & " " & [Startklasse_Text] AS Name_Ablauf
FROM Startklasse INNER JOIN (Tanz_Runden INNER JOIN Rundentab ON Tanz_Runden.Runde=Rundentab.Runde) ON Startklasse.Startklasse=Rundentab.Startklasse;

