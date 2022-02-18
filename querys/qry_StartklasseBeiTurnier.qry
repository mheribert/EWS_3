SELECT Startklasse.Startklasse, Startklasse.Startklasse_text, Startklasse.Reihenfolge, Startklasse_Turnier.Turniernr, Startklasse.isStartklasse
FROM Startklasse INNER JOIN Startklasse_Turnier ON Startklasse.Startklasse = Startklasse_Turnier.Startklasse
UNION SELECT Startklasse.Startklasse, Startklasse.Startklasse_text, Startklasse.Reihenfolge, Turnier.Turniernum, Startklasse.isStartklasse
FROM Startklasse, Turnier where Startklasse.isStartklasse=false
ORDER BY Startklasse.Reihenfolge;

