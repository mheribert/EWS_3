SELECT IIf(Nz([name_team])="",[He_Nachname] & "-" & [Da_Nachname],[name_team]) AS Musik_Name, Paare.Startkl, Paare.Musik_FT AS Lieder, "Fuﬂtechnik" AS Pfad
FROM Paare
UNION
SELECT IIf(Nz([name_team])="",[He_Nachname] & "-" & [Da_Nachname],[name_team]) AS Musik_Name, Startkl, Musik_Akro, "Akrobatik" AS Pfad FROM Paare
UNION
SELECT IIf(Nz([name_team])="",[He_Nachname] & "-" & [Da_Nachname],[name_team]) AS Musik_Name, Startkl, Musik_Stell, "Stellprobe" AS Pfad FROM Paare
UNION
SELECT IIf(Nz([name_team])="",[He_Nachname] & "-" & [Da_Nachname],[name_team]) AS Musik_Name, Startkl, Musik_Form, "Formation" AS Pfad FROM Paare
UNION SELECT IIf(Nz([name_team])="",[He_Nachname] & "-" & [Da_Nachname],[name_team]) AS Musik_Name, Startkl, Musik_Sieg, "Siegerehrung" AS Pfad FROM Paare;

