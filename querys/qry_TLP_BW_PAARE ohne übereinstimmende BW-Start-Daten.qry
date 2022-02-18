SELECT TLP_BW_PAARE.Buchnr, TLP_BW_PAARE.Anrede, TLP_BW_PAARE.Nachname, TLP_BW_PAARE.Vorname, TLP_BW_PAARE.Startm, TLP_BW_PAARE.Clubnr, TLP_BW_PAARE.Clubname_kurz, TLP_BW_PAARE.[Geb-Dat-geprüft], TLP_BW_PAARE.Geburtsjahr, TLP_BW_PAARE.LRRVERB
FROM TLP_BW_PAARE LEFT JOIN [BW-Start-Daten] ON TLP_BW_PAARE.[Nachname] = [BW-Start-Daten].[Nachname]
WHERE ((([BW-Start-Daten].Nachname) Is Null));

