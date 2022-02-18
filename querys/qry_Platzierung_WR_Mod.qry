SELECT View_Paare.Turniernr, Rundentab.Rundenreihenfolge, View_Paare.Startkl, View_Paare.Startnr, [WR_Vorname] & " " & [WR_Nachname] AS VollerName, View_Paare.Startklasse_text, "Wertungsrichter: " & [WR_Kuerzel] AS WRKurz, Auswertung.Platz AS Platz_WR, View_Paare.Platz, Turnier.Turnier_Name, Turnier.T_Datum, Turnier.Veranst_Clubnr, Turnier.Veranst_Name, Turnier.Veranst_Ort, Tanz_Runden.Rundentext, View_Paare.Name, View_Paare.Verein_Name, View_Majoritaet.disqualifiziert, Tanz_Runden.R_NAME_ABLAUF, View_Majoritaet.Majoritaet
FROM Wert_Richter
INNER JOIN ((Tanz_Runden
INNER JOIN (((((Turnier
INNER JOIN View_Paare ON Turnier.Turniernum = View_Paare.Turniernr)
INNER JOIN Paare_Rundenqualifikation ON View_Paare.TP_ID = Paare_Rundenqualifikation.TP_ID)
INNER JOIN Rundentab ON (Rundentab.RT_ID = Paare_Rundenqualifikation.RT_ID) AND (View_Paare.RT_ID_Ausgeschieden = Rundentab.RT_ID))
INNER JOIN Startklasse_Wertungsrichter ON View_Paare.Startkl = Startklasse_Wertungsrichter.Startklasse)
INNER JOIN View_Majoritaet ON (Paare_Rundenqualifikation.RT_ID = View_Majoritaet.RT_ID) AND (Paare_Rundenqualifikation.TP_ID = View_Majoritaet.TP_ID)) ON Tanz_Runden.Runde = Rundentab.Runde)
INNER JOIN Auswertung ON Paare_Rundenqualifikation.PR_ID = Auswertung.PR_ID) ON (Wert_Richter.WR_ID = Startklasse_Wertungsrichter.WR_ID) AND (Wert_Richter.WR_ID = Auswertung.WR_ID)
WHERE (((Rundentab.RT_ID)=[Formulare]![A-Programmübersicht]![Report_RT_ID]))
UNION SELECT View_Paare.Turniernr, Rundentab.Rundenreihenfolge, View_Paare.Startkl, View_Paare.Startnr, "Moderation" AS VollerName, View_Paare.Startklasse_text, "" AS WRKurz, MajoritaetKurz AS Platz_WR, View_Paare.Platz, Turnier.Turnier_Name, Turnier.T_Datum, Turnier.Veranst_Clubnr, Turnier.Veranst_Name, Turnier.Veranst_Ort, Tanz_Runden.Rundentext, View_Paare.Name, View_Paare.Verein_Name, View_Majoritaet.disqualifiziert, Tanz_Runden.R_NAME_ABLAUF, View_Majoritaet.Majoritaet

FROM Wert_Richter
INNER JOIN ((Tanz_Runden
INNER JOIN (((((Turnier
INNER JOIN View_Paare ON Turnier.Turniernum = View_Paare.Turniernr)
INNER JOIN Paare_Rundenqualifikation ON View_Paare.TP_ID = Paare_Rundenqualifikation.TP_ID)
INNER JOIN Rundentab ON (Rundentab.RT_ID = Paare_Rundenqualifikation.RT_ID) AND (View_Paare.RT_ID_Ausgeschieden = Rundentab.RT_ID))
INNER JOIN Startklasse_Wertungsrichter ON View_Paare.Startkl = Startklasse_Wertungsrichter.Startklasse)
INNER JOIN View_Majoritaet ON (Paare_Rundenqualifikation.RT_ID = View_Majoritaet.RT_ID) AND (Paare_Rundenqualifikation.TP_ID = View_Majoritaet.TP_ID)) ON Tanz_Runden.Runde = Rundentab.Runde)
INNER JOIN Auswertung ON Paare_Rundenqualifikation.PR_ID = Auswertung.PR_ID) ON (Wert_Richter.WR_ID = Startklasse_Wertungsrichter.WR_ID) AND (Wert_Richter.WR_ID = Auswertung.WR_ID)
WHERE (((Rundentab.RT_ID)=[Formulare]![A-Programmübersicht]![Report_RT_ID]))
and 

Wert_Richter.WR_ID in (
select min(wr.WR_ID) 
from Wert_Richter wr, 
Startklasse_Wertungsrichter ws, 
Rundentab rt 

where wr.wr_id=ws.wr_id 
and wr.turniernr = [Formulare]![A-Programmübersicht]![akt_turnier]
and ws.Startklasse=rt.Startklasse 
and rt.rt_id=[Formulare]![A-Programmübersicht]![Report_RT_ID])
ORDER BY 7;

